# -*- coding: utf-8 -*-
"""
配置管理模块
============================================================

【模块职责】
    管好"程序运行需要的各种设置"以及"把设置保存到硬盘"，具体包括：
      1. 定义数据库连接配置（ConnectionConfig）
      2. 定义全局运行配置（AppConfig）：模板路径、输出目录、会计期间等
      3. 把配置保存到用户目录，实现"下次打开还记得上次的连接信息"
      4. 密码加密存储（避免别人直接打开配置文件就看到密码）

【新手须知：配置为什么要存到 %APPDATA% 而不是程序目录？】
    %APPDATA% 是 Windows 给当前用户存放个人数据的目录。
    把配置存在这里的好处：
      - 程序安装目录可能没有写权限（如 Program Files），写这里不会被拒绝；
      - 换个用户登录，用的就是各自的配置，互不干扰。
"""
from __future__ import annotations

import base64        # 标准的 Base64 编解码（Fernet 加密结果本身是字节，需要转成文本）
import json          # 把 Python 字典 <-> JSON 文本 互相转换
import os            # 读取环境变量（APPDATA）、修改文件权限
from dataclasses import asdict, dataclass, field   # 数据类相关工具
from pathlib import Path
from typing import Optional


# ===========================================================================
# 一、路径约定
# ===========================================================================
# Path(__file__) 表示"当前这个 .py 文件本身"。
# .resolve() 得到它的绝对路径；.parent 表示"上一层目录"。
# 连续两次 .parent：config/settings.py -> config 目录 -> 项目根目录
PROJECT_ROOT = Path(__file__).resolve().parent.parent

# 用户配置目录：%APPDATA%/YiYangReport
# os.environ.get("APPDATA", ...)：优先取环境变量 APPDATA 的值；
# 如果取不到（理论上非 Windows 才会发生），就用用户主目录兜底。
USER_CONFIG_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "YiYangReport"

# 如果目录不存在就创建（parents=True 表示父目录也一起创建；
# exist_ok=True 表示"已存在也不报错"）
USER_CONFIG_DIR.mkdir(parents=True, exist_ok=True)

# 配置文件与密钥文件的完整路径
CONFIG_FILE = USER_CONFIG_DIR / "user_config.json"   # 保存各项设置
KEY_FILE = USER_CONFIG_DIR / ".secret.key"           # 保存加密用密钥

# 默认模板文件名（只读，程序绝不修改它）
DEFAULT_TEMPLATE = "2026版新修订企业财务报模板.xls"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output"

# 常用报表（9 张）。
# 【设计意图】财务人员日常最常生成的就是这 9 张，
#   在"选择要生成的报表"窗口里用特殊颜色标注，
#   并提供"只看常用9表"按钮，避免在一长串表页里反复寻找、误选。
# 用元组（tuple）而不是列表，是因为这份清单是"固定不变的常量"，
#   元组不可修改，能防止被其它代码意外改动。
COMMON_REPORT_CODES = (
    "LK01", "LK02", "LK03",
    "LK25", "LK26", "LK27", "LK28",
    "LK31", "LK33",
)


# 首选 ODBC 驱动清单（越靠前越优先）。
# 【设计意图】本工具的使用者是普通财务人员，他们不懂"ODBC 驱动"是什么。
#   因此程序不要求用户选择驱动，而是按下面的顺序，
#   自动从"本机已安装的驱动"里挑一个能连通的来用。
PREFERRED_SQLSERVER_DRIVERS = [
    "ODBC Driver 18 for SQL Server",
    "ODBC Driver 17 for SQL Server",
    "ODBC Driver 13 for SQL Server",
    "ODBC Driver 11 for SQL Server",
    "SQL Server Native Client 11.0",
    "SQL Server Native Client 10.0",
    "SQL Server",
]


def list_sqlserver_drivers() -> list[str]:
    """
    返回本机已安装、可用于 SQL Server 的 ODBC 驱动名（按优先级从高到低）。

    【新手必学：pyodbc.drivers()】
        它会列出本机所有 ODBC 驱动，但里面也混着 Excel、Access、MySQL 等
        无关驱动，所以我们要按"已知的 SQL Server 驱动清单"过滤并排序。

    返回值：
        - 有可用驱动：返回按优先级排好的驱动名列表；
        - 连 pyodbc 都没装：返回空列表（由上层给出友好提示）。
    """
    try:
        import pyodbc
    except ImportError:
        return []

    installed = set(pyodbc.drivers())
    # 第一步：按预设优先级，挑出"已安装"的那些
    ordered = [d for d in PREFERRED_SQLSERVER_DRIVERS if d in installed]
    # 第二步：把名字里含 "SQL Server/SQLSERVER" 的其他驱动补充到末尾，
    #         以防将来出现清单里没列到的新驱动。
    for d in sorted(installed):
        if d in ordered:
            continue
        up = d.upper()
        if "SQL SERVER" in up or "SQLSERVER" in up:
            ordered.append(d)
    return ordered


# ===========================================================================
# 二、数据结构
# ===========================================================================
@dataclass
class ConnectionConfig:
    """
    SQL Server 连接配置。

    【新手必学：@dataclass 数据类】
        普通写法要自己写 __init__ 把参数一个个赋值给 self.属性，很啰嗦。
        加上 @dataclass 装饰器后，Python 会自动帮我们生成 __init__、__repr__ 等，
        我们只要像下面这样"声明字段和默认值"即可：

            @dataclass
            class ConnectionConfig:
                server: str = "localhost"

        等价于手写了 __init__(self, server="localhost") 并赋值。
        字段后面 `: str` 是"类型提示"，帮助阅读和理解，运行时不做强制检查。

    【注意】字段顺序很重要：有默认值的字段不能排在没默认值的字段前面。
    """

    server: str = "localhost"          # 服务器地址，如 192.168.1.10 或 localhost
    instance: str = ""                  # 实例名，如 SQLEXPRESS；默认实例留空
    port: int = 1433                    # 端口。1433 是 SQL Server 默认端口

    # 【重要：database 的含义已经变了】
    #   以前这里填的是"年度库"（如 UFDATA_001_2026），要求客户自己记住库名，
    #   一旦填错，程序就会安安静静地查错账套的数据 —— 这是最危险的情况。
    #
    #   现在这里填的是**登录时连的库**，固定为用友系统库（默认 UFSystem）。
    #   年度库不再由人填写，而是在生成时由"账套号 + 会计年度"自动解析出来
    #   （解析规则见 core/u8_account.py，数据来自 UFSystem.dbo.UA_AccountDatabase）。
    database: str = "UFSystem"

    # 用友系统库（账套注册库）的名字。
    # 绝大多数现场都叫 UFSystem，但个别改过库名的环境可以在这里调整，
    # 程序查账套表时会用这个库名做三段式限定。
    system_database: str = "UFSystem"

    # 运行时解析出来的年度库（如 UFDATA_001_2026）。
    # 【说明】它不是让用户填的，而是每次生成前根据账套号+年度查出来再写进来。
    #   存进配置只是为了日志能看出"上次用的是哪个库"，不参与任何判断。
    target_database: str = ""

    username: str = ""                  # SQL 登录名
    password: str = ""                  # SQL 登录密码（内存中为明文；存盘时会加密）
    use_windows_auth: bool = False      # 是否使用 Windows 集成认证
    # ODBC 驱动名。留空("")表示"自动检测"：
    #   程序会按优先级依次尝试本机已安装的 SQL Server 驱动；
    #   连接成功后会把实际可用的驱动名回填到这里，下次优先使用。
    driver: str = ""
    timeout: int = 30                   # 连接/查询超时（秒）
    trust_server_certificate: bool = True  # 是否信任服务器证书（内网常用 True）

    def build_conn_str(self) -> str:
        """
        把上面的字段"拼装"成 pyodbc 能识别的连接字符串。

        【连接字符串长什么样？】
            DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost,1433;
            DATABASE=UFDATA_001_2026;UID=sa;PWD=xxx;TrustServerCertificate=yes;
            它就是一串用分号分隔的"键=值"文本。
        """
        # 先决定 SERVER 部分的写法（三种情况互斥）：
        server_part = self.server
        if self.instance:
            # 有实例名：写成 服务器\实例名，例如 localhost\SQLEXPRESS
            server_part = f"{self.server}\\{self.instance}"
        elif self.port:
            # 没实例名但有端口：写成 服务器,端口，例如 192.168.1.10,1433
            server_part = f"{self.server},{self.port}"

        # 决定驱动名：若配置里留空，说明要"自动"，
        # 就先挑一个本机可用的驱动占位（真正的依次尝试在 Database.connect 里做）。
        driver = self.driver.strip()
        if not driver:
            candidates = list_sqlserver_drivers()
            driver = candidates[0] if candidates else "SQL Server"

        # 用一个列表收集每个"键=值"片段
        parts = [
            f"DRIVER={{{driver}}}",   # 驱动名两边要加大括号（驱动名本身含空格）
            f"SERVER={server_part}",
        ]
        # 【关键】连接时只进系统库，不进年度库。
        #   原因：账套与年度的对应关系必须先查系统库才知道，
        #   在"还不知道该连哪个年度库"的时候，直接连年度库是无从谈起的事。
        login_db = self.system_database or self.database
        if login_db:
            parts.append(f"DATABASE={login_db}")

        # 认证方式二选一
        if self.use_windows_auth:
            parts.append("Trusted_Connection=yes")
        else:
            parts.append(f"UID={self.username}")
            parts.append(f"PWD={self.password}")

        if self.trust_server_certificate:
            parts.append("TrustServerCertificate=yes")

        # ";".join(列表) 会把列表元素用分号连接成一个字符串，末尾再补一个分号
        return ";".join(parts) + ";"

    def to_safe_dict(self) -> dict:
        """
        转成字典，但把密码清空。
        用途：把配置回显到界面时，不要把密码显示/传递出去，更安全。
        """
        d = asdict(self)   # dataclass 自带工具：把对象转成普通字典
        d["password"] = ""
        return d


@dataclass
class AppConfig:
    """应用全局配置：把连接配置 + 其他运行设置打包在一起。"""

    # field(default_factory=...) 的含义：
    #   这里默认值是一个"新对象"。如果直接写 connection: ConnectionConfig = ConnectionConfig()，
    #   Python 会让所有实例共用同一个对象，可能互相干扰。
    #   用 default_factory=ConnectionConfig 表示"每次需要默认值时，都新建一个"。
    connection: ConnectionConfig = field(default_factory=ConnectionConfig)
    template_path: str = str(PROJECT_ROOT / DEFAULT_TEMPLATE)
    output_dir: str = str(DEFAULT_OUTPUT_DIR)

    # 会计期间（要生成哪年哪月的报表）
    fiscal_year: int = 2026
    fiscal_month: int = 5

    # 账套号（U8 年度库后缀中的账套号，如 001）
    account_set: str = "001"

    # 最近一次选中的报表代码列表（记住用户上次的选择）
    # 注意：列表这类"可变对象"同样要用 default_factory
    last_selected_reports: list = field(default_factory=list)


# ===========================================================================
# 三、密码加解密（基于 cryptography 库的 Fernet 算法）
# ===========================================================================
def _get_fernet():
    """
    获取 Fernet 加密器对象；第一次使用时自动生成密钥文件。

    【Fernet 是什么？】
        Fernet 是 cryptography 库提供的一种"对称加密"方案：
        用同一个密钥既能加密也能解密。它安全、简单、不容易用错，适合我们这种场景。

    【返回值】
        - 正常情况：返回一个 Fernet 对象
        - 如果没安装 cryptography 库：返回 None（程序会降级为"不加密"）
    """
    try:
        from cryptography.fernet import Fernet
    except ImportError:
        # 没装加密库时降级处理：仍然能运行，只是密码以明文保存
        return None

    # 密钥不存在则生成一个并存盘
    if not KEY_FILE.exists():
        KEY_FILE.write_bytes(Fernet.generate_key())   # 生成随机密钥并写入文件
        try:
            # 尽量把密钥文件权限设为"仅本人可读写"（0o600 是 Linux 风格写法）
            # Windows 上可能不支持，失败也不影响运行，所以用 try 兜住
            os.chmod(KEY_FILE, 0o600)
        except OSError:
            pass

    key = KEY_FILE.read_bytes()
    return Fernet(key)


def encrypt_password(plain: str) -> str:
    """
    加密密码。

    参数 plain：明文密码
    返回：加密后的字符串。这样即使有人打开配置文件，也看不懂密码。
    """
    if not plain:
        return ""

    f = _get_fernet()
    if f is None:
        # 没有加密库时的"降级方案"：加一个 PLAIN: 前缀标记，表示这是明文。
        # 解密时看到这个前缀就知道不用真解密。
        return "PLAIN:" + plain

    # f.encrypt 得到的是一段字节(bytes)，用 decode("ascii") 转成文本才能存进 JSON
    token = f.encrypt(plain.encode("utf-8"))
    return token.decode("ascii")


def decrypt_password(cipher: str) -> str:
    """解密密码，是 encrypt_password 的逆操作。"""
    if not cipher:
        return ""

    # 处理"降级方案"产生的明文前缀
    if cipher.startswith("PLAIN:"):
        return cipher[len("PLAIN:"):]

    f = _get_fernet()
    if f is None:
        return ""

    try:
        return f.decrypt(cipher.encode("ascii")).decode("utf-8")
    except Exception:
        # 例如换了电脑导致密钥不匹配，解密会失败，此时返回空字符串
        return ""


# ===========================================================================
# 四、读写配置文件
# ===========================================================================
def save_config(cfg: AppConfig) -> None:
    """
    保存配置到磁盘。

    【做法】
        1. 先把 AppConfig 对象转成字典；
        2. 把密码单独加密；
        3. 用 json.dumps 转成文本，写入文件。
    """
    data = asdict(cfg)

    # 密码单独处理：取出明文 → 加密 → 放回字典
    pwd = cfg.connection.password
    data["connection"]["password"] = encrypt_password(pwd)
    # 加个标记，方便读取时知道"这个 password 字段是加密过的"
    data["connection"]["_password_encrypted"] = True

    # ensure_ascii=False 让中文原样写入（不转成 \uXXXX 转义），便于人工查看
    # indent=2 让 JSON 有缩进，更易读
    CONFIG_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def load_config() -> AppConfig:
    """
    从磁盘读取配置；文件不存在或损坏时，返回一份默认配置。

    【新手须知：为什么要处理"文件不存在"？】
        程序第一次运行时，配置文件还不存在。
        如果直接读取就会抛出异常导致程序崩溃。
        所以这里先判断存在性，再动手读——这叫"防御式编程"。
    """
    if not CONFIG_FILE.exists():
        return AppConfig()

    # 读取并解析 JSON；若文件损坏（比如被手动改坏了），也要能优雅降级
    try:
        data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return AppConfig()

    # 取出连接配置部分，并按需解密密码
    conn_data = data.get("connection", {})
    if conn_data.get("_password_encrypted"):
        conn_data["password"] = decrypt_password(conn_data.get("password", ""))
    # 这个标记只是临时用的，删掉它，免得传进 ConnectionConfig 时报错
    conn_data.pop("_password_encrypted", None)

    # 【新手难点：字典过滤】
    #   ConnectionConfig.__dataclass_fields__ 会列出该类所有合法字段名。
    #   下面这行只保留"合法字段"，丢弃多余的键，防止旧配置里残留的字段导致报错。
    conn = ConnectionConfig(**{
        k: v for k, v in conn_data.items()
        if k in ConnectionConfig.__dataclass_fields__
    })
    # 上面的 {...} 是"字典推导式"，**conn 是把字典当作关键字参数传入构造函数

    # 【旧配置迁移】老版本的 database 里存的是年度库（UFDATA_xxx_yyyy）。
    #   如果直接沿用，程序就会"登录就进了某个年度库"，
    #   万一那个库已被删除或改名，连测试连接都会失败。
    #   所以这里自动纠正为系统库，让新逻辑从干净状态开始。
    if conn.database.upper().startswith("UFDATA"):
        conn.database = conn.system_database

    return AppConfig(
        connection=conn,
        # data.get(键, 默认值)：取不到就用默认值，保证健壮性
        template_path=data.get("template_path", str(PROJECT_ROOT / DEFAULT_TEMPLATE)),
        output_dir=data.get("output_dir", str(DEFAULT_OUTPUT_DIR)),
        fiscal_year=data.get("fiscal_year", 2026),
        fiscal_month=data.get("fiscal_month", 5),
        account_set=data.get("account_set", "001"),
        last_selected_reports=data.get("last_selected_reports", []),
    )


def ensure_dirs() -> None:
    """确保运行需要用到目录都存在（输出目录、日志目录）。"""
    DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    (PROJECT_ROOT / "logs").mkdir(parents=True, exist_ok=True)