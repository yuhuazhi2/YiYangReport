# -*- coding: utf-8 -*-
"""
数据库访问层（SQL Server 2016 / 用友 U8 后台）
============================================================

【模块职责】
    1. Database 类：管理 pyodbc 连接（连接复用）。
    2. 提供 query / query_one / scalar 等便捷查询方法。
    3. 每次查询自动记录 SQL 审计日志（便于与用友前台数据比对）。
    4. 支持连接"指定数据库"或运行时"切换数据库"（U8 多账套/多年度库）。

【新手须知：什么是"分层"？】
    这个模块属于"数据访问层"（Data Access Layer）。
    它只负责"跟数据库打交道"，不关心财务逻辑，也不管界面长什么样。
    好处是：将来换数据库（比如换 PostgreSQL），只需改这一层，其他代码不动。

【新手须知：为什么要"连接复用"？】
    每次查询都重新连一次数据库非常慢（TCP 握手 + 身份验证要几百毫秒）。
    所以我们在对象里保存一个连接 _conn，反复使用同一个连接，速度快很多。

【新手须知：为什么连接要 autocommit=True？】
    财务报表只需要"读"数据，不需要修改用友的账。
    打开自动提交可以避免长时间占用事务锁，影响别人用用友。
"""
from __future__ import annotations

# typing 用于"类型提示"，不是强制的，但能让编辑器帮你检查错误、代码更易读
from typing import Any, Iterable, Optional, Sequence
import time

from config.settings import ConnectionConfig, list_sqlserver_drivers
from core.logging_setup import SQLAudit, get_logger

# 每个模块都获取自己的日志器，__name__ 会自动填成 "core.db"
logger = get_logger(__name__)

# 【新手须知：try/except ImportError 的用途】
#   pyodbc 是第三方库，可能没装。若直接 import 失败，整个项目都没法运行。
#   这里捕获 ImportError 并把它设为 None，让"不连数据库的场景"也能导入本模块
#   （比如做静态检查、看代码结构）。真正的连接操作会检查 pyodbc 是否为 None。
try:
    import pyodbc
except ImportError:  # 便于无驱动的环境导入本模块做静态检查
    pyodbc = None  # type: ignore


class DatabaseError(Exception):
    """
    自定义异常类。

    【新手必学：为什么要自定义异常？】
        Python 自带很多异常（ValueError、KeyError...），但它们的名字不够具体。
        自定义 DatabaseError 让上层代码能精确捕获"数据库相关"的错误：
            try:
                db.query(...)
            except DatabaseError as e:
                QMessageBox.warning(self, "错误", str(e))
        如果只用 Exception，就无法区分"是数据库错了"还是"是别的错了"。

    【新手须知：括号里的 Exception】
        表示 DatabaseError 继承自 Exception。继承后它就是一种"异常"，
        可以被 raise 抛出、被 except 捕获。
    """


class Database:
    """SQL Server 数据库封装（连接复用）"""

    def __init__(self, config: ConnectionConfig):
        """
        构造函数（创建对象时自动调用）。

        【新手须知：__init__ 和 self】
            self 代表"这个对象自己"。self.config = config 的意思是：
            把外部传进来的配置，存到这个对象的一个属性上，
            以后其他方法里就能用 self.config 拿到它。

        参数：
            config: ConnectionConfig 对象，包含服务器、账号、库名等连接信息。
        """
        self.config = config
        self._conn = None      # 连接对象，尚未连接时是 None
                               # 下划线开头的属性表示"内部使用"，是 Python 的命名约定

    # ------------------------------------------------------------------
    # 连接管理
    # ------------------------------------------------------------------
    def connect(self) -> None:
        """建立数据库连接（若已连接则会被覆盖为新的连接）。

        【自动选择 ODBC 驱动（要点功能）】
            本工具面向普通财务人员，他们不懂"ODBC 驱动"。
            因此当配置里**没有指定驱动**（driver 为空）时，本方法会
            自动列出本机已安装的 SQL Server 驱动，按优先级**依次尝试**，
            哪个能连通就用哪个，并把成功的驱动名回填到配置里，
            下次启动直接优先使用它。

            若配置里已指定了具体驱动（比如上次连成功过），则只试那一个，
            避免每次都要把所有驱动挨个试一遍、拖慢启动。
        """
        # 检查 pyodbc 是否可用
        if pyodbc is None:
            raise DatabaseError(
                "未安装 pyodbc，请先执行: pip install pyodbc，"
                "并确保已安装 'ODBC Driver 17 for SQL Server'"
            )

        cfg = self.config
        # 记住调用前的驱动设置，若所有尝试都失败就还原，方便下次重新自动尝试
        original_driver = cfg.driver

        # ---- 第1步：确定"要尝试的驱动候选清单" ----
        # 先拿到"本机实际已安装"的驱动列表（可能为空）
        available = list_sqlserver_drivers()

        if cfg.driver.strip():
            # 情况A：配置里已指定驱动（例如上次连成功过）
            #   把它排在第一位优先尝试，但**后面追加其它可用驱动作为备选**。
            # 【新手须知】为什么还要追加备选？
            #   万一这个驱动后来被卸载/损坏了，程序仍能自动回落到别的驱动，
            #   财务人员不用懂原理，也不会因为驱动失效就无法用软件。
            specified = cfg.driver.strip()
            candidates = [specified] + [d for d in available if d != specified]
        else:
            # 情况B：未指定驱动（全新安装/首次使用）→ 完全自动
            candidates = available

        # 兜底：一个可用驱动都没有
        if not candidates:
            raise DatabaseError(
                "本机未找到可用的 SQL Server ODBC 驱动。\n"
                "请安装 'ODBC Driver 17 for SQL Server'（微软官网免费下载）后重试。"
            )
        logger.info("将依次尝试以下 ODBC 驱动：%s", "、".join(candidates))

        # ---- 第2步：逐个驱动尝试连接，成功即返回 ----
        errors: list[str] = []
        for drv in candidates:
            cfg.driver = drv   # 临时把驱动设为当前候选，以便 build_conn_str 使用
            try:
                self._conn = pyodbc.connect(
                    cfg.build_conn_str(),      # 由配置对象拼出连接字符串
                    timeout=cfg.timeout,        # 连接超时秒数
                    autocommit=True,            # 自动提交，不占事务锁
                )
                # 连通了！把可用驱动名保留在配置里（下次优先用它）
                logger.info("使用驱动 [%s] 连接成功：%s\\%s",
                            drv, cfg.server, cfg.instance or "(默认实例)")
                return
            except Exception as exc:  # noqa: BLE001
                # 该驱动不可用，记录原因后继续尝试下一个
                errors.append(f"  · {drv} → {exc}")
                logger.warning("驱动 [%s] 连接失败，继续尝试下一个：%s", drv, exc)

        # ---- 第3步：全部驱动都失败 ----
        cfg.driver = original_driver   # 还原，保持"自动"状态以便下次重试
        detail = "\n".join(errors)
        # 【新手须知：raise ... from None】
        #   这里不再保留单个原始异常（因为尝试了多个），
        #   用 from None 让报错信息更干净，直接把汇总详情作为消息抛出。
        raise DatabaseError(
            f"连接失败，已尝试 {len(candidates)} 个 ODBC 驱动，均未成功：\n{detail}"
        ) from None

    def close(self) -> None:
        """关闭数据库连接（释放资源）。"""
        if self._conn is not None:
            try:
                self._conn.close()
            except Exception:
                # 关闭失败也无所谓（比如连接已断），直接忽略。
                # 【新手须知】pass 是"什么都不做"的占位语句。
                pass
            self._conn = None    # 置空，标记"当前没有连接"
            logger.info("数据库连接已关闭")

    def __enter__(self):
        """
        支持 with 语句（进入时自动连接）。

        用法：
            with Database(cfg) as db:
                rows = db.query("SELECT ...")
            # 离开缩进块时自动 close()

        【新手必学：上下文管理器】
            实现了 __enter__ 和 __exit__ 的类，就能配合 with 使用，
            保证资源"用完自动清理"，即使中途报错也会清理。
        """
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """支持 with 语句（退出时自动关闭）。返回 False 表示不吞掉异常。"""
        self.close()
        return False

    @property
    def connected(self) -> bool:
        """
        判断当前连接是否可用。

        【新手必学：@property 装饰器】
            加上它之后，调用时不用写括号：
                if db.connected:      # 注意：不是 db.connected()
            看起来像访问属性，实际执行的是这个方法。
            适合表达"派生出来的状态"，读起来更自然。
        """
        if self._conn is None:
            return False
        try:
            # 真正执行一条最简单的语句，验证连接还活着
            self._conn.cursor().execute("SELECT 1")
            return True
        except Exception:
            # 连接可能已超时/断开
            return False

    def test_connection(self) -> tuple[bool, str]:
        """
        测试连接，返回 (是否成功, 提示消息)。

        【新手必学：返回多个值】
            Python 的 return a, b 实际返回的是元组 (a, b)，
            调用方可以这样接：
                ok, msg = db.test_connection()

        界面上的"测试连接"按钮就调用这个方法。
        """
        try:
            self.connect()
            cur = self._conn.cursor()
            cur.execute("SELECT @@VERSION")   # SQL Server 内置变量，返回版本信息
            ver = cur.fetchone()[0]
            # 版本字符串有多行，只取第一行更简洁
            short = ver.splitlines()[0] if ver else "未知版本"
            return True, f"连接成功\n{short}"
        except Exception as exc:
            return False, f"连接失败：{exc}"

    def switch_database(self, database: str) -> None:
        """
        切换当前连接的默认数据库（用于 U8 多账套/多年度库）。

        【背景知识：用友 U8 的数据组织】
            用友 U8 通常是"一套账"对应一个数据库，命名类似 UFDATA_001_2026，
            其中 001 是账套号，2026 是年度。
            想查不同年度/账套的数据，切换数据库最方便。
        """
        self._ensure_conn()
        # 【安全提示】库名不能用参数占位符（?）传递，只能拼接到 SQL 里，
        # 所以用方括号 [] 包住库名，防止特殊字符或 SQL 注入。
        self._conn.execute(f"USE [{database}]")
        self.config.database = database     # 同步更新配置，方便下次复用
        logger.info("已切换数据库：%s", database)

    def list_databases(self) -> list[str]:
        """列出所有数据库（用于连接后选择 U8 年度库）。"""
        rows = self.query(
            # database_id > 4 是为了排除系统库(master/model/msdb/tempdb)
            "SELECT name FROM sys.databases WHERE database_id > 4 ORDER BY name"
        )
        # 【新手必学：列表推导式】
        #   [r[0] for r in rows] 等价于：
        #       result = []
        #       for r in rows:
        #           result.append(r[0])
        #   把"每行的第一列（库名）"抽出来组成新列表。
        return [r[0] for r in rows]

    def _ensure_conn(self):
        """
        内部辅助方法：确保有连接可用（没有就自动连）。

        下划线开头表示"私有方法"，只在本类内部使用。
        """
        if self._conn is None:
            self.connect()

    # ------------------------------------------------------------------
    # 查询接口
    # ------------------------------------------------------------------
    def query(self, sql: str, params: Optional[Sequence[Any]] = None,
              database: Optional[str] = None) -> list[tuple]:
        """
        执行查询，返回全部行（list[tuple]，每个元组是一行）。

        参数：
            sql      : SQL 语句，用 ? 作为参数占位符
            params   : 与 ? 一一对应的参数值（可防止 SQL 注入）
            database : 可选。若指定，先 USE 切到该库再执行查询

        【新手必学：为什么要用参数占位符 ?】
            ❌ 错误写法（危险）：
                sql = f"SELECT * FROM Code WHERE cCode = '{code}'"
                如果有人传入 code = "' OR '1'='1"，就成了 SQL 注入。
            ✅ 正确写法：
                db.query("SELECT * FROM Code WHERE cCode = ?", [code])
                数据库会安全地处理这个值，杜绝注入。

        【新手必学：Optional[...] 类型提示】
            Optional[Sequence[Any]] 表示"可以是序列，也可以是 None"。
            Optional[X] 等价于 Union[X, None]。
        """
        self._ensure_conn()
        start = time.perf_counter()   # 开始计时（高精度计时器）
        try:
            cur = self._conn.cursor()
            if database:
                cur.execute(f"USE [{database}]")
            # params 为 None 时用空列表代替，避免 pyodbc 报错
            cur.execute(sql, params or [])
            rows = cur.fetchall()                 # 取出所有结果行
            elapsed = time.perf_counter() - start # 计算耗时
            SQLAudit.record(sql, params, elapsed, len(rows))   # 写审计日志
            # 把 pyodbc 的 Row 对象统一转成普通元组，方便后续处理
            return [tuple(r) for r in rows]
        except Exception as exc:
            elapsed = time.perf_counter() - start
            # 出错了也要记审计日志，标注错误内容
            SQLAudit.record(sql, params, elapsed, 0, error=str(exc))
            logger.error("查询失败: %s | SQL=%s", exc, sql)
            raise DatabaseError(f"查询失败：{exc}") from exc

    def query_dict(self, sql: str, params: Optional[Sequence[Any]] = None,
                   database: Optional[str] = None) -> list[dict]:
        """
        执行查询，返回字典列表（每个字典的键是列名，值是单元格）。

        返回示例：
            [{"cCode": "1001", "cName": "库存现金"},
             {"cCode": "1002", "cName": "银行存款"}]

        【新手必学：dict(zip(cols, r))】
            - zip(cols, r) 把"列名列表"和"一行数据"两两配对：
                [("cCode", "1001"), ("cName", "库存现金")]
            - dict(...) 把这种配对列表转成字典。
            组合起来就是"把一行数据包装成 {列名: 值}"。
            这就是为什么取数时用 query_dict 更方便——可以直接 row["cName"]。
        """
        self._ensure_conn()
        start = time.perf_counter()
        try:
            cur = self._conn.cursor()
            if database:
                cur.execute(f"USE [{database}]")
            cur.execute(sql, params or [])
            # cur.description 包含结果集的列信息，d[0] 是列名
            cols = [d[0] for d in cur.description]
            rows = cur.fetchall()
            elapsed = time.perf_counter() - start
            SQLAudit.record(sql, params, elapsed, len(rows))
            return [dict(zip(cols, r)) for r in rows]
        except Exception as exc:
            elapsed = time.perf_counter() - start
            SQLAudit.record(sql, params, elapsed, 0, error=str(exc))
            raise DatabaseError(f"查询失败：{exc}") from exc

    def query_one(self, sql: str, params: Optional[Sequence[Any]] = None):
        """返回单行（没有结果则返回 None）。常用于"只查一条记录"的场景。"""
        rows = self.query(sql, params)
        # 【新手须知：三目（条件）表达式】
        #   rows[0] if rows else None
        #   等价于：如果 rows 非空就取第一行，否则返回 None。
        #   【重要】空列表在 Python 里视为 False，所以 if rows 就是在判断"有没有数据"。
        return rows[0] if rows else None

    def scalar(self, sql: str, params: Optional[Sequence[Any]] = None,
               default: Any = None) -> Any:
        """
        返回单个值（第一行第一列）。例如查合计金额。

        参数 default：没有结果时返回的默认值（默认 None）。
        """
        row = self.query_one(sql, params)
        if row is None:
            return default
        return row[0]

    def execute(self, sql: str, params: Optional[Sequence[Any]] = None) -> int:
        """
        执行非查询语句（INSERT/UPDATE/DELETE/建表等），返回受影响行数。

        【注意】本项目主要做"读账"，一般不会用到这个方法写用友的库，
        保留它是为了通用性（比如创建临时表、临时视图辅助取数）。
        """
        self._ensure_conn()
        start = time.perf_counter()
        try:
            cur = self._conn.cursor()
            cur.execute(sql, params or [])
            n = cur.rowcount    # 受影响行数
            elapsed = time.perf_counter() - start
            SQLAudit.record(sql, params, elapsed, n)
            return n
        except Exception as exc:
            elapsed = time.perf_counter() - start
            SQLAudit.record(sql, params, elapsed, 0, error=str(exc))
            raise DatabaseError(f"执行失败：{exc}") from exc

    # ------------------------------------------------------------------
    # 元数据辅助（"元数据"= 描述数据的数据，比如有哪些表、表里有哪些列）
    # ------------------------------------------------------------------
    def list_tables(self, database: str, like: str = "") -> list[str]:
        """
        列出指定库中的表名。

        参数 like：模糊匹配关键字。留空则列出全部表。
                   比如 like="accvouch" 只列出名字含 accvouch 的表。
        """
        sql = ("SELECT name FROM sys.tables "
               "WHERE name LIKE ? ORDER BY name")
        # 用 % 包住关键字实现"包含"匹配（% 是 SQL 通配符）
        rows = self.query(sql, [f"%{like}%"], database=database)
        return [r[0] for r in rows]

    def list_columns(self, database: str, table: str) -> list[tuple]:
        """
        列出指定表的列（列名、数据类型、最大长度）。

        【实用场景】
            拿到一个陌生的 U8 年度库，先用这个方法看看
            某张表有哪些字段，再决定怎么取数。
        """
        sql = ("SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH "
               "FROM INFORMATION_SCHEMA.COLUMNS "
               "WHERE TABLE_NAME = ? ORDER BY ORDINAL_POSITION")
        # ORDINAL_POSITION 是列在表中的顺序号，按它排序能保持原始列顺序
        return self.query(sql, [table], database=database)