# -*- coding: utf-8 -*-
"""
日志与 SQL 审计模块
============================================================

【模块职责】
    1. 初始化应用日志：把运行过程记录到 logs/app.log，便于事后排查问题。
    2. 提供 get_logger(name)：全项目统一用它获取"日志器"。
    3. SQLAudit：专门记录每一条发往数据库的 SQL 语句，用于核查取数是否正确。

【新手须知：为什么要记日志而不是用 print？】
    - print 只能输出到屏幕，程序一关就没了；
    - 日志可以同时写到文件，保留历史，方便出问题时追溯；
    - 日志有"级别"（DEBUG/INFO/WARNING/ERROR），可以按重要程度筛选；
    - 生产环境一般不打印到屏幕，而是写文件。
"""
from __future__ import annotations

import logging                    # Python 标准日志库
import time                       # 用于计时（计算一段代码跑了多久）
from logging.handlers import RotatingFileHandler   # 会"自动滚动"的文件处理器
from pathlib import Path


# 计算项目根目录（本文件在 core/ 下，往上两级就是项目根目录）
PROJECT_ROOT = Path(__file__).resolve().parent.parent
LOG_DIR = PROJECT_ROOT / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)    # 目录不存在就创建

APP_LOG = LOG_DIR / "app.log"           # 应用运行日志
SQL_AUDIT_LOG = LOG_DIR / "sql_audit.log"   # SQL 审计日志

# 模块级"开关"，防止 init_logging() 被重复执行导致日志重复输出
_initialized = False


def init_logging(level: int = logging.INFO) -> None:
    """
    初始化"根日志器"（整个程序共用的最上层日志器）。

    【关于"重复初始化"】
        如果多次调用会往根日志器上挂多个处理器，导致同一条日志打印好几遍。
        所以我们用 _initialized 这个"开关"保证只初始化一次（这叫"幂等"）。

    参数 level：日志级别。INFO 表示"记录 INFO 及以上级别的日志"，
               更细的 DEBUG 级别会被忽略。
    """
    # global 声明：告诉 Python 下面要修改的是"模块级的 _initialized"，
    # 而不是新建一个同名的局部变量。
    global _initialized
    if _initialized:
        return

    # 定义日志格式：时间 [级别] 日志器名: 消息
    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    root = logging.getLogger()      # 获取根日志器
    root.setLevel(level)

    # ---- 处理器1：写入文件（带滚动，防止日志文件无限增大）----
    # RotatingFileHandler 参数含义：
    #   maxBytes=5*1024*1024 → 单个文件最大 5MB
    #   backupCount=5        → 最多保留 5 个历史备份（app.log.1、app.log.2 ...）
    #   encoding="utf-8"     → 中文不乱码
    fh = RotatingFileHandler(
        APP_LOG, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    fh.setFormatter(fmt)
    root.addHandler(fh)

    # ---- 处理器2：同时输出到控制台 ----
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    root.addHandler(ch)

    _initialized = True


def get_logger(name: str) -> logging.Logger:
    """
    获取一个命名日志器。

    【用法约定】
        每个模块顶部写：logger = get_logger(__name__)
        __name__ 会自动是当前模块的名字（如 "core.db"），
        这样日志里就能看出是哪一块代码打印的，便于排查。
    """
    init_logging()   # 保证日志系统已初始化（重复调用也无害）
    return logging.getLogger(name)


class SQLAudit:
    """
    SQL 审计记录器。

    【为什么需要它？】
        财务报表最怕"取数取错了却没人发现"。
        把每次查询的 SQL、参数、耗时、返回行数都记下来，
        就能和用友前台/后台数据逐条比对，快速定位问题。

    【新手须知：类方法 vs 实例方法】
        本类的方法都用 @classmethod 装饰，第一个参数是 cls（类本身）。
        这样的方法"不用创建对象就能调用"，例如直接 SQLAudit.record(...)。
        而实例方法第一个参数是 self（对象自身），必须先创建对象。
    """

    _logger: logging.Logger | None = None   # 类属性：所有实例共用同一个日志器

    @classmethod
    def _get_logger(cls) -> logging.Logger:
        """懒加载：第一次调用时才创建审计日志器，之后复用。"""
        if cls._logger is None:
            lg = logging.getLogger("sql_audit")
            lg.setLevel(logging.INFO)
            # propagate=False：不让这条日志"向上冒泡"到根日志器，
            # 避免它同时被写进 app.log（我们想让它单独待在 sql_audit.log）
            lg.propagate = False

            fh = RotatingFileHandler(
                SQL_AUDIT_LOG, maxBytes=10 * 1024 * 1024, backupCount=10,
                encoding="utf-8",
            )
            fh.setFormatter(logging.Formatter("%(asctime)s %(message)s"))
            lg.addHandler(fh)
            cls._logger = lg
        return cls._logger

    @classmethod
    def record(cls, sql: str, params=None, elapsed: float = 0.0,
               rowcount: int = 0, error: str = "", tag: str = "") -> None:
        """
        记录一次数据库操作。

        参数：
            sql      : 执行的 SQL 语句
            params   : SQL 参数（用参数化查询可防止 SQL 注入）
            elapsed  : 耗时（秒）
            rowcount : 影响/返回的行数
            error    : 若有错误，记录错误信息
            tag      : 可选标记，写明"这条 SQL 是哪张报表发起的"（如 "LK01"）。

        【新手必学：为什么要加 tag？】
            一张报表生成失败时，维护人员打开 sql_audit.log 会看到几十上百条 SQL。
            如果每行都带 [LK01] 这样的标记，就能一眼筛出"出问题那张表取过哪些数"，
            不用凭时间顺序去猜。这是"可追溯性"（traceability）的典型做法。

        【重要：本日志只写文件，绝不进界面】
            类属性 _logger 设置了 propagate=False，保证这些内容不会冒泡到
            根日志器（也就不会出现在控制台或 app.log）。
            因此"SQL 语句"这一类技术细节，天然被隔离在审计文件里。
        """
        lg = cls._get_logger()

        # 把 SQL 中的换行和多余空格压成一个空格，让日志更紧凑、每行一条
        sql_one = " ".join(str(sql).split())

        # 三目写法：有错误就标 ERR(错误内容)，否则标 OK
        status = "OK" if not error else f"ERR({error})"

        # 有 tag 就显示成 "[LK01] "，没有就留空，保证格式统一不乱
        tag_part = f"[{tag}] " if tag else ""

        # 用 "{}" 占位符格式化，比直接用 f-string 更专业：
        # 这样日志库能延迟处理字符串拼接，且更容易按需过滤。
        lg.info(
            "[%s] %srows=%d elapsed=%.3fs params=%s sql=%s",
            status, tag_part, rowcount, elapsed, params, sql_one,
        )


class timed_query:
    """
    计时并审计一段数据库操作的"上下文管理器"。

    【新手必学：with 语句与上下文管理器】
        常见用法：
            with timed_query("SELECT ...") as tq:
                rows = 执行查询()
                tq.rowcount = len(rows)
        在 with 块开始时自动调用 __enter__，结束时（无论是否出错）
        自动调用 __exit__。非常适合"进入时开资源、退出时关资源"
        以及本例这种"自动计时"的场景。
    """

    def __init__(self, sql: str, params=None):
        self.sql = sql
        self.params = params
        self.start = 0.0
        self.rowcount = 0

    def __enter__(self):
        """进入 with 块时执行：记录起始时间。"""
        self.start = time.perf_counter()   # perf_counter 是高精度计时器，适合测耗时
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """
        退出 with 块时执行：计算耗时并写审计日志。

        参数 exc_type/exc_val/exc_tb 描述"是否发生了异常"：
           - 没出错时，exc_type 为 None
           - 出错时，exc_type 是异常类型
        返回 False 表示"我不处理这个异常，让它继续往上抛"。
        """
        elapsed = time.perf_counter() - self.start
        if exc_type is None:
            # 正常结束：记录成功
            SQLAudit.record(self.sql, self.params, elapsed, self.rowcount)
        else:
            # 出错：记录错误信息，方便排查
            SQLAudit.record(self.sql, self.params, elapsed, self.rowcount,
                            error=str(exc_val))
        return False