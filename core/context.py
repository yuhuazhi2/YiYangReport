# -*- coding: utf-8 -*-
"""
报表生成上下文与取数缓存
============================================================

【模块职责】
    1. 持有数据库连接、会计期间、账套信息等"运行期参数"。
    2. 提供带缓存的取数接口（fetch / fetch_dict），
       使多张报表可复用同一批 U8 数据，避免重复查询。
    3. 记录每张报表的取数来源（SQL + 结果快照），供核对与审计。

【什么是"上下文"（Context）？】
    "上下文"是一个贯穿整个生成流程的"共享对象"。
    生成报表时，43 张表都依赖同样的数据库、同样的会计期间。
    与其每张表都各自传一堆参数，不如把它们打包成一个 context 一路传下去。
    这是很常见的设计手法，Qt 等框架里也大量使用（如事件上下文）。

【为什么需要缓存？（性能关键）】
    例如 LK01 资产负债表和 LK02 利润表，可能都要查 GL_accsum（科目余额表）。
    如果每张表都去查一遍数据库，43 张表会重复查很多次，速度极慢。
    有了缓存，第一次查到结果后存起来，同一次生成中再要就直接返回。
    这就是"批量取数 + 单连接复用"性能策略的核心。

【缓存键约定】
    "库.表:期间:参数"，例如 "GL_accsum:2026-05"。
    把"库、表、期间、参数"都编进键里，保证"不同的查询不会撞车"。
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Optional

from core.db import Database
from core.logging_setup import get_logger

logger = get_logger(__name__)


@dataclass
class Period:
    """
    会计期间（年 + 月）。

    【新手须知：@dataclass 在本例的好处】
        Period(2026, 5) 就能直接构造，比手写 __init__ 省事。
        而且两个 Period(2026, 5) 对象对比时会被认为"相等"，
        这在做键、做判断时很方便。
    """

    year: int
    month: int

    @property
    def key(self) -> str:
        """
        标准化字符串键，如 2026 年 5 月 → "2026-05"。

        【新手须知：f"{self.month:02d}" 的 :02d】
            :02d 是"格式化说明符"：
              - 0  → 不足位补 0
              - 2  → 总宽度至少 2 位
              - d  → 十进制整数
            所以 5 会变成 "05"，10 仍是 "10"。
            这样字符串排序、比较时 "2026-05" < "2026-10" 才符合直觉。
        """
        return f"{self.year}-{self.month:02d}"

    def __str__(self) -> str:
        """
        定义"把对象转成字符串"时的样子。

        【新手必学：__str__ 魔术方法】
            它是 Python 的"魔术方法"（dunder method）之一。
            定义了它之后：
                print(period)        → 打印 "2026年5月"
                str(period)          → 得到 "2026年5月"
            如果没有定义，打印出来会是 <Period object at 0x...> 这种看不懂的东西。
        """
        return f"{self.year}年{self.month}月"


@dataclass
class FetchRecord:
    """
    一次取数记录（用于核对与审计）。

    【用途】
        生成完成后，用户能查看"这张报表的数据是从哪条 SQL 查出来的"，
        再拿去和用友前台的数据对比，快速定位差异。

    【新手难点：field(default_factory=...) 为什么不能直接写默认值？】
        ❌ 错误写法：created_at: str = datetime.now()...
           这样所有 FetchRecord 会共享"导入时"的那一个时间戳，全错。
        ✅ 正确写法：default_factory=lambda: datetime.now()...
           default_factory 是一个"函数"，每次创建新对象时才调用它，
           所以每条记录都能拿到"当下"的时间。
        【原因】Python 的默认参数只在函数定义时求值一次；
               可变或"每次都要新值"的默认值，必须用 default_factory。
    """

    report_code: str
    cache_key: str
    sql: str
    params: Any
    rowcount: int
    elapsed: float
    created_at: str = field(
        default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    )


class ReportContext:
    """
    报表生成上下文。

    在生成流程开始时创建一次，贯穿所有报表的 fetch/write。

    【生命周期示意】
        生成开始 → 创建 ReportContext → LK01 取数(填缓存) → LK02 复用缓存
                → ... → LK33 → 生成结束（对象被丢弃）
    """

    def __init__(self, db: Database, period: Period, account_set: str = "001"):
        """
        参数：
            db          : 已连接的 Database 对象。
            period      : 会计期间（Period）。
            account_set : 账套号，默认 "001"。
        """
        self.db = db
        self.period = period
        self.account_set = account_set

        # _cache：真正的取数缓存。键是字符串，值是查询结果（行列表）。
        self._cache: dict[str, Any] = {}

        # _audit：取数审计列表，每取一次就追加一条 FetchRecord。
        self._audit: list[FetchRecord] = []

        # report_data：每张报表记录其"关键指标值"，供勾稽核对使用。
        # 结构形如：{"LK01": {"货币资金_期末": 12345.67, ...}, "LK02": {...}}
        self.report_data: dict[str, dict] = {}

    # ------------------------------------------------------------------
    # 缓存
    # ------------------------------------------------------------------
    def make_key(self, table: str, *params, database: str = "") -> str:
        """
        构造缓存键，格式："库.表:期间:参数"。

        【新手必学：*params 可变位置参数】
            参数名前加一个星号，表示"接收任意多个位置参数，打包成元组"。
            调用示例：
                ctx.make_key("GL_accsum")              → params = ()
                ctx.make_key("Code", "1001")           → params = ("1001",)
                ctx.make_key("GL_accvouch", "2026", 5) → params = ("2026", 5)
            这样这个函数既能接收 0 个、也能接收多个参数，非常灵活。

        【新手必学：**kwargs 可变关键字参数】（本函数用 database 关键字更清晰）
            如果写成 **kwargs，则接收"任意多个 键=值"参数，打包成字典。
        """
        # 如果指定了数据库，就在键前面带上库名，避免跨库撞车
        db_part = f"{database}." if database else ""

        # 【新手必学：生成器表达式 + join】
        #   ":".join(str(p) for p in params)
        #   str(p) for p in params 是一个"生成器"，逐个把参数转成字符串，
        #   然后 join 用 ":" 把它们拼起来。
        #   结果示例：("2026", "05") → "2026:05"
        arg_part = ":".join(str(p) for p in params)

        return f"{db_part}{table}:{self.period.key}:{arg_part}"

    def cache_get(self, key: str) -> Optional[Any]:
        """取出缓存。没缓存过时返回 None。"""
        # 【新手必学：dict.get()】
        #   self._cache.get(key) 等价于：
        #       if key in self._cache: return self._cache[key]
        #       else: return None
        #   比 self._cache[key] 更安全：键不存在时不会抛 KeyError，而是返回 None。
        return self._cache.get(key)

    def cache_set(self, key: str, value: Any) -> None:
        """写入缓存。"""
        self._cache[key] = value

    def cache_has(self, key: str) -> bool:
        """判断缓存里有没有这个键。"""
        # in 运算符对字典来说，判断的是"键"是否存在
        return key in self._cache

    def clear_cache(self) -> None:
        """清空缓存（例如切换期间/账套后需要重新取数时调用）。"""
        self._cache.clear()

    # ------------------------------------------------------------------
    # 带缓存的查询
    # ------------------------------------------------------------------
    def fetch(self, key: str, sql: str, params=None,
              database: str = "", report_code: str = "") -> Any:
        """
        带缓存的查询：若缓存存在则直接返回，否则执行 SQL 并缓存。

        参数：
            key         : 缓存键（建议用 make_key 生成）。
            sql         : 要执行的 SQL。
            params      : SQL 参数。
            database    : 可选的数据库名。
            report_code : 哪张报表发起的取数（记入审计，便于核对）。

        【流程】
            先看缓存 → 命中则秒回
                    → 未命中则查库 → 结果存缓存 + 记账 → 返回
        """
        if key in self._cache:
            return self._cache[key]

        start = time.perf_counter()
        # database 为空字符串时传 None，因为 Database.query 约定用 None 表示"不切换库"
        # 【关键】tag=report_code：把"哪张报表发的查询"带给审计日志，
        #   这样 sql_audit.log 里每行都能看出归属，出问题时可按报表代码直接筛。
        #   注意它只进日志文件，不会出现在界面【运行日志】里。
        rows = self.db.query(sql, params, database=database or None, tag=report_code)
        elapsed = time.perf_counter() - start

        self._cache[key] = rows
        self._audit.append(FetchRecord(
            report_code=report_code,
            cache_key=key,
            sql=sql,
            params=params,
            rowcount=len(rows),
            elapsed=elapsed,
        ))
        return rows

    def fetch_dict(self, key: str, sql: str, params=None,
                   database: str = "", report_code: str = "") -> Any:
        """
        带缓存的查询，返回"字典列表"版本。

        和 fetch 的区别只在于底层调用了 db.query_dict，
        返回的是 [{列名: 值}, ...]，方便按列名取值。
        """
        if key in self._cache:
            return self._cache[key]

        start = time.perf_counter()
        rows = self.db.query_dict(sql, params, database=database or None,
                                  tag=report_code)
        elapsed = time.perf_counter() - start

        self._cache[key] = rows
        self._audit.append(FetchRecord(
            report_code=report_code, cache_key=key, sql=sql, params=params,
            rowcount=len(rows), elapsed=elapsed,
        ))
        return rows

    # ------------------------------------------------------------------
    # 报表数据登记（供核对）
    # ------------------------------------------------------------------
    def register_report_data(self, code: str, data: dict) -> None:
        """
        登记某报表的关键取数值，供勾稽核对使用。

        【为什么要单独登记？】
            fetch 缓存的是"原始查询行"，勾稽关系需要的是"报表口径的指标值"
            （比如"货币资金_期末"）。报表在 write 阶段算出这些值后，
            顺手登记进来，核对模块就能直接用，不用再解析 Excel。
        """
        self.report_data[code] = data

    def get_report_data(self, code: str) -> Optional[dict]:
        """取出某报表登记的数据（没有则返回 None）。"""
        return self.report_data.get(code)

    # ------------------------------------------------------------------
    # 审计
    # ------------------------------------------------------------------
    @property
    def audit_records(self) -> list[FetchRecord]:
        """
        返回全部取数审计记录。

        【注意】返回的是副本（list(...)），防止外部修改内部列表
        （防御性编程，思路和 template.py 里返回副本一样）。
        """
        return list(self._audit)

    def audit_summary(self) -> str:
        """
        生成一句审计汇总，如：
            "共 12 次取数，总耗时 3.45s，缓存 12 项"

        【新手必学：sum 配合生成器表达式】
            sum(r.elapsed for r in self._audit)
            把每条记录的耗时累加起来。
            f"{total:.2f}" 表示保留两位小数。
        """
        total = sum(r.elapsed for r in self._audit)
        return (f"共 {len(self._audit)} 次取数，"
                f"总耗时 {total:.2f}s，缓存 {len(self._cache)} 项")