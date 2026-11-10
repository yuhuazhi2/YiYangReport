# -*- coding: utf-8 -*-
"""测试用的"假对象"与造数工具。

【为什么需要这个文件？】
    报表类的 fetch() 依赖 context 取数、write() 依赖 hook 写单元格。
    但测试不该依赖真实环境：
      · 真的连 SQL Server → 测试只能在部署了 U8 的机器上跑；
      · 真的拉起 Excel    → 慢,且可能残留 EXCEL.EXE 进程。

    所以这里用两个"假对象"顶替：
      FakeCtx  —— 假装是 ReportContext,fetch() 直接返回我们指定的行；
      FakeHook —— 假装是 _SheetHook,把写入的单元格记在字典里。

    这样报表逻辑（取数分流、汇总、写哪一格）就能在毫秒级验证完,
    不需要数据库,也不需要 Excel。

【使用范围】
    只测"逻辑"。真实 Excel 的写入效果（格式、日期误判等）不在这里测 ——
    那类问题必须用真 Excel 才能暴露,属手工验证范畴。
"""
from __future__ import annotations

from core.context import Period


# 测试里用的"本年库名" / "上年库名"。
# 【为什么要给假库名，而不是留空？】
#   LK02 的上年取数逻辑会先判断 ctx.prev_database 是否为空：
#   为空就跳过上年查询（视为首年）。若测试里也留空，
#   那些想验证"上年能取到数"的用例就永远走不到取数分支。
TARGET_DB = "UFDATA_001_2026"
PREV_DB = "UFDATA_001_2025"
# 【默认"无上年数据"】
#   绝大多数既有测试只关心"本年怎么填"，不关心 P 列。
#   若默认给一个上年库，每次生成都会多查一轮，所有"查询次数"断言
#   都会被翻倍打断 —— 而那些断言本身是对的，只是不该被上年打扰。
#   所以默认 prev_database=""（等价于账套首年），
#   需要验证上年功能的用例显式传 prev_database=PREV_DB。
NO_PREV_DB = ""


class FakeCtx:
    """假的报表上下文：fetch() 返回预设的行,不做任何数据库操作。

    参数：
        rows       : 本年查询结果,形如 [(ccode, iperiod, amount), ...]
        year       : 会计年度（默认 2026）
        month      : 会计期间月份（默认 5）
        fail       : 为 True 时 fetch 抛异常,用于测"取数失败"的容错路径
        prev_rows  : 对比期（上年/上期）查询结果。
                     None = 该账套无对比期数据（如首年）
        prev_fail  : 为 True 时"查上年库"的那次 fetch 抛异常

    【为什么把本年和上年的假数据分开？】
        两年查的是不同的库、不同的年度，返回的必然是不同批次数据。
        若共用一个列表，就测不出"P 列取的是上年而不是本年"这类关键错误 ——
        而这正是本功能最需要守住的地方。
    """

    def __init__(self, rows, year: int = 2026, month: int = 5,
                 fail: bool = False,
                 prev_rows=None, prev_fail: bool = False,
                 target_database: str = TARGET_DB,
                 prev_database: str = NO_PREV_DB,
                 balance_rows=None, balance_prev_rows=None):
        self.period = Period(year, month)
        self._rows = rows
        self._fail = fail
        self._prev_rows = prev_rows
        self._prev_fail = prev_fail
        # 【LK03 需求31 新增：余额类查询（GL_accsum）的假数据】
        #
        #   同一张报表里有两条取数线：
        #     · gl_cashtable —— 现金流量项目，形如 (cCashItem, iYPeriod, md, mc)
        #     · GL_accsum    —— 科目余额，    形如 (ccode, iYPeriod, mb)
        #   两者的列数与语义都不同，不能共用一批假数据。
        #
        #   若共用一个列表，余额聚合会把 '01' 当成科目编码去匹配
        #   '1001'/'1002'，永远匹配不上 → C45 静默留空，
        #   而测试还以为"取数逻辑没问题"。
        #
        #   默认 None → 余额线返回空列表（模拟"未取到余额"），
        #   这与改动前的行为一致（改动前也是匹配不上、留空），
        #   所以既有 525 条测试不受影响。
        self._balance_rows = balance_rows
        self._balance_prev_rows = balance_prev_rows
        self.target_database = target_database
        self.prev_database = prev_database
        # 记录每次 fetch 的调用信息,便于断言"查了哪些科目"
        self.fetch_calls: list[dict] = []

    def make_key(self, table: str, *params, database: str = "") -> str:
        # 缓存键的具体格式与测试无关,只要唯一即可
        return f"{table}:{':'.join(str(p) for p in params)}"

    def fetch(self, key: str, sql: str, params=None, database: str = "",
              report_code: str = ""):
        self.fetch_calls.append({
            "key": key,
            "sql": sql,
            "params": params,
            "database": database,
            "report_code": report_code,
        })
        # 【LK03 需求31：余额类查询的假数据分流】
        #   为什么用 SQL 文本判断，而不是用 key？
        #     key 的格式由 make_key 决定，属于"报表自己起的名字"，
        #     改一次命名就会让这里的判断失效，而且不会报错。
        #     SQL 里出现哪张表是**取数函数**决定的，与报表层命名无关。
        #
        #   【★ 只在显式给了余额假数据时才分流 ★】
        #     GL_accsum 不是 LK03 专用的 —— LK01 资产负债表也查这张表。
        #     若只凭 "GL_accsum" 就一律返回余额假数据，
        #     LK01 会拿到 None→[]，整张表静默变成空白（实测 29 条测试失败）。
        #     所以仅当调用方**确实传了**余额假数据（非 None）时才走这条分支，
        #     其余情况完全按原有逻辑返回，对既有报表零影响。
        has_balance_fixture = (self._balance_rows is not None
                               or self._balance_prev_rows is not None)
        if has_balance_fixture and "GL_accsum" in (sql or ""):
            if report_code.endswith(("-上年", "-上期")):
                return list(self._balance_prev_rows or [])
            return list(self._balance_rows or [])
        # 【靠 report_code 区分本期与对比期】
        #   生产代码里：LK02 的上年那批查询记为 "LK02-上年"，
        #               LK03 的上期那批查询记为 "LK03-上期"。
        #   用后缀判断，比"看 database 是不是 PREV_DB"更直接，
        #   因为测试也可能故意把两者设成同一个库（跨年度单库场景）——
        #   那种情况下按库名分辨根本区分不开。
        #
        #   【为什么把 "-上期" 也纳入？】
        #     LK03 现金流量表的 D 列取"上一期"：1 月时是上年 12 月，
        #     其余月份是同年上月，语义上叫"上期"更准确。
        #     现有 LK02 只用 "-上年"，此改动对它没有任何影响。
        is_prev = report_code.endswith(("-上年", "-上期"))
        if is_prev:
            if self._prev_fail:
                raise RuntimeError("模拟上年取数失败")
            if self._prev_rows is None:
                return []
            return list(self._prev_rows)

        if self._fail:
            raise RuntimeError("模拟取数失败")
        return list(self._rows)

    def register_report_data(self, code: str, data: dict) -> None:
        # 勾稽核对用的登记接口,测试中不需要做事
        pass


class FakeHook:
    """假的表页句柄：把写入的单元格记录在 self.written 里。

    只实现报表实际用到的方法。若某个报表调用了未实现的方法,
    测试会立刻报 AttributeError —— 这恰好能提示"框架接口有变动"。
    """

    def __init__(self, formula_values: dict | None = None):
        # {单元格地址: 写入的值}
        self.written: dict[str, object] = {}
        # {单元格地址: 数字格式串}
        # 【为什么单独存一份】
        #   报表可能为"负数格"改用不带填充的格式，
        #   而值本身不变。若只记录 written，就测不出
        #   "格式有没有被改对" —— 而这正是本次修复的全部内容。
        self.number_formats: dict[str, str] = {}
        # 记录写入顺序,便于验证"先设格式再写值"这类顺序要求
        self.events: list[tuple] = []
        # {单元格地址: 公式计算结果}
        #
        # 【为什么需要它？】
        #   真实 Excel 里，公式行（小计/净额）的值由 Excel 算出，
        #   程序只读不写 —— 所以 written 里永远没有这些格子。
        #   但 LK03 需要"读取公式行的值判断正负，再修正格式"，
        #   若假对象读不到值，那条逻辑在测试里永远走不到，
        #   等于没测。
        #   这里允许测试预设"Excel 算出来的结果"，让读取路径可验证。
        self.formula_values: dict[str, object] = dict(formula_values or {})

    # ---- 写 ----
    def set(self, cell: str, value) -> None:
        self.written[cell] = value
        self.events.append(("set", cell, value))

    def set_rc(self, row: int, col: int, value) -> None:
        self.set(f"{col_letter(col)}{row}", value)

    def set_text(self, cell: str, value) -> None:
        self.written[cell] = value
        self.events.append(("set_text", cell, value))

    def set_number_format(self, cell: str, fmt: str) -> None:
        """记录数字格式设置（对应 _SheetHook.set_number_format）。"""
        self.number_formats[cell] = fmt
        self.events.append(("number_format", cell, fmt))

    def set_text_rc(self, row: int, col: int, value) -> None:
        self.set_text(f"{col_letter(col)}{row}", value)

    # ---- 读 ----
    def get(self, cell: str):
        # 写入过的格子以实际写入值为准；否则取"公式计算结果"
        # 【为什么写入优先？】
        #   若某格既被写入又有预设值，说明预设已过时，
        #   应以程序刚写进去的值为准 —— 否则测出来的是旧数据。
        if cell in self.written:
            return self.written[cell]
        return self.formula_values.get(cell)

    def get_rc(self, row: int, col: int):
        return self.get(f"{col_letter(col)}{row}")

    # ---- 定位（页眉填写会用到）----
    def find_cell(self, keyword: str, start_row: int = 1, max_row: int = 0,
                  max_col: int = 0) -> tuple[int, int]:
        # 假对象里没有真实表格内容,一律"找不到",页眉函数会安全跳过
        return (0, 0)

    def find_row(self, keyword: str, col: int = 1, start_row: int = 1,
                 max_row: int = 0, exact: bool = False) -> int:
        return 0

    # ---- 边界 ----
    @property
    def used_rows(self) -> int:
        return 60

    @property
    def used_cols(self) -> int:
        return 12

    @property
    def name(self) -> str:
        return "FAKE"


def col_letter(col: int) -> str:
    """把列号转成 Excel 列字母（1→A, 2→B, ..., 27→AA）。

    【为什么不用 chr(64+col)？】
        那个写法只支持 1~26 列。而本模板里有 49 列、256 列的表页
        （LK21、LK06-1）,超过 26 列就会算出乱码。
        这里用标准的进制转换实现,支持任意列数。
    """
    letters = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def make_rows(codes_map: dict, month: int) -> list:
    """造"本期 + 本年1月"两批查询结果,方便写测试。

    参数：
        codes_map : {科目编码: (年初余额mb, 期末余额me)}
        month     : 本期月份（如 5）
    返回：
        [(ccode, iperiod, mb, me), ...]

    【为什么要这样造数？】
        本项目的关键口径是"年初取 1 月的 mb、期末取本期的 me"。
        造数时让两批数据**刻意不同**,才能验证程序没取错期间：
        若程序误取本期 mb,结果就会不符而测试失败。

        【特别注意】本期记录的 mb 故意填 0.0。
        若程序错误地取本期 mb 当年初,就会得到 0 —— 一眼就能发现。
    """
    rows = []
    for code, (begin_val, end_val) in codes_map.items():
        if month == 1:
            # 本期就是 1 月：同一行同时提供 mb(年初) 与 me(期末)
            rows.append((code, 1, begin_val, end_val))
        else:
            rows.append((code, 1, begin_val, 0.0))      # 1月记录
            rows.append((code, month, 0.0, end_val))    # 本期记录
    return rows


def approx(a, b, tol: float = 0.01) -> bool:
    """浮点比较：绝对值差小于容差即视为相等。"""
    if a is None or b is None:
        return a is b
    return abs(float(a) - float(b)) <= tol