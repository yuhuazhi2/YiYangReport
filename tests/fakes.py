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


class FakeCtx:
    """假的报表上下文：fetch() 返回预设的行,不做任何数据库操作。

    参数：
        rows  : 模拟的查询结果,形如 [(ccode, iperiod, mb, me), ...]
        year  : 会计年度（默认 2026）
        month : 会计期间月份（默认 5）
        fail  : 为 True 时 fetch 抛异常,用于测"取数失败"的容错路径
    """

    def __init__(self, rows, year: int = 2026, month: int = 5,
                 fail: bool = False):
        self.period = Period(year, month)
        self._rows = rows
        self._fail = fail
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
            "report_code": report_code,
        })
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

    def __init__(self):
        # {单元格地址: 写入的值}
        self.written: dict[str, object] = {}
        # 记录写入顺序,便于验证"先设格式再写值"这类顺序要求
        self.events: list[tuple] = []

    # ---- 写 ----
    def set(self, cell: str, value) -> None:
        self.written[cell] = value
        self.events.append(("set", cell, value))

    def set_rc(self, row: int, col: int, value) -> None:
        self.set(f"{col_letter(col)}{row}", value)

    def set_text(self, cell: str, value) -> None:
        self.written[cell] = value
        self.events.append(("set_text", cell, value))

    def set_text_rc(self, row: int, col: int, value) -> None:
        self.set_text(f"{col_letter(col)}{row}", value)

    # ---- 读 ----
    def get(self, cell: str):
        return self.written.get(cell)

    def get_rc(self, row: int, col: int):
        return self.written.get(f"{col_letter(col)}{row}")

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