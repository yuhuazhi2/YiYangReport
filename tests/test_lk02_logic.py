# -*- coding: utf-8 -*-
"""LK02 利润表的逻辑测试。

运行方式（在项目根目录执行）：
    python -m unittest discover -s tests -v

【本文件测什么？】
    1. 会计期间是否写入 C4，且必须用 set_text（不能用 set）
    2. C4 是否只写一次（防止通用页眉叠加写第二处年月）
    3. 【结构护栏】公式行不得出现在 ROW_SPECS（否则会覆盖模板公式）
    4. 【结构护栏】不同行不得占用同一 Excel 行
    5. 【结构护栏】excel_row 必须来自 EXCEL_ROW 对照表
    6. 【结构护栏】ROW_SPECS 的格子是否都登记进 cell_map
    7. 直接读模板，校验坐标确实落在有项目名的行上

    第 3~6 条是**长期护栏**：以后逐条追加项目时若写错，
    测试会立刻失败，不必等生成报表才发现。

【不测什么？】
    真实 Excel 的写入效果（格式、日期误判）—— 那必须用真 Excel，
    属手工验证范畴，见 docs/报表开发规范.md 的说明。
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

# 让测试既能通过 discover 运行，也能单独 python tests/xxx.py 运行
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from tests.fakes import FakeCtx, FakeHook, approx  # noqa: E402

from reports.lk02_profit import (            # noqa: E402
    COL_OF_MONTH,
    COL_YEAR_TOTAL,
    EXCEL_ROW,
    FORMULA_ROWS,
    PERIOD_CELL,
    ROW_SPECS,
    LK02Report,
)


def make_amount_rows(codes_map: dict) -> list:
    """造 LK02 的查询结果：[(ccode, iperiod, amount), ...]。

    参数 codes_map : {科目编码: {月份: 金额}}

    【为什么不用 tests/fakes.py 的 make_rows()？】
        那个是给 LK01 造的 `(ccode, iperiod, mb, me)` —— 余额格式。
        LK02 走 GL_accvouch 聚合，只有三列 `(ccode, iperiod, amount)`。
        这里若强行复用，会因为列数/语义不同而产生"看起来跑通了、
        实际上字段错位"的测试假阳性。宁可另写一个 6 行的小函数。
    """
    rows = []
    for code, by_month in codes_map.items():
        for month, amount in by_month.items():
            rows.append((code, month, amount))
    return rows


def run_report(rows=None, month: int = 5, year: int = 2026, fail: bool = False):
    """跑一遍 LK02 的 fetch + write，返回 (hook, ctx)。"""
    ctx = FakeCtx(rows or [], year=year, month=month, fail=fail)
    hook = FakeHook()
    rep = LK02Report()
    rep.fetch(ctx)
    rep.write(hook, ctx)
    return hook, ctx


class TestPeriodCell(unittest.TestCase):
    """验证会计期间写入 C4（需求第 1 条）。"""

    def test_period_written_as_text(self):
        """C4 应等于"2026年5月"，且必须用 set_text 写入。"""
        hook, _ = run_report(month=5, year=2026)
        self.assertEqual(hook.get(PERIOD_CELL), "2026年5月")

        # C4 位于金额表头区，普通 set() 会被 Excel 识别成日期序列号
        # （2026-05-01 → 46143，再按金额格式显示成 46,143.00）。
        kinds = [e[0] for e in hook.events if e[1] == PERIOD_CELL]
        self.assertIn("set_text", kinds,
                      f"C4 未用 set_text 写入，实际事件：{hook.events}")
        self.assertNotIn("set", kinds,
                         "C4 不应使用普通 set 写入（会被识别成日期）")

    def test_period_follows_selected_month(self):
        """换个月份生成时，C4 必须跟着变（不能写死）。

        【这条防的是"把年月写死成常量"】
            C4 的年月来自 ctx.period（界面录入值）。
            若有人图省事写成死值（例如固定 "2026年5月"），
            那么无论用户选哪个月，报表页眉永远显示 5 月 ——
            而且数字是对的（取数用的是真实期间），
            只有页眉错，极易被漏检。
        """
        for year, month, expect in (
            (2026, 5, "2026年5月"),
            (2026, 8, "2026年8月"),
            (2025, 11, "2025年11月"),
            (2026, 1, "2026年1月"),
            (2027, 12, "2027年12月"),
        ):
            with self.subTest(year=year, month=month):
                hook, _ = run_report(month=month, year=year)
                self.assertEqual(hook.get(PERIOD_CELL), expect)

    def test_no_duplicate_period_written(self):
        """同一页不应出现两处年月（fill_period=False 的意义）。

        FakeHook 的 find_cell 恒返回 (0,0)，所以通用页眉不会真的写；
        这里用事件数确认"C4 只被写了一次"，防止将来误改为
        fill_period=True 又叠加一次。
        """
        hook, _ = run_report()
        writes = [e for e in hook.events if e[1] == PERIOD_CELL]
        self.assertEqual(len(writes), 1,
                         f"C4 应只写一次，实际 {len(writes)} 次：{writes}")


class TestExportRevenue(unittest.TestCase):
    """需求第 2 条：行次2「其中：出口主营业务收入」（Excel 第 10 行）。"""

    def test_d10_is_february(self):
        """D10 必须是 2 月数据（iYPeriod=202602），不是 1 月也不是累计。

        【这条对应现场明确指示】
            需求原话：D10 是填列 2 月份的数，仅把 iYPeriod 指向 202602，
            其他取数条件相同，取 md 借方金额合计填入 D10。

        【为什么值得单独测 D10？】
            本实现是"一次查全年、按月分流"，与"逐月查 12 次"写法不同。
            若分流环节把 month 搞错（例如用了 month-1 或从 0 开始），
            1 月的数就会落到 D10，而整张表每格看起来都"有数字"。
            本条只给 2 月数据，D10 有值、C10 必须为空 —— 一眼可辨。
        """
        hook, _ = run_report([("600102", 2, 4321.0)])

        self.assertTrue(approx(hook.get("D10"), 4321.0),
                        f"D10（2月）应为 4321，实际 {hook.get('D10')}")
        self.assertIsNone(hook.get("C10"),
                          f"C10（1月）无数据时应空白，实际 {hook.get('C10')}")
        self.assertTrue(approx(hook.get("O10"), 4321.0),
                        f"O10 本年累计应为 4321，实际 {hook.get('O10')}")

    def test_month_to_cell_mapping_explicit(self):
        """逐月核对 12 个「月份 → 格位」对应关系（不靠 COL_OF_MONTH 自证）。

        【为什么把 12 条映射全部写死在测试里？】
            若测试也用 COL_OF_MONTH 推期望值，那本表写错时测试跟着一起错，
            等于没测。这里把"Excel 上第 10 行第几列是几月"这份**事实**
            独立写一遍，与代码里的映射互相印证。

        【对应关系（模板第 8 行表头已逐格确认）】
            C=1月 D=2月 E=3月 F=4月 G=5月 H=6月
            I=7月 J=8月 K=9月 L=10月 M=11月 N=12月
        """
        expect_map = {
            1: "C10", 2: "D10", 3: "E10", 4: "F10", 5: "G10", 6: "H10",
            7: "I10", 8: "J10", 9: "K10", 10: "L10", 11: "M10", 12: "N10",
        }

        for month, cell in expect_map.items():
            with self.subTest(month=month, cell=cell):
                # 只给该月一个特征值，其余月份无数据
                hook, _ = run_report([("600102", month, 1000.0 * month)])
                self.assertTrue(
                    approx(hook.get(cell), 1000.0 * month),
                    f"{month}月应写入 {cell}={1000.0 * month}，"
                    f"实际 {hook.get(cell)}")

    def test_twelve_months_written(self):
        """12 个月各自写入对应列，O 列 = 12 个月之和。"""
        rows = make_amount_rows({
            "600102": {m: m * 100.0 for m in range(1, 13)},
        })
        hook, _ = run_report(rows, month=12)

        for month, col in COL_OF_MONTH.items():
            expect = month * 100.0
            actual = hook.get(f"{col}10")
            self.assertTrue(
                approx(actual, expect),
                f"{col}10（{month}月）期望 {expect}，实际 {actual}")

        # 100+200+...+1200 = 7800
        self.assertTrue(approx(hook.get("O10"), 7800.0),
                        f"O10 应为 7800，实际 {hook.get('O10')}")

    def test_missing_months_stay_blank(self):
        """只有部分月份有数据时，其余月份格必须保持空白（不写 0）。

        【为什么这条重要？】
            "该月为 0" 与 "该月未取数" 在报表上是两回事：
            前者说明确实没有出口收入，后者说明需要人去查。
            若统一写 0，就再也分不清了。
        """
        rows = make_amount_rows({"600102": {3: 500.0, 7: 800.0}})
        hook, _ = run_report(rows)

        self.assertTrue(approx(hook.get("E10"), 500.0))   # 3月
        self.assertTrue(approx(hook.get("I10"), 800.0))   # 7月
        self.assertTrue(approx(hook.get("O10"), 1300.0))  # 累计

        for month, col in COL_OF_MONTH.items():
            if month in (3, 7):
                continue
            self.assertIsNone(
                hook.get(f"{col}10"),
                f"{col}10（{month}月）无数据时应保持空白，"
                f"实际 {hook.get(f'{col}10')}")

    def test_partial_year_holes_stay_blank(self):
        """选 8 月但只录到 6 月：7、8 月空白，O10 只累计 1~6 月。

        【这条对应现场提醒 20261103】
            操作录入日期通常是"凭证实际录完的月份"，但也有回查历史、
            漏结转、补录等情形。所以**绝不能假设 1~N 月连续有数**。

            场景：当前期间 = 8 月，但 U8 里 7、8 月还没做结转损益。
            正确表现：C10~H10 有值，I10、J10 空白，O10 = 1~6 月之和。
            错误表现：补 0 到 8 月（谎报"7、8 月无出口收入"），
                      或只填到 6 月就把 O10 算成"全年"。
        """
        rows = make_amount_rows({"600102": {m: 100.0 * m for m in range(1, 7)}})
        hook, _ = run_report(rows, month=8)

        # 1~6 月有值
        for month in range(1, 7):
            cell = f"{COL_OF_MONTH[month]}10"
            self.assertTrue(approx(hook.get(cell), 100.0 * month),
                            f"{cell}（{month}月）应为 {100.0 * month}")

        # 7、8 月必须空白（不能补 0）
        for month in (7, 8):
            cell = f"{COL_OF_MONTH[month]}10"
            self.assertIsNone(hook.get(cell),
                              f"{cell}（{month}月）无结转记录时应空白，"
                              f"实际 {hook.get(cell)}")

        # 累计 = 100+200+...+600 = 2100
        self.assertTrue(approx(hook.get("O10"), 2100.0),
                        f"O10 应为 2100（1~6月），实际 {hook.get('O10')}")

    def test_future_months_filled_not_truncated(self):
        """选 8 月但已录到 10 月：9、10 月照常填列，不擅自截断。

        【为什么不截断？】
            选 8 月却已有 9、10 月数据，说明该账套走在期间前面
            （提前录入或回查历史）。擅自截掉等于隐瞒已发生的事实；
            而"按实际数据填列 + 日志提示"既完整又不误导。

            【注意】这与"普通月份生成"不冲突：正常情况 9、10 月
            还没有凭证，自然取不到数，该情形不会触发。
        """
        rows = make_amount_rows({
            "600102": {m: 100.0 * m for m in range(1, 11)},   # 1~10 月都有
        })
        hook, _ = run_report(rows, month=8)

        # 9、10 月必须照常写入，不能被期间截断
        self.assertTrue(approx(hook.get("K10"), 900.0),
                        f"K10（9月）应填 900，实际 {hook.get('K10')}")
        self.assertTrue(approx(hook.get("L10"), 1000.0),
                        f"L10（10月）应填 1000，实际 {hook.get('L10')}")

        # 累计 = 1~10 月之和 = 5500
        self.assertTrue(approx(hook.get("O10"), 5500.0),
                        f"O10 应为 5500（1~10月），实际 {hook.get('O10')}")

    def test_non_consecutive_months(self):
        """月份不连续（1、2、5 月有数）时，只填这三个月。

        【为什么单测？】
            若有人把"填数循环"改成 `for m in range(1, max(ok)+1)`，
            就会把 3、4 月也写进去（值为 0 或 KeyError）。
            这条给不连续的月份，能立刻发现这类"按最大月补全"的误改。
        """
        rows = make_amount_rows({"600102": {1: 10.0, 2: 20.0, 5: 50.0}})
        hook, _ = run_report(rows, month=5)

        self.assertTrue(approx(hook.get("C10"), 10.0))   # 1月
        self.assertTrue(approx(hook.get("D10"), 20.0))   # 2月
        self.assertTrue(approx(hook.get("G10"), 50.0))   # 5月

        # 3、4 月必须空白
        for month in (3, 4):
            cell = f"{COL_OF_MONTH[month]}10"
            self.assertIsNone(hook.get(cell),
                              f"{cell}（{month}月）无数据应空白，"
                              f"实际 {hook.get(cell)}")

        self.assertTrue(approx(hook.get("O10"), 80.0))

    def test_no_data_writes_nothing(self):
        """完全取不到数时，整行（含 O10）保持空白。"""
        hook, _ = run_report([])
        for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
            self.assertIsNone(hook.get(f"{col}10"),
                              f"{col}10 无数据时不应写入")

    def test_fetch_failure_not_crash(self):
        """取数抛异常时不崩溃，且不写任何数。"""
        hook, _ = run_report([], fail=True)
        for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
            self.assertIsNone(hook.get(f"{col}10"))

    def test_sum_cross_month_records(self):
        """同月多条记录应累加（GROUP BY 之外的兜底）。

        【为什么要测？】
            SQL 已 GROUP BY ccode, iperiod，理论上每月每科目只有一行。
            但若将来取数函数被人改成"取明细"或 UNION 了别的来源，
            同月就会出现多行。报表层的累加逻辑必须能正确处理
            （这里用 += 而非 = 赋值，就是为了这个）。
        """
        rows = [
            ("600102", 5, 100.0),
            ("600102", 5, 250.0),   # 同月第二条
            ("600102", 6, 70.0),
        ]
        hook, _ = run_report(rows)

        self.assertTrue(approx(hook.get("G10"), 350.0),
                        f"5月应为 100+250=350，实际 {hook.get('G10')}")
        self.assertTrue(approx(hook.get("H10"), 70.0))
        self.assertTrue(approx(hook.get("O10"), 420.0))

    def test_invalid_period_skipped(self):
        """iperiod 超范围（如 U8 的第 13 调整期）应跳过，不写入非法列。

        【为什么？】
            C~N 只有 12 列。若把 13 期的数硬塞进来，
            COL_OF_MONTH 取值会 KeyError；若静默丢弃又不留痕迹，
            日后"某月数对不上"永远查不出原因。所以显式跳过并记警告。
        """
        rows = [
            ("600102", 13, 9999.0),
            ("600102", 2, 100.0),
        ]
        hook, _ = run_report(rows)

        self.assertTrue(approx(hook.get("D10"), 100.0))
        # 13 期不得混入累计
        self.assertTrue(approx(hook.get("O10"), 100.0),
                        f"O10 不应包含 13 期，实际 {hook.get('O10')}")

    def test_other_codes_ignored(self):
        """不属于本行的科目不得混入（防"取回全部科目"这类误改）。"""
        rows = make_amount_rows({
            "600102": {5: 100.0},
            "600101": {5: 7777.0},    # 内销，属行次3，不应进本行
            "6001":   {5: 8888.0},    # 一级科目，不应进本行
        })
        hook, _ = run_report(rows)

        self.assertTrue(approx(hook.get("G10"), 100.0),
                        f"G10 应只含 600102=100，实际 {hook.get('G10')}")

    def test_income_rows_share_one_query(self):
        """所有收入行应合并成一次查询（同方向、同摘要、同对方科目）。

        【为什么不是"总共只查 1 次"？】
            成本行取贷方 mc、收入行取借方 md，direction 不同 →
            分组键不同 → 必须分 2 条 SQL。
            若强行合并，SQL 里只能选一个 SUM 列，必有一组取错。
            所以"2 次查询"是正确性的代价，只多一次往返。

            本条只检查"收入组内部合并"：
            params 里应同时出现所有收入科目。
        """
        _, ctx = run_report([])

        # 找出所有 SUM(md) 的查询（收入组）
        income_calls = [c for c in ctx.fetch_calls if "SUM(md) AS amount" in c["sql"]]
        self.assertEqual(
            len(income_calls), 1,
            f"所有收入行应合并在同一个 SUM(md) 查询里，"
            f"实际有 {len(income_calls)} 条")

        params = income_calls[0]["params"]
        for code in ("600101", "600102", "6051"):
            self.assertIn(code, params, f"收入组查询应包含科目 {code}")

    def test_query_params_carry_filters(self):
        """SQL 参数必须带上年度、期间区间、科目、摘要、对方科目。

        【为什么逐项断言？】
            参数顺序错位不会报错，只会静默取到错误的数。
            这条把"参数里有什么、在第几位"钉死，顺序错位立刻暴露。

        【参数顺序（对应 SQL 里 ? 出现顺序）】
            ① iyear              目标年度，如 2026
            ② iYPeriod 下限      如 202601（1月）
            ③ iYPeriod 上限      如 202612（12月）
            ④… ccode 列表
            … cdigest 关键字（带 %）
            … ccode_equal 列表
        """
        _, ctx = run_report([])
        params = ctx.fetch_calls[0]["params"]

        self.assertEqual(params[0], 2026, "第1个参数应为年度 iyear")
        self.assertEqual(params[1], 202601, "第2个参数应为 iYPeriod 下限")
        self.assertEqual(params[2], 202612, "第3个参数应为 iYPeriod 上限")
        # 【为什么不再断言 params[3] 是某一个固定科目？】
        #   出口(600102)与内销(600101)口径完全相同，会被合并成同一组，
        #   SQL 里变成 `ccode IN ('600101','600102')`。
        #   科目在 params 里的顺序取决于 sorted()，且后续新增同行口径的行
        #   还会继续追加。所以这里只断言"两个科目都在"，
        #   不把具体下标钉死，避免每加一行就要改这条测试。
        self.assertIn("600101", params, "参数应包含内销科目 600101")
        self.assertIn("600102", params, "参数应包含出口科目 600102")

        # 其后是摘要关键字（带 %），最后是对方科目
        self.assertIn("%结转损益%", params)
        self.assertIn("%损益结转%", params)
        self.assertIn("4103", params)

    def test_sql_filters_by_yperiod(self):
        """SQL 必须按 iYPeriod 区间过滤（需求现场确认的写法）。

        【为什么要单独断言 SQL 文本？】
            参数对了不代表 SQL 用了它。若 WHERE 里漏写
            `iYPeriod BETWEEN ? AND ?`，params 里仍会有那两个值，
            但 SQL 不会使用 —— 数据库可能报"参数过多"，
            也可能静默按前 N 个参数执行，后者不报错但取数范围错。
            这条把 SQL 文本也钉住。
        """
        _, ctx = run_report([])
        sql = ctx.fetch_calls[0]["sql"]

        self.assertIn("iYPeriod BETWEEN ?", sql,
                      f"SQL 未按 iYPeriod 过滤，实际 SQL：{sql}")
        self.assertIn("iyear = ?", sql, "SQL 应保留 iyear 作为冗余保险")

        # 【为什么只断言 "ccode IN (" 而不写死占位符个数？】
        #   出口(600102)与内销(600101)合并成同一组后，这里是 IN (?, ?)。
        #   若写死 "ccode IN (?)"，每新增一个同口径科目就要改一次测试，
        #   而这条测试的目的是"确认有按科目过滤"，
        #   占位符个数由 test_both_codes_in_single_sql 单独覆盖。
        self.assertIn("ccode IN (", sql,
                      f"SQL 未按科目过滤，实际 SQL：{sql}")

    def test_formula_rows_never_written(self):
        """公式行（营业利润/利润总额/净利润）在任何输入下都不得被写入。"""
        rows = make_amount_rows({"600102": {m: 1000.0 for m in range(1, 13)}})
        hook, _ = run_report(rows, month=12)

        for row_no in FORMULA_ROWS:
            excel_row = EXCEL_ROW[row_no]
            for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
                self.assertIsNone(
                    hook.get(f"{col}{excel_row}"),
                    f"{col}{excel_row}（行次{row_no}）是公式行，程序不得写入")


class TestStructuralConstants(unittest.TestCase):
    """常量本身的正确性（坐标写错会静默填到别的项目上）。"""

    def test_month_columns(self):
        """12 个月列应为 C~N，且顺序连续。"""
        self.assertEqual(list(COL_OF_MONTH.keys()), list(range(1, 13)))
        self.assertEqual(COL_OF_MONTH[1], "C")
        self.assertEqual(COL_OF_MONTH[12], "N")
        self.assertEqual(COL_YEAR_TOTAL, "O")

    def test_excel_row_mapping(self):
        """行次 → Excel 行号，须与模板逐格核对结果一致。

        【为什么把关键值写死在测试里？】
            这张对照表是"抄模板"得来的，一旦有人凭记忆改动，
            整张表会整体错行，而报表上每一格看起来都"正常"。
            这里固定几个锚点（首行/汇总行/末行）立刻能发现偏移。
        """
        self.assertEqual(EXCEL_ROW[1], 9, "行次1 营业收入应在 Excel 第9行")
        self.assertEqual(EXCEL_ROW[5], 13, "行次5 营业成本应在 Excel 第13行")
        self.assertEqual(EXCEL_ROW[22], 30, "行次22 营业利润应在 Excel 第30行")
        self.assertEqual(EXCEL_ROW[25], 33, "行次25 利润总额应在 Excel 第33行")
        self.assertEqual(EXCEL_ROW[27], 35, "行次27 净利润应在 Excel 第35行")
        self.assertEqual(EXCEL_ROW[26], 34, "行次26 所得税费用应在 Excel 第34行")

    def test_formula_rows(self):
        """带模板公式的汇总行必须是 22/25/27。"""
        self.assertEqual(set(FORMULA_ROWS), {22, 25, 27})


class TestRowSpecs(unittest.TestCase):
    """行项目配置的结构护栏（随 ROW_SPECS 增长自动生效）。"""

    def test_formula_rows_not_in_specs(self):
        """公式行不得出现在 ROW_SPECS 中。

        【为什么这条最重要？】
            模板里营业利润/利润总额/净利润三行自带公式。
            若程序往里写数值，会覆盖公式 —— 下个月再生成时
            公式已经没了，只剩上个月算出的死数，且不报任何错。
            这类错误在报表上完全看不出来。
        """
        formula_excel_rows = {EXCEL_ROW[n] for n in FORMULA_ROWS}
        for spec in ROW_SPECS:
            self.assertNotIn(
                spec.excel_row, formula_excel_rows,
                f"「{spec.label}」的 Excel 行 {spec.excel_row} 属公式行，"
                f"程序不应写入（会覆盖模板公式）")

    def test_no_duplicate_excel_rows(self):
        """不同行项目不得占用同一 Excel 行。"""
        seen: dict[int, str] = {}
        for spec in ROW_SPECS:
            self.assertNotIn(
                spec.excel_row, seen,
                f"Excel 行 {spec.excel_row} 被「{seen.get(spec.excel_row)}」"
                f"和「{spec.key}」同时占用")
            seen[spec.excel_row] = spec.key

    def test_rows_exist_in_excel_row_map(self):
        """spec 的 excel_row 必须来自 EXCEL_ROW 的值（防止手写行号）。"""
        allowed = set(EXCEL_ROW.values())
        for spec in ROW_SPECS:
            self.assertIn(
                spec.excel_row, allowed,
                f"「{spec.key}」的 excel_row={spec.excel_row} "
                f"不在 EXCEL_ROW 对照表中，可能手写错了行号")

    def test_direction_valid(self):
        """direction 只能是 debit / credit / signed。

        【为什么允许 signed？】
            有些科目（如 6117 其他收益）在结转时借贷两侧都可能出现，
            必须同时取 SUM(md) 与 SUM(mc) 再按净额折算，
            不能用单一方向。详见 ProfitRow 的 direction 说明。
        """
        for spec in ROW_SPECS:
            self.assertIn(
                spec.direction, ("credit", "debit", "signed"),
                f"「{spec.key}」的 direction={spec.direction!r} 非法")

    def test_codes_are_strings(self):
        """科目编码必须是字符串。"""
        for spec in ROW_SPECS:
            for c in (*spec.codes, *spec.exclude_codes):
                self.assertIsInstance(
                    c, str,
                    f"「{spec.key}」的科目编码 {c!r} 不是字符串")

    def test_codes_not_empty(self):
        """已配置的行必须有科目（否则填不出数，且原因不明）。"""
        for spec in ROW_SPECS:
            self.assertTrue(
                spec.codes,
                f"「{spec.key}」的 codes 为空 —— 若该行本期不取数，"
                f"应登记为待定项，而不是留一条空配置")

    def test_cell_map_matches_specs(self):
        """ROW_SPECS 涉及的格子都应登记进 cell_map。

        【为什么？】
            cell_map 决定父类 verify() 读回哪些格做勾稽核对。
            新增行时若忘了登记，该行数据不进核对清单，
            日后"勾稽通过"可能是假象（漏检）。
        """
        cm_values = set(LK02Report.cell_map.values())
        for spec in ROW_SPECS:
            cells = [f"{col}{spec.excel_row}" for col in COL_OF_MONTH.values()]
            if spec.fill_total:
                cells.append(f"{COL_YEAR_TOTAL}{spec.excel_row}")
            for cell in cells:
                self.assertIn(
                    cell, cm_values,
                    f"「{spec.key}」的 {cell} 未登记 cell_map")


class TestDomesticRevenue(unittest.TestCase):
    """需求第 3 条：行次3「其中：内销主营业务收入」（Excel 第 11 行）。

    口径与第 2 条完全相同，仅科目编码 600101（内销）对 600102（出口）。
    """

    def test_d11_is_february(self):
        """D11 必须是 2 月内销数（600101），不能混入出口(600102)。"""
        hook, _ = run_report([("600101", 2, 8765.0)])

        self.assertTrue(approx(hook.get("D11"), 8765.0),
                        f"D11（2月内销）应为 8765，实际 {hook.get('D11')}")
        self.assertIsNone(hook.get("C11"),
                          f"C11 无数据时应空白，实际 {hook.get('C11')}")
        self.assertTrue(approx(hook.get("O11"), 8765.0))

    def test_twelve_months_written(self):
        """12 个月各自写入对应列，O11 = 12 个月之和。"""
        rows = make_amount_rows({
            "600101": {m: m * 10.0 for m in range(1, 13)},
        })
        hook, _ = run_report(rows, month=12)

        for month, col in COL_OF_MONTH.items():
            expect = month * 10.0
            actual = hook.get(f"{col}11")
            self.assertTrue(
                approx(actual, expect),
                f"{col}11（{month}月）期望 {expect}，实际 {actual}")

        # 10+20+...+120 = 780
        self.assertTrue(approx(hook.get("O11"), 780.0),
                        f"O11 应为 780，实际 {hook.get('O11')}")

    def test_export_and_domestic_not_crossed(self):
        """出口(600102)与内销(600101)各归各行，绝不串用。

        【为什么这是本需求最关键的测试？】
            两行口径完全相同，只有科目编码差一位（600102 / 600101）、
            Excel 行号差一行（10 / 11）。若分流环节把科目-行映射搞反，
            或两行共用同一个累加器，就会出现：
                · 出口行的数跑到内销行
                · 两行显示同一个数
                · 某一行翻倍
            而报表上每格都"有数字"，肉眼看不出任何异常。

            本条给两个明显不同的金额，任一处串位立刻暴露。
        """
        rows = make_amount_rows({
            "600102": {m: 1000.0 for m in range(1, 7)},   # 出口 1~6 月各 1000
            "600101": {m: 7.0 for m in range(1, 7)},       # 内销 1~6 月各 7
        })
        hook, _ = run_report(rows, month=6)

        for month, col in COL_OF_MONTH.items():
            if month > 6:
                continue
            with self.subTest(month=month):
                # 出口行（第10行）
                self.assertTrue(
                    approx(hook.get(f"{col}10"), 1000.0),
                    f"{col}10（出口 {month}月）应为 1000，"
                    f"实际 {hook.get(f'{col}10')}")
                # 内销行（第11行）
                self.assertTrue(
                    approx(hook.get(f"{col}11"), 7.0),
                    f"{col}11（内销 {month}月）应为 7，"
                    f"实际 {hook.get(f'{col}11')}")

        self.assertTrue(approx(hook.get("O10"), 6000.0),
                        f"O10 出口累计应为 6000，实际 {hook.get('O10')}")
        self.assertTrue(approx(hook.get("O11"), 42.0),
                        f"O11 内销累计应为 42，实际 {hook.get('O11')}")

    def test_domestic_blank_when_only_export_has_data(self):
        """只有出口数据时，内销整行必须空白（不能凭空填数或写 0）。"""
        rows = make_amount_rows({"600102": {m: 500.0 for m in range(1, 4)}})
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("O10"), 1500.0))
        for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
            self.assertIsNone(hook.get(f"{col}11"),
                              f"{col}11 内销无数据时应空白，"
                              f"实际 {hook.get(f'{col}11')}")

    def test_single_query_covers_both_rows(self):
        """两行口径相同，必须合并成一次查询（不能各查一次）。

        【为什么要锁这条？】
            若将来有人给其中一行单独加了筛选条件（例如只对内销
            过滤某个摘要），分组键就会不同 → 变成 2 次查询。
            这可能是有意为之，但也可能是改错了。本条失败时
            会提示"两行口径已不再相同，请确认是否有意为之"。
        """
        _, ctx = run_report([])

        # 只看收入组（SUM(md)），因为成本行是另一组
        income_calls = [c for c in ctx.fetch_calls if "SUM(md) AS amount" in c["sql"]]
        self.assertEqual(
            len(income_calls), 1,
            f"出口与内销口径相同，应在同一个收入组查询里，"
            f"实际收入组有 {len(income_calls)} 条查询")

        params = income_calls[0]["params"]
        self.assertIn("600101", params)
        self.assertIn("600102", params)

    def test_both_codes_in_single_sql(self):
        """合并后的 SQL 应一次包含两个科目（IN 列表）。"""
        _, ctx = run_report([])
        params = ctx.fetch_calls[0]["params"]

        self.assertIn("600101", params)
        self.assertIn("600102", params)

    def test_formula_rows_never_written_with_domestic(self):
        """内销有数据时，公式行仍不得被写入。"""
        rows = make_amount_rows({"600101": {m: 1000.0 for m in range(1, 13)}})
        hook, _ = run_report(rows, month=12)

        for row_no in FORMULA_ROWS:
            excel_row = EXCEL_ROW[row_no]
            for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
                self.assertIsNone(
                    hook.get(f"{col}{excel_row}"),
                    f"{col}{excel_row}（行次{row_no}）是公式行，不得写入")


class TestOtherBusinessRevenue(unittest.TestCase):
    """需求第 4 条：行次4「其中：其他业务收入」（Excel 第 12 行）。

    口径与第 2、3 条完全相同，仅科目编码 6051（一级码，四位）。
    """

    def test_c12_is_january(self):
        """C12 必须是 1 月其他业务收入（6051）。"""
        hook, _ = run_report([("6051", 1, 3210.0)])

        self.assertTrue(approx(hook.get("C12"), 3210.0),
                        f"C12（1月）应为 3210，实际 {hook.get('C12')}")
        self.assertTrue(approx(hook.get("O12"), 3210.0))
        self.assertIsNone(hook.get("D12"),
                          f"D12 无数据时应空白，实际 {hook.get('D12')}")

    def test_twelve_months_written(self):
        """12 个月各自写入对应列，O12 = 12 个月之和。"""
        rows = make_amount_rows({"6051": {m: m * 3.0 for m in range(1, 13)}})
        hook, _ = run_report(rows, month=12)

        for month, col in COL_OF_MONTH.items():
            expect = month * 3.0
            actual = hook.get(f"{col}12")
            self.assertTrue(
                approx(actual, expect),
                f"{col}12（{month}月）期望 {expect}，实际 {actual}")

        # 3+6+...+36 = 234
        self.assertTrue(approx(hook.get("O12"), 234.0),
                        f"O12 应为 234，实际 {hook.get('O12')}")

    def test_three_rows_not_crossed(self):
        """出口(10行)/内销(11行)/其他业务(12行) 三行互不串用。

        【为什么这条比上一轮的两行互斥更重要？】
            现在有三行口径完全相同，只有科目编码与 Excel 行号不同：
                600102 → 第 10 行
                600101 → 第 11 行
                6051   → 第 12 行
            一旦分流映射写错（例如字典 key 重复、或行号抄错），
            可能出现"某两行显示同一个数"或"某行空着"，
            而报表上每格都"有数字"，完全看不出异常。

            本条给三个数量级明显不同的金额，任一处串位立刻暴露。
        """
        rows = make_amount_rows({
            "600102": {3: 1000.0},   # 出口
            "600101": {3: 100.0},    # 内销
            "6051":   {3: 1.0},      # 其他业务
        })
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("E10"), 1000.0),
                        f"E10（出口）应为 1000，实际 {hook.get('E10')}")
        self.assertTrue(approx(hook.get("E11"), 100.0),
                        f"E11（内销）应为 100，实际 {hook.get('E11')}")
        self.assertTrue(approx(hook.get("E12"), 1.0),
                        f"E12（其他业务）应为 1，实际 {hook.get('E12')}")

    def test_other_business_blank_when_others_have_data(self):
        """只有出口/内销数据时，第 12 行必须整行空白。"""
        rows = make_amount_rows({
            "600102": {1: 500.0},
            "600101": {1: 300.0},
        })
        hook, _ = run_report(rows, month=1)

        for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
            self.assertIsNone(hook.get(f"{col}12"),
                              f"{col}12 其他业务收入无数据时应空白，"
                              f"实际 {hook.get(f'{col}12')}")

    def test_exact_match_not_prefix(self):
        """必须精确匹配 6051，不得把下级科目 605101 等一并算入。

        【口径已现场确认（20261103）：精确匹配是正确的】
            理由有二：
              ① 本账套的 6051 没有下级科目（无 605101、605102…）；
              ② 即使将来设了下级，U8 会在 6051 本身记一行汇总，
                 下级科目的金额已包含在 6051 这条记录里。

            所以若改成前缀匹配（LIKE '6051%'）或把下级码追加进 codes，
            就会把同一笔金额**算两遍**：
                6051 本身已含下级合计
              + 605101、605102 再各算一次
              = 金额翻倍，而报表上只显示一个看似合理的数，肉眼无法发现。

            本条给一级码和两个下级码不同金额：
              · 正确实现：C12 = 3（只认 6051）
              · 误用前缀匹配：C12 = 3+30+300 = 333（错，翻百倍）
              · 误追加下级码：C12 = 333（同样错）
        """
        rows = [
            ("6051", 1, 3.0),        # 一级科目（应入）
            ("605101", 1, 30.0),     # 下级（不应入）
            ("605102", 1, 300.0),    # 下级（不应入）
        ]
        hook, _ = run_report(rows, month=1)

        self.assertTrue(approx(hook.get("C12"), 3.0),
                        f"C12 应只含 6051=3（不含下级），"
                        f"实际 {hook.get('C12')}")

    def test_three_codes_in_single_sql(self):
        """三个收入科目应合并进同一次查询（收入组）的 IN 列表。

        【为什么限定在 SUM(md) 的查询里找？】
            成本行（640102）取 SUM(mc)，是另一组。
            本条只验证"三个收入科目合并在同一组"。
        """
        _, ctx = run_report([])

        income_calls = [c for c in ctx.fetch_calls if "SUM(md) AS amount" in c["sql"]]
        self.assertEqual(
            len(income_calls), 1,
            f"三个收入科目应合并为 1 条 SUM(md) 查询，"
            f"实际 {len(income_calls)} 条")

        params = income_calls[0]["params"]
        for code in ("600102", "600101", "6051"):
            self.assertIn(code, params, f"SQL 参数缺少科目 {code}")

    def test_formula_rows_never_written_with_other_business(self):
        """其他业务收入有数据时，公式行仍不得被写入。"""
        rows = make_amount_rows({"6051": {m: 1000.0 for m in range(1, 13)}})
        hook, _ = run_report(rows, month=12)

        for row_no in FORMULA_ROWS:
            excel_row = EXCEL_ROW[row_no]
            for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
                self.assertIsNone(
                    hook.get(f"{col}{excel_row}"),
                    f"{col}{excel_row}（行次{row_no}）是公式行，不得写入")


class TestExportCost(unittest.TestCase):
    """需求第 5 条：行次6「其中：出口主营业务成本」（Excel 第 14 行）。

    【本类的核心价值：方向不同】
        前四条收入行取借方 md，本行是成本行取贷方 mc。
        结转损益时收入与成本分录方向相反（见模块头部口径③）。
        若照抄上一行的 "debit"，本行会永远空白且不报错。
    """

    def test_c14_is_january(self):
        """C14 必须是 1 月出口成本（640102）。"""
        hook, _ = run_report([("640102", 1, 6789.0)])

        self.assertTrue(approx(hook.get("C14"), 6789.0),
                        f"C14（1月）应为 6789，实际 {hook.get('C14')}")
        self.assertTrue(approx(hook.get("O14"), 6789.0))
        self.assertIsNone(hook.get("D14"),
                          f"D14 无数据时应空白，实际 {hook.get('D14')}")

    def test_twelve_months_written(self):
        """12 个月各自写入对应列，O14 = 12 个月之和。"""
        rows = make_amount_rows({"640102": {m: m * 5.0 for m in range(1, 13)}})
        hook, _ = run_report(rows, month=12)

        for month, col in COL_OF_MONTH.items():
            expect = month * 5.0
            actual = hook.get(f"{col}14")
            self.assertTrue(
                approx(actual, expect),
                f"{col}14（{month}月）期望 {expect}，实际 {actual}")

        # 5+10+...+60 = 390
        self.assertTrue(approx(hook.get("O14"), 390.0),
                        f"O14 应为 390，实际 {hook.get('O14')}")

    def test_income_and_cost_not_crossed(self):
        """收入行(第10行)与成本行(第14行)互不串用。

        【为什么必须测？】
            本期新增了"成本行"，与收入行相比：
                · 科目不同（640102 vs 600102）
                · **方向不同**（credit vs debit）
                · Excel 行不同（14 vs 10）

            若方向写错（成本行也用 debit），取数结果为空 ——
            本条会因 E14 为 None 而失败，直接暴露。
            若行号写错（例如把成本写到第 10 行），本条同样失败。
        """
        rows = make_amount_rows({
            "600102": {3: 1000.0},    # 收入（debit 侧）
            "640102": {3: 600.0},     # 成本（credit 侧）
        })
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("E10"), 1000.0),
                        f"E10（出口收入）应为 1000，实际 {hook.get('E10')}")
        self.assertTrue(approx(hook.get("E14"), 600.0),
                        f"E14（出口成本）应为 600，实际 {hook.get('E14')}")

    def test_cost_row_uses_credit_direction(self):
        """成本行的 direction 必须是 credit（不能照抄收入行的 debit）。

        【为什么直接断言配置而不是看运行结果？】
            运行结果为空也可能因为"该月无数据"，
            只看结果无法区分"方向配错"与"确实没数据"。
            这条直接检查配置，把「640102 必须 credit」钉死。
        """
        spec = next(s for s in ROW_SPECS if s.key == "出口主营业务成本")

        self.assertEqual(spec.direction, "credit",
                         f"出口主营业务成本应取贷方 mc，"
                         f"实际 direction={spec.direction!r}")
        self.assertEqual(spec.codes, ("640102",))
        self.assertEqual(spec.excel_row, EXCEL_ROW[6], "应写第 14 行")

    def test_income_rows_use_debit_direction(self):
        """三条收入行的 direction 必须是 debit（与成本行相反）。

        【对称检查】
            与上一条配对：防止有人"统一改成 credit"或"统一改成 debit"。
            两组方向必须相反，这是由结转损益的分录结构决定的，
            不是代码风格偏好。
        """
        income_keys = ("出口主营业务收入", "内销主营业务收入", "其他业务收入")
        for key in income_keys:
            spec = next(s for s in ROW_SPECS if s.key == key)
            self.assertEqual(
                spec.direction, "debit",
                f"「{key}」应取借方 md，实际 direction={spec.direction!r}")

    def test_cost_row_is_optional(self):
        """成本行标注 optional=True（640102 现场暂未启用）。"""
        spec = next(s for s in ROW_SPECS if s.key == "出口主营业务成本")

        self.assertTrue(
            spec.optional,
            "640102 现场标注为暂未启用，应设 optional=True，"
            "让'无数据'走 INFO 而非 WARNING")

    def test_optional_row_blank_when_no_data(self):
        """640102 无数据时，第 14 行整行空白（不写 0）。"""
        rows = make_amount_rows({"600102": {1: 500.0}})   # 只有收入，无成本
        hook, _ = run_report(rows, month=1)

        for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
            self.assertIsNone(hook.get(f"{col}14"),
                              f"{col}14 无数据时应空白，"
                              f"实际 {hook.get(f'{col}14')}")

    def test_optional_row_fills_once_data_appears(self):
        """一旦 640102 开始有数，按正常逻辑填列（optional 不阻止填数）。"""
        rows = make_amount_rows({"640102": {1: 100.0, 2: 200.0}})
        hook, _ = run_report(rows, month=2)

        self.assertTrue(approx(hook.get("C14"), 100.0))
        self.assertTrue(approx(hook.get("D14"), 200.0))
        self.assertTrue(approx(hook.get("O14"), 300.0))

    def test_cost_query_uses_mc(self):
        """SQL 必须对成本行使用 SUM(mc)，不能是 SUM(md)。"""
        _, ctx = run_report([])

        # 成本行与收入行方向不同 → 分组键不同 → 两次查询
        sqls = [c["sql"] for c in ctx.fetch_calls]

        self.assertTrue(
            any("SUM(mc) AS amount" in s for s in sqls),
            f"应有一条 SQL 使用 SUM(mc)，实际：{sqls}")
        self.assertTrue(
            any("SUM(md) AS amount" in s for s in sqls),
            f"应有一条 SQL 使用 SUM(md)，实际：{sqls}")

    def test_direction_split_creates_two_queries(self):
        """收入行取借方、成本行取贷方，但仍只发 1 次查询。

        【这条测试的期望值曾经是错的，记录一下】
            我最初认为"方向不同 → 必须分 2 次查询"，理由是
            "一条 SQL 只能选一个 SUM 列，合并必有一组取错"。

            这个判断是错的。正确做法是**一次同时 SUM(md) 与 SUM(mc)**
            返回两列，报表层按各行自己的 direction 选用对应列：
                · SQL 结构统一 → 可合并成 1 次查询；
                · 不用的那一列忽略即可，不会取错。

            收益：少一次数据库往返；且所有行取自同一快照，
            不会出现"收入与成本来自不同时刻"的错位。
            代价：多返回一列（聚合后行数本就不多，可忽略）。
        """
        _, ctx = run_report([])

        # 【为什么现在是 3 次？】
        #   需求第 11 条新增 6117 其他收益，direction="signed"，
        #   它走独立取数函数（同时取 md 与 mc），与收入组/成本组都不同
        #   → 共 3 组查询：debit / credit / signed。
        self.assertEqual(
            len(ctx.fetch_calls), 3,
            f"收入组(debit)、成本组(credit)、方向不定组(signed) 各 1 次，"
            f"共 3 次查询，实际 {len(ctx.fetch_calls)} 次")

    def test_excel_row_mapping_for_cost(self):
        """出口主营业务成本对应 Excel 第 14 行（行次 6）。

        【为什么单独断言？】
            EXCEL_ROW 里第 6 行对应 14：
                行次 5 = 13（营业成本）
                行次 6 = 14（出口主营业务成本）
                行次 7 = 15（内销主营业务成本）
            容易误写成 13 或 15；写错会把成本数填到相邻行上。
        """
        self.assertEqual(EXCEL_ROW[6], 14)


class TestDomesticCost(unittest.TestCase):
    """需求第 6 条：行次7「其中：内销主营业务成本」（Excel 第 15 行）。

    口径与第 5 条（出口成本）相同，科目 640101，方向 credit（取贷方 mc）。
    新增降级机制：640101 无数据时改用一级科目 6401。
    """

    def test_c15_is_january(self):
        """C15 必须是 1 月内销成本（640101）。"""
        hook, _ = run_report([("640101", 1, 2468.0)], month=1)

        self.assertTrue(approx(hook.get("C15"), 2468.0),
                        f"C15（1月）应为 2468，实际 {hook.get(chr(39)+chr(67)+chr(49)+chr(53)+chr(39))}")
        self.assertTrue(approx(hook.get("O15"), 2468.0))
        self.assertIsNone(hook.get("D15"))

    def test_twelve_months_written(self):
        """12 个月各自写入对应列，O15 = 12 个月之和。"""
        rows = make_amount_rows({"640101": {m: m * 2.0 for m in range(1, 13)}})
        hook, _ = run_report(rows, month=12)

        for month, col in COL_OF_MONTH.items():
            self.assertTrue(approx(hook.get(f"{col}15"), month * 2.0))
        self.assertTrue(approx(hook.get("O15"), 156.0))

    def test_direction_is_credit(self):
        """内销成本行的 direction 必须是 credit（与收入行相反）。"""
        spec = next(s for s in ROW_SPECS if s.key == "内销主营业务成本")

        self.assertEqual(spec.direction, "credit")
        self.assertEqual(spec.codes, ("640101",))
        self.assertEqual(spec.excel_row, EXCEL_ROW[7], "应写第 15 行")

    def test_fallback_codes_configured(self):
        """备用科目必须配成 6401（现场要求）。"""
        spec = next(s for s in ROW_SPECS if s.key == "内销主营业务成本")

        self.assertEqual(spec.fallback_codes, ("6401",),
                         "640101 不存在时应降级用一级科目 6401")

    def test_use_fallback_when_main_absent(self):
        """640101 整年无数据时，改用 6401 填列第 15 行。"""
        hook, _ = run_report([("6401", 1, 888.0), ("6401", 2, 999.0)], month=2)

        self.assertTrue(approx(hook.get("C15"), 888.0),
                        f"降级后 C15 应为 6401 的 888，实际 {hook.get(chr(39)+chr(67)+chr(49)+chr(53)+chr(39))}")
        self.assertTrue(approx(hook.get("D15"), 999.0))
        self.assertTrue(approx(hook.get("O15"), 1887.0))

    def test_no_fallback_when_main_has_data(self):
        """640101 有数据时绝不掺入 6401（防重复计算）。"""
        hook, _ = run_report([("640101", 1, 555.0), ("6401", 1, 888.0)], month=1)

        self.assertTrue(approx(hook.get("C15"), 555.0),
                        f"应用 640101 的 555，实际 {hook.get(chr(39)+chr(67)+chr(49)+chr(53)+chr(39))}")

    def test_blank_when_both_absent(self):
        """主备科目都没有数据时，整行空白。"""
        hook, _ = run_report([], month=1)

        for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
            self.assertIsNone(hook.get(f"{col}15"))

    def test_not_crossed_with_export_cost(self):
        """内销成本(15行)与出口成本(14行)互不串用。"""
        rows = make_amount_rows({
            "640102": {3: 1000.0},
            "640101": {3: 100.0},
        })
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("E14"), 1000.0))
        self.assertTrue(approx(hook.get("E15"), 100.0))

    def test_excel_row_mapping(self):
        """内销主营业务成本对应 Excel 第 15 行（行次 7）。"""
        self.assertEqual(EXCEL_ROW[7], 15)

class TestOtherBusinessCost(unittest.TestCase):
    """需求第 7 条：行次8「其中：其他业务成本」（Excel 第 16 行）。

    口径与第 5、6 条相同，科目 6402，方向 credit（取贷方 mc）。
    """

    def test_c16_is_january(self):
        """C16 必须是 1 月其他业务成本（6402）。"""
        hook, _ = run_report([("6402", 1, 1357.0)], month=1)

        self.assertTrue(approx(hook.get("C16"), 1357.0),
                        f"C16（1月）应为 1357，实际 {hook.get('C16')}")
        self.assertTrue(approx(hook.get("O16"), 1357.0))
        self.assertIsNone(hook.get("D16"))

    def test_twelve_months_written(self):
        """12 个月各自写入对应列，O16 = 12 个月之和。"""
        rows = make_amount_rows({"6402": {m: m * 4.0 for m in range(1, 13)}})
        hook, _ = run_report(rows, month=12)

        for month, col in COL_OF_MONTH.items():
            self.assertTrue(approx(hook.get(f"{col}16"), month * 4.0))
        self.assertTrue(approx(hook.get("O16"), 312.0))

    def test_direction_is_credit(self):
        """其他业务成本行的 direction 必须是 credit。"""
        spec = next(s for s in ROW_SPECS if s.key == "其他业务成本")

        self.assertEqual(spec.direction, "credit")
        self.assertEqual(spec.codes, ("6402",))
        self.assertEqual(spec.excel_row, EXCEL_ROW[8], "应写第 16 行")

    def test_no_fallback_configured(self):
        """其他业务成本未配备用科目（现场未要求）。"""
        spec = next(s for s in ROW_SPECS if s.key == "其他业务成本")

        self.assertEqual(spec.fallback_codes, ())

    def test_not_crossed_with_other_cost_rows(self):
        """三行成本（14/15/16）互不串用。"""
        rows = make_amount_rows({
            "640102": {3: 1000.0},
            "640101": {3: 100.0},
            "6402":   {3: 1.0},
        })
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("E14"), 1000.0),
                        f"E14（出口成本）应为 1000，实际 {hook.get('E14')}")
        self.assertTrue(approx(hook.get("E15"), 100.0),
                        f"E15（内销成本）应为 100，实际 {hook.get('E15')}")
        self.assertTrue(approx(hook.get("E16"), 1.0),
                        f"E16（其他业务成本）应为 1，实际 {hook.get('E16')}")

    def test_blank_when_no_data(self):
        """6402 无数据时整行空白（不写 0）。"""
        hook, _ = run_report([], month=1)

        for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
            self.assertIsNone(hook.get(f"{col}16"),
                              f"{col}16 无数据时应空白，"
                              f"实际 {hook.get(f'{col}16')}")

    def test_excel_row_mapping(self):
        """其他业务成本对应 Excel 第 16 行（行次 8）。"""
        self.assertEqual(EXCEL_ROW[8], 16)


class TestTaxAndSellingExpense(unittest.TestCase):
    """需求第 8 条：两行成本费用项目（口径与其他业务成本完全相同）。

        ① 行次 9「其中：税金及附加」科目 6403 → C17:N17 + O17
        ② 行次 10「其中：销售费用」  科目 6601 → C18:N18 + O18
    """

    def test_c17_is_january(self):
        """C17 必须是 1 月税金及附加（6403）。"""
        hook, _ = run_report([("6403", 1, 4321.0)], month=1)

        self.assertTrue(approx(hook.get("C17"), 4321.0),
                        f"C17（1月）应为 4321，实际 {hook.get('C17')}")
        self.assertTrue(approx(hook.get("O17"), 4321.0))
        self.assertIsNone(hook.get("D17"))

    def test_c18_is_january(self):
        """C18 必须是 1 月销售费用（6601）。"""
        hook, _ = run_report([("6601", 1, 5678.0)], month=1)

        self.assertTrue(approx(hook.get("C18"), 5678.0),
                        f"C18（1月）应为 5678，实际 {hook.get('C18')}")
        self.assertTrue(approx(hook.get("O18"), 5678.0))
        self.assertIsNone(hook.get("D18"))

    def test_twelve_months_written(self):
        """两行各自 12 个月全部写入，O 列累计正确。"""
        rows = make_amount_rows({
            "6403": {m: m * 10.0 for m in range(1, 13)},
            "6601": {m: m * 20.0 for m in range(1, 13)},
        })
        hook, _ = run_report(rows, month=12)

        for month, col in COL_OF_MONTH.items():
            self.assertTrue(
                approx(hook.get(f"{col}17"), month * 10.0),
                f"{col}17（{month}月）应为 {month * 10.0}")
            self.assertTrue(
                approx(hook.get(f"{col}18"), month * 20.0),
                f"{col}18（{month}月）应为 {month * 20.0}")

        self.assertTrue(approx(hook.get("O17"), 780.0))
        self.assertTrue(approx(hook.get("O18"), 1560.0))

    def test_direction_is_credit(self):
        """两行方向都必须是 credit（成本费用类取贷方 mc）。"""
        for key, code, excel_row in (
            ("税金及附加", "6403", 17),
            ("销售费用", "6601", 18),
        ):
            with self.subTest(key=key):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.direction, "credit")
                self.assertEqual(spec.codes, (code,))
                self.assertEqual(spec.excel_row, excel_row)

    def test_two_rows_not_crossed(self):
        """税金及附加(17)与销售费用(18)互不串用。"""
        rows = make_amount_rows({
            "6403": {3: 1000.0},
            "6601": {3: 1.0},
        })
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("E17"), 1000.0),
                        f"E17（税金及附加）应为 1000，实际 {hook.get('E17')}")
        self.assertTrue(approx(hook.get("E18"), 1.0),
                        f"E18（销售费用）应为 1，实际 {hook.get('E18')}")

    def test_blank_when_no_data(self):
        """两行都无数据时整行空白（不写 0）。"""
        hook, _ = run_report([], month=1)

        for row in (17, 18):
            for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
                self.assertIsNone(hook.get(f"{col}{row}"),
                                  f"{col}{row} 无数据时应空白，"
                                  f"实际 {hook.get(f'{col}{row}')}")

    def test_no_fallback_configured(self):
        """两行都未配备用科目（现场未要求）。"""
        for key in ("税金及附加", "销售费用"):
            spec = next(s for s in ROW_SPECS if s.key == key)
            self.assertEqual(spec.fallback_codes, ())

    def test_excel_row_mapping(self):
        """行次 9 → Excel 第 17 行；行次 10 → Excel 第 18 行。"""
        self.assertEqual(EXCEL_ROW[9], 17)
        self.assertEqual(EXCEL_ROW[10], 18)

    def test_cost_group_merged_in_one_query(self):
        """所有成本费用行合并在同一次 SUM(mc) 查询。"""
        _, ctx = run_report([])

        cost_calls = [c for c in ctx.fetch_calls if "SUM(mc) AS amount" in c["sql"]]
        self.assertEqual(
            len(cost_calls), 1,
            f"所有成本费用行应合并为 1 条 SUM(mc) 查询，"
            f"实际 {len(cost_calls)} 条")

        params = cost_calls[0]["params"]
        for code in ("640101", "640102", "6402", "6403", "6601"):
            self.assertIn(code, params,
                          f"成本组查询应包含科目 {code}")


class TestAdminAndRndExpense(unittest.TestCase):
    """需求第 9 条：管理费用与研发费用（首对"净额 + 单列"组合）。

        ① 行次 11「其中：管理费用」= 6602 － 660253 → C19:N19 + O19
        ② 行次 12「其中：研发费用」= 660253        → C20:N20 + O20

    【为什么要一起测？】
        这两行共享同一个科目 660253：第19行作**减项**、第20行作**加项**。
        单独测任一行都看不出"共享科目是否被重复或漏算"，
        必须成对验证「第19行 + 第20行 = 6602 全额」这个恒等式。
    """

    def test_admin_is_net_of_rnd(self):
        """C19 = 6602 － 660253（净额列示）。"""
        hook, _ = run_report([("6602", 1, 1000.0), ("660253", 1, 300.0)],
                             month=1)

        self.assertTrue(approx(hook.get("C19"), 700.0),
                        f"C19（管理费用净额）应为 700，实际 {hook.get('C19')}")

    def test_rnd_is_detail_code(self):
        """C20 = 660253（取明细科目本身）。"""
        hook, _ = run_report([("6602", 1, 1000.0), ("660253", 1, 300.0)],
                             month=1)

        self.assertTrue(approx(hook.get("C20"), 300.0),
                        f"C20（研发费用）应为 300，实际 {hook.get('C20')}")

    def test_admin_plus_rnd_equals_total(self):
        """第19行 + 第20行 = 6602 全额（关键恒等式）。

        【为什么这条最重要？】
            660253 同时归属两行：第19行减、第20行加。
            若映射写成单值字典（后写覆盖前写），会出现：
                · 只减不加 → 两行合计偏小（少一笔研发费用）
                · 只加不减 → 两行合计偏大（多算一笔）
                · 完全丢失 → 某一行为空
            而报表上每格都"有数字"，看不出任何异常。

            本条用一组具体数字验证恒等式，任一处串位立刻失败。
        """
        total_admin = 1000.0
        rnd = 300.0
        hook, _ = run_report([
            ("6602", 1, total_admin),
            ("660253", 1, rnd),
        ], month=1)

        admin = hook.get("C19")
        rnd_cell = hook.get("C20")
        self.assertTrue(
            approx((admin or 0) + (rnd_cell or 0), total_admin),
            f"管理费用 {admin} + 研发费用 {rnd_cell} 应等于 6602 全额 "
            f"{total_admin}，实际和 {(admin or 0) + (rnd_cell or 0)}")

    def test_twelve_months_both_rows(self):
        """两行各自 12 个月正确，O 列累计正确。"""
        rows = make_amount_rows({
            "6602": {m: m * 100.0 for m in range(1, 13)},
            "660253": {m: m * 30.0 for m in range(1, 13)},
        })
        hook, _ = run_report(rows, month=12)

        for month, col in COL_OF_MONTH.items():
            self.assertTrue(
                approx(hook.get(f"{col}19"), month * 70.0),
                f"{col}19（{month}月）应为 {month * 70.0}，"
                f"实际 {hook.get(f'{col}19')}")
            self.assertTrue(
                approx(hook.get(f"{col}20"), month * 30.0),
                f"{col}20（{month}月）应为 {month * 30.0}，"
                f"实际 {hook.get(f'{col}20')}")

        # 1..12 月 70 的倍数之和 = 70 * 78 = 5460
        self.assertTrue(approx(hook.get("O19"), 5460.0))
        # 30 * 78 = 2340
        self.assertTrue(approx(hook.get("O20"), 2340.0))
        # 两者相加应等于 6602 全年 = 100 * 78 = 7800
        self.assertTrue(approx(hook.get("O19") + hook.get("O20"), 7800.0))

    def test_admin_without_rnd(self):
        """660253 无记录时，管理费用 = 6602 全额（不减、不留空）。

        【注意列位】month=5 对应 5 月列 G，不是 C（1月）。
        """
        hook, _ = run_report([("6602", 5, 888.0)], month=5)

        self.assertTrue(approx(hook.get("G19"), 888.0),
                        f"G19（5月）应为 6602 全额 888，实际 {hook.get('G19')}")
        self.assertIsNone(hook.get("H20"),
                          f"研发费用无数据时应空白，实际 {hook.get('H20')}")

    def test_rnd_only_skips_admin(self):
        """只有 660253、没有 6602 时，C19 不写负数。

        【为什么？】
            与 LK01 的"只有减项"情形同一原则：has 只由**加项**设置。
            此时若照算，C19 会变成 -300（负数管理费用，不可能）。
            宁可留空让人查，也不写一个方向错误的数。
        """
        hook, _ = run_report([("660253", 1, 300.0)], month=1)

        self.assertIsNone(hook.get("C19"),
                          f"不应写入负数，实际 {hook.get('C19')}")
        # 但研发费用行应正常有数
        self.assertTrue(approx(hook.get("C20"), 300.0),
                        f"C20 应为 300，实际 {hook.get('C20')}")

    def test_direction_both_credit(self):
        """两行方向都必须是 credit。"""
        for key, excel_row in (("管理费用", 19), ("研发费用", 20)):
            with self.subTest(key=key):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.direction, "credit")
                self.assertEqual(spec.excel_row, excel_row)

    def test_admin_config_declares_exclude(self):
        """管理费用行必须声明减项 660253（配置层面固化口径）。"""
        spec = next(s for s in ROW_SPECS if s.key == "管理费用")

        self.assertEqual(spec.codes, ("6602",))
        self.assertEqual(spec.exclude_codes, ("660253",),
                         "管理费用应按净额列示，须减 660253 研发费用")

    def test_rnd_config_uses_detail_code(self):
        """研发费用行必须取明细码 660253（不是 6602）。"""
        spec = next(s for s in ROW_SPECS if s.key == "研发费用")

        self.assertEqual(spec.codes, ("660253",))
        self.assertEqual(spec.exclude_codes, ())

    def test_two_rows_not_crossed_with_other_expense(self):
        """管理费用/研发费用 不与其他成本费用行串用。"""
        rows = make_amount_rows({
            "6602": {3: 1000.0},
            "660253": {3: 300.0},
            "6601": {3: 5.0},      # 销售费用，不应进 19/20 行
            "6603": {3: 7.0},      # 其他科目，不应进 19/20 行
        })
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("E19"), 700.0),
                        f"E19 应为 700，实际 {hook.get('E19')}")
        self.assertTrue(approx(hook.get("E20"), 300.0),
                        f"E20 应为 300，实际 {hook.get('E20')}")
        self.assertTrue(approx(hook.get("E18"), 5.0),
                        f"E18（销售费用）应不受影响，实际 {hook.get('E18')}")

    def test_blank_when_no_data(self):
        """两行都无数据时整行空白。"""
        hook, _ = run_report([], month=1)

        for row in (19, 20):
            for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
                self.assertIsNone(hook.get(f"{col}{row}"),
                                  f"{col}{row} 无数据时应空白，"
                                  f"实际 {hook.get(f'{col}{row}')}")

    def test_excel_row_mapping(self):
        """行次 11 → Excel 第 19 行；行次 12 → Excel 第 20 行。"""
        self.assertEqual(EXCEL_ROW[11], 19)
        self.assertEqual(EXCEL_ROW[12], 20)

    def test_shared_code_in_both_rows(self):
        """660253 必须同时出现在两行的科目映射中（一减一加）。

        【为什么直接检查配置而非只看结果？】
            运行结果只能说明"这次取到了正确的数"，
            但无法说明"映射结构本身是否支持一个科目归属两行"。
            若将来有人把 code_to_rows 改成单值字典，
            某些场景下结果仍可能"碰巧正确"（例如两行恰好同号），
            这条从配置层面把"660253 属于两行"的事实钉死。
        """
        admin = next(s for s in ROW_SPECS if s.key == "管理费用")
        rnd = next(s for s in ROW_SPECS if s.key == "研发费用")

        self.assertIn("660253", admin.exclude_codes,
                      "660253 应作为管理费用的减项")
        self.assertIn("660253", rnd.codes,
                      "660253 应作为研发费用的加项")


class TestFinanceExpense(unittest.TestCase):
    """需求第 10 条：财务费用及其两个"其中"明细。

        ① 行次 13「其中：财务费用」6603    → C21:N21 + O21
        ② 行次 14「其中：利息费用」660311  → C22:N22 + O22
        ③ 行次 15「其中：利息收入」660302  → C23:N23 + O23

    【本类的核心验证点：第21行取全额、不减去两个下级】
        与第 9 条（管理费用减研发费用）形成对照：
        研发费用是**并列项目**（模板第20行不带"其中："），故要减；
        利息费用/利息收入带"其中："前缀，属**明细说明**，故不减。
        若这里也减，财务费用会偏小，且报表上看不出异常。
    """

    def test_c21_is_january(self):
        """C21 必须是 1 月财务费用（6603）全额。"""
        hook, _ = run_report([("6603", 1, 2600.0)], month=1)

        self.assertTrue(approx(hook.get("C21"), 2600.0),
                        f"C21（1月）应为 2600，实际 {hook.get('C21')}")
        self.assertTrue(approx(hook.get("O21"), 2600.0))
        self.assertIsNone(hook.get("D21"))

    def test_c22_is_january(self):
        """C22 必须是 1 月利息费用（660311）。"""
        hook, _ = run_report([("660311", 1, 1800.0)], month=1)

        self.assertTrue(approx(hook.get("C22"), 1800.0),
                        f"C22（1月）应为 1800，实际 {hook.get('C22')}")
        self.assertTrue(approx(hook.get("O22"), 1800.0))

    def test_c23_is_january(self):
        """C23 必须是 1 月利息收入（660302）。"""
        hook, _ = run_report([("660302", 1, 300.0)], month=1)

        self.assertTrue(approx(hook.get("C23"), 300.0),
                        f"C23（1月）应为 300，实际 {hook.get('C23')}")
        self.assertTrue(approx(hook.get("O23"), 300.0))

    def test_finance_expense_is_gross_not_net(self):
        """第21行取 6603 全额，**不减去**两个下级（与第9条对照）。

        【为什么这条最关键？】
            若照抄第 9 条管理费用的写法给本行加 exclude_codes
            （减去 660311、660302），财务费用会偏小，
            而营业利润会偏大 —— 两处同时错，报表上却每格都正常。

            本条给：6603=2600、660311=1800、660302=300
            正确：C21 = 2600（全额）
            误减：C21 = 2600-1800-300 = 500（错）
        """
        hook, _ = run_report([
            ("6603", 1, 2600.0),
            ("660311", 1, 1800.0),
            ("660302", 1, 300.0),
        ], month=1)

        self.assertTrue(
            approx(hook.get("C21"), 2600.0),
            f"C21（财务费用）应取 6603 全额 2600，不减去下级，"
            f"实际 {hook.get('C21')}")

    def test_finance_expense_config_has_no_exclude(self):
        """配置层面固化：财务费用行不得设 exclude_codes。"""
        spec = next(s for s in ROW_SPECS if s.key == "财务费用")

        self.assertEqual(spec.codes, ("6603",))
        self.assertEqual(
            spec.exclude_codes, (),
            "财务费用取 6603 全额；利息费用/利息收入是\"其中\"明细，"
            "非并列项目，不应扣除")

    def test_detail_rows_have_no_exclude_either(self):
        """两个明细行同样不设减项。"""
        for key, code in (("利息费用", "660311"), ("利息收入", "660302")):
            with self.subTest(key=key):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.codes, (code,))
                self.assertEqual(spec.exclude_codes, ())

    def test_three_rows_not_crossed(self):
        """三行互不串用（给数量级差异明显的金额）。"""
        rows = make_amount_rows({
            "6603": {3: 2600.0},
            "660311": {3: 1800.0},
            "660302": {3: 300.0},
        })
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("E21"), 2600.0),
                        f"E21（财务费用）应为 2600，实际 {hook.get('E21')}")
        self.assertTrue(approx(hook.get("E22"), 1800.0),
                        f"E22（利息费用）应为 1800，实际 {hook.get('E22')}")
        self.assertTrue(approx(hook.get("E23"), 300.0),
                        f"E23（利息收入）应为 300，实际 {hook.get('E23')}")

    def test_twelve_months_three_rows(self):
        """三行各自 12 个月正确，O 列累计正确。"""
        rows = make_amount_rows({
            "6603": {m: m * 100.0 for m in range(1, 13)},
            "660311": {m: m * 60.0 for m in range(1, 13)},
            "660302": {m: m * 10.0 for m in range(1, 13)},
        })
        hook, _ = run_report(rows, month=12)

        for month, col in COL_OF_MONTH.items():
            self.assertTrue(approx(hook.get(f"{col}21"), month * 100.0))
            self.assertTrue(approx(hook.get(f"{col}22"), month * 60.0))
            self.assertTrue(approx(hook.get(f"{col}23"), month * 10.0))

        # 1..12 之和 = 78
        self.assertTrue(approx(hook.get("O21"), 7800.0))
        self.assertTrue(approx(hook.get("O22"), 4680.0))
        self.assertTrue(approx(hook.get("O23"), 780.0))

    def test_direction_all_credit(self):
        """三行方向都必须是 credit。

        【为什么利息收入也取贷方？】
            利息收入是财务费用的减项，结转时分录为
            「借 本年利润 / 贷 财务费用-利息收入」，
            金额同样出现在贷方 mc，与其他费用科目方向一致。
        """
        for key, excel_row in (
            ("财务费用", 21),
            ("利息费用", 22),
            ("利息收入", 23),
        ):
            with self.subTest(key=key):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.direction, "credit")
                self.assertEqual(spec.excel_row, excel_row)

    def test_blank_when_no_data(self):
        """三行都无数据时整行空白。"""
        hook, _ = run_report([], month=1)

        for row in (21, 22, 23):
            for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
                self.assertIsNone(hook.get(f"{col}{row}"),
                                  f"{col}{row} 无数据时应空白，"
                                  f"实际 {hook.get(f'{col}{row}')}")

    def test_excel_row_mapping(self):
        """行次 13→21、14→22、15→23。"""
        self.assertEqual(EXCEL_ROW[13], 21)
        self.assertEqual(EXCEL_ROW[14], 22)
        self.assertEqual(EXCEL_ROW[15], 23)

    def test_not_crossed_with_admin_and_rnd(self):
        """财务费用系列不与管理费用/研发费用串用。"""
        rows = make_amount_rows({
            "6602": {3: 1000.0},
            "660253": {3: 300.0},
            "6603": {3: 2600.0},
            "660311": {3: 1800.0},
        })
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("E19"), 700.0),
                        f"E19（管理费用净额）应为 700，实际 {hook.get('E19')}")
        self.assertTrue(approx(hook.get("E20"), 300.0))
        self.assertTrue(approx(hook.get("E21"), 2600.0),
                        f"E21 应为 2600，实际 {hook.get('E21')}")
        self.assertTrue(approx(hook.get("E22"), 1800.0))


class TestOtherIncome(unittest.TestCase):
    """需求第 11 条：行次16「加：其他收益」（Excel 第 24 行）。

    科目 6117，本表首个 direction="signed" 的行。

    【本类的核心验证点：借贷两侧都要能取到，且符号正确】
        6117 在结转时既可能出现在借方（收益增加）也可能出现在贷方
        （收益冲回），现场确认"只能是唯一一个方向出现"。
        若固定取 md 或 mc 单列，当它出现在另一侧时本行空白 —— 静默漏数。
    """

    def test_debit_side_positive(self):
        """6117 出现在借方 → 正数（收益增加）。"""
        hook, _ = run_report([("6117", 1, 500.0, 0.0)], month=1)

        self.assertTrue(approx(hook.get("C24"), 500.0),
                        f"C24（借方收益）应为 +500，实际 {hook.get('C24')}")
        self.assertTrue(approx(hook.get("O24"), 500.0))

    def test_credit_side_negative(self):
        """6117 出现在贷方 → 负数（收益冲回，带负号填列）。

        【为什么这条最关键？】
            若实现成"固定取 md"，此场景取到 0，C24 空白 ——
            报表上看不出异常，只是"本月没有其他收益"。
            本条用贷方数据强制暴露该错误。
        """
        hook, _ = run_report([("6117", 1, 0.0, 300.0)], month=1)

        self.assertTrue(approx(hook.get("C24"), -300.0),
                        f"C24（贷方冲回）应为 -300，实际 {hook.get('C24')}")
        self.assertTrue(approx(hook.get("O24"), -300.0))

    def test_signed_net_when_both_sides(self):
        """两侧同时有值时按净额填列（借方 - 贷方）。

        【与现场口径的关系】
            现场说"只能唯一一个方向出现"，正常情况不会两侧都有。
            但若数据不遵守，仍按净额填列并打 WARNING（见 fetch），
            而不是静默丢弃一侧。
            本条给 500 / 200，期望净额 300。
        """
        hook, _ = run_report([("6117", 1, 500.0, 200.0)], month=1)

        self.assertTrue(approx(hook.get("C24"), 300.0),
                        f"C24（净额 500-200）应为 300，实际 {hook.get('C24')}")

    def test_twelve_months_signed(self):
        """12 个月正确；既有借方月也有贷方月时各自符号正确。

        数据设计：奇数月借方 100、偶数月贷方 100
            → 1月 +100, 2月 -100, 3月 +100, ...
        累计 = 6*100 - 6*100 = 0
        """
        rows = []
        for m in range(1, 13):
            if m % 2 == 1:
                rows.append(("6117", m, 100.0, 0.0))   # 借方
            else:
                rows.append(("6117", m, 0.0, 100.0))   # 贷方

        hook, _ = run_report(rows, month=12)

        for month, col in COL_OF_MONTH.items():
            expect = 100.0 if month % 2 == 1 else -100.0
            self.assertTrue(
                approx(hook.get(f"{col}24"), expect),
                f"{col}24（{month}月）应为 {expect}，"
                f"实际 {hook.get(f'{col}24')}")

        # 6 个 +100 与 6 个 -100 相抵
        self.assertTrue(approx(hook.get("O24"), 0.0),
                        f"O24 累计应为 0，实际 {hook.get('O24')}")

    def test_direction_is_signed(self):
        """配置层面固化：6117 必须是 signed 模式。"""
        spec = next(s for s in ROW_SPECS if s.key == "其他收益")

        self.assertEqual(spec.direction, "signed")
        self.assertEqual(spec.codes, ("6117",))
        self.assertEqual(spec.excel_row, EXCEL_ROW[16], "应写第 24 行")

    def test_blank_when_no_data(self):
        """无数据时整行空白（不写 0）。"""
        hook, _ = run_report([], month=1)

        for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
            self.assertIsNone(hook.get(f"{col}24"),
                              f"{col}24 无数据时应空白，"
                              f"实际 {hook.get(f'{col}24')}")

    def test_uses_two_column_query(self):
        """signed 模式必须走双列取数（SQL 含两个 SUM）。

        【为什么要看 SQL 文本？】
            若有人把 direction 误改成 debit/credit，
            SQL 会变成单列 SUM，贷方月就取不到数。
            本条直接断言 SQL 里有 SUM(md) AS debit_amount
            与 SUM(mc) AS credit_amount 两个别名。
        """
        _, ctx = run_report([])

        signed_calls = [
            c for c in ctx.fetch_calls
            if "debit_amount" in c["sql"] and "credit_amount" in c["sql"]
        ]
        self.assertEqual(
            len(signed_calls), 1,
            f"应恰有 1 条双列取数（signed 模式），"
            f"实际 {len(signed_calls)} 条")

        params = signed_calls[0]["params"]
        self.assertIn("6117", params)
        self.assertIn("4103", params)

    def test_not_crossed_with_finance_expense(self):
        """其他收益(24)不与财务费用(21)串用。"""
        rows = [
            ("6117", 3, 500.0, 0.0),      # 其他收益
            ("6603", 3, 2600.0, 0.0),     # 财务费用（模拟，正常应是 mc）
        ]
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("E24"), 500.0),
                        f"E24（其他收益）应为 500，实际 {hook.get('E24')}")
        # 财务费用行不应受 6117 影响
        self.assertNotEqual(hook.get("E21"), 500.0,
                            "财务费用不应被其他收益的数污染")

    def test_excel_row_mapping(self):
        """行次 16 → Excel 第 24 行。

        【为什么单独断言？】
            EXCEL_ROW 里第 16 行对应 24：
                行次 15 = 23（利息收入）
                行次 16 = 24（其他收益）
                行次 17 = 25（投资收益）
            容易误写成 23 或 25。
        """
        self.assertEqual(EXCEL_ROW[16], 24)





class TestTemplateConsistency(unittest.TestCase):
    """直接读模板 .xls，校验坐标是否真的落在有项目名的行上。"""

    def test_excel_rows_have_labels(self):
        """EXCEL_ROW 里每个 Excel 行，模板 A 列应有项目名。

        【为什么重要？】
            这条能抓住"整张表错行"——例如把行次 1 配成 Excel 第 10 行，
            那行是"其中：出口主营业务收入"，A 列虽非空但内容不对。
            这里至少保证"该行不是空行"，配合上面的锚点断言
            可覆盖绝大部分错行情形。

        【若模板文件不存在则跳过】，不影响其他测试。
        """
        try:
            import xlrd
        except ImportError:
            self.skipTest("未安装 xlrd")

        from config.settings import DEFAULT_TEMPLATE, PROJECT_ROOT

        tpl = PROJECT_ROOT / DEFAULT_TEMPLATE
        if not tpl.exists():
            self.skipTest(f"模板不存在：{tpl}")

        wb = xlrd.open_workbook(str(tpl))
        sh = wb.sheet_by_name("LK02")

        for row_no, excel_row in EXCEL_ROW.items():
            # xlrd 从 0 开始；A 列是索引 0
            val = str(sh.cell_value(excel_row - 1, 0)).strip()
            self.assertTrue(
                val,
                f"行次 {row_no} 应为 Excel 第 {excel_row} 行，"
                f"但该行 A 列是空的 —— 行号可能写错了")

    def test_formula_rows_are_summary_rows(self):
        """公式行（22/25/27）在模板中应是"二、三、四"开头的汇总项目。"""
        try:
            import xlrd
        except ImportError:
            self.skipTest("未安装 xlrd")

        from config.settings import DEFAULT_TEMPLATE, PROJECT_ROOT

        tpl = PROJECT_ROOT / DEFAULT_TEMPLATE
        if not tpl.exists():
            self.skipTest(f"模板不存在：{tpl}")

        wb = xlrd.open_workbook(str(tpl))
        sh = wb.sheet_by_name("LK02")

        for row_no in FORMULA_ROWS:
            excel_row = EXCEL_ROW[row_no]
            val = str(sh.cell_value(excel_row - 1, 0)).strip()
            self.assertTrue(
                val.startswith(("二、", "三、", "四、")),
                f"行次 {row_no}（Excel 第 {excel_row} 行）应是汇总行，"
                f"实际 A 列文字为 {val!r}")



if __name__ == "__main__":
    unittest.main(verbosity=2)
