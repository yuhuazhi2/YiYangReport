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

    def test_single_query_for_specs(self):
        """多个口径相同的行应合并成一次查询（性能与一致性）。"""
        _, ctx = run_report([])
        self.assertEqual(len(ctx.fetch_calls), 1,
                         f"应只发一次查询，实际 {len(ctx.fetch_calls)} 次")

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
        self.assertEqual(params[3], "600102", "第4个参数应为科目编码")

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
        self.assertIn("ccode IN (?)", sql)

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
        """direction 只能是 credit / debit。"""
        for spec in ROW_SPECS:
            self.assertIn(
                spec.direction, ("credit", "debit"),
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