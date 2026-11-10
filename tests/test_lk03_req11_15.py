# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
import unittest
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from tests.fakes import approx  # noqa: E402
from tests.test_lk03_logic import (  # noqa: E402
    call_with_item,
    current_calls,
    make_cash_rows,
    prev_calls,
    run_report,
)

from reports.lk03_cashflow import (        # noqa: E402
    EXCEL_ROW,
    FORMULA_ROWS,
    PENDING_ROWS,
    ROW_SPECS,
    LK03Report,
)


# (需求号, 项目编码, 模板行次, Excel 行号, key, 校验用金额)
INVEST_INFLOWS = (
    (11, "08", 13, 21, "收回投资收到的现金", 10000.0),
    (12, "09", 14, 22, "取得投资收益收到的现金", 2000.0),
    (13, "10", 15, 23, "处置长期资产收回的现金净额", 300.0),
    (14, "11", 16, 24, "处置子公司收到现金净额", 40.0),
    (15, "12", 17, 25, "收到其他与投资活动有关的现金", 5.0),
)
"""LK03 需求第 11~15 条：投资活动五项现金流入 C21~C25 / D21~D25。"""

class TestConfig(unittest.TestCase):
    """配置层面逐项核对。"""

    def test_each_row_config(self):
        """五项：项目编码、行次、方向逐条核对。"""
        for req, item, row_no, excel_row, key, _amt in INVEST_INFLOWS:
            with self.subTest(req=req, item=item):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.cash_items, (item,),
                                 f"需求{req}：项目编码应为 {item!r}")
                self.assertEqual(spec.excel_row, EXCEL_ROW[row_no],
                                 f"需求{req}：应写 Excel 第 {excel_row} 行")
                self.assertEqual(spec.direction, "inflow",
                                 f"需求{req}：收到类现金应为 inflow")

    def test_no_duplicate_items(self):
        """五行项目编码不得重复。

        【为什么？】
            若两条配置用同一个编码，同一条取数记录会同时累加到
            两个格子 —— 报表上两行显示同一个数，且都"有值"。
        """
        items = [item for _, item, *_ in INVEST_INFLOWS]
        self.assertEqual(len(items), len(set(items)),
                         f"项目编码有重复：{items}")

    def test_no_duplicate_excel_rows(self):
        """五行不得占用同一 Excel 行。"""
        excel_rows = [row[3] for row in INVEST_INFLOWS]
        self.assertEqual(len(excel_rows), len(set(excel_rows)))

    def test_rows_not_pending(self):
        """行次 13~17 已实现，应从 PENDING_ROWS 移除。"""
        pending_rows = {row_no for _, _, row_no in PENDING_ROWS}
        for _req, _item, row_no, *_ in INVEST_INFLOWS:
            self.assertNotIn(row_no, pending_rows,
                             f"行次 {row_no} 已实现，应从 PENDING_ROWS 移除")

    def test_cells_in_cell_map(self):
        """十格（C/D 各五）都必须登记进 cell_map。"""
        values = set(LK03Report.cell_map.values())
        for _req, _item, _row_no, excel_row, _key, _amt in INVEST_INFLOWS:
            for col in ("C", "D"):
                cell = f"{col}{excel_row}"
                self.assertIn(cell, values, f"{cell} 未登记 cell_map")

class TestNetValues(unittest.TestCase):
    """★ 五行各归各格，且方向与算法正确。"""

    def test_each_row_own_value(self):
        """五项同时给数，各落各格。

        【为什么数量级是 10000 / 2000 / 300 / 40 / 5？】
            任一处串位，相邻两行的断言必然有一个不成立。
            若都用相近的数，串位后可能"看起来差不多"而被忽略。
        """
        rows = make_cash_rows({
            item: (amt, 0.0) for _r, item, _n, _e, _k, amt in INVEST_INFLOWS
        })
        hook, _ = run_report(rows)

        for req, _item, _row_no, excel_row, key, amt in INVEST_INFLOWS:
            cell = f"C{excel_row}"
            self.assertTrue(
                approx(hook.get(cell), amt),
                f"需求{req}：{cell}（{key}）应为 {amt}，实际 {hook.get(cell)}")

    def test_net_is_debit_minus_credit(self):
        """净额 = 借方 - 贷方（与需求 2 条同向）。"""
        hook, _ = run_report(make_cash_rows({"08": (5000.0, 1200.0)}))
        self.assertTrue(approx(hook.get("C21"), 3800.0),
                        f"C21 应为 5000-1200=3800，实际 {hook.get('C21')}")

    def test_md_only(self):
        """只有借方时等于 md。"""
        hook, _ = run_report(make_cash_rows({"09": (700.0, 0.0)}))
        self.assertTrue(approx(hook.get("C22"), 700.0),
                        f"C22 应为 700，实际 {hook.get('C22')}")

    def test_mc_only_negative(self):
        """只有贷方时净额为负（冲回）。"""
        hook, _ = run_report(make_cash_rows({"10": (0.0, 400.0)}))
        self.assertTrue(approx(hook.get("C23"), -400.0),
                        f"C23 应为 -400，实际 {hook.get('C23')}")

    def test_net_zero_written(self):
        """有记录且净额为 0 时写 0（不是留空）。"""
        hook, _ = run_report(make_cash_rows({"11": (600.0, 600.0)}))
        self.assertEqual(hook.get("C24"), 0.0,
                         f"C24 有记录且净额为 0 时应写 0，"
                         f"实际 {hook.get('C24')}")

    def test_no_data_blank(self):
        """无记录时十格全部保持空白。"""
        hook, _ = run_report([])
        for _req, _item, _n, excel_row, _k, _a in INVEST_INFLOWS:
            for col in ("C", "D"):
                cell = f"{col}{excel_row}"
                self.assertIsNone(hook.get(cell),
                                  f"{cell} 无数据时应空白，"
                                  f"实际 {hook.get(cell)}")

    def test_fetch_failure_not_crash(self):
        """取数失败时不崩溃，且不写这十格。"""
        hook, _ = run_report([], fail=True)
        for _req, _item, _n, excel_row, _k, _a in INVEST_INFLOWS:
            self.assertIsNone(hook.get(f"C{excel_row}"))
            self.assertIsNone(hook.get(f"D{excel_row}"))

    def test_other_items_ignored(self):
        """未实现的项目编码不得混入这五行。"""
        hook, _ = run_report(make_cash_rows({
            "08": (1000.0, 0.0),
            "07": (9999.0, 0.0),   # 未实现
            "13": (8888.0, 0.0),   # 未实现
        }))
        self.assertTrue(approx(hook.get("C21"), 1000.0),
                        f"C21 应只含 '08'=1000，实际 {hook.get('C21')}")

    def test_not_crossed_with_operating_rows(self):
        """投资活动五行不得与经营活动已实现行（'01'~'04'）串用。"""
        hook, _ = run_report(make_cash_rows({
            "01": (100000.0, 0.0),
            "02": (20000.0, 0.0),
            "03": (3000.0, 0.0),
            "04": (0.0, 40000.0),
            "08": (500.0, 0.0),
            "09": (60.0, 0.0),
            "10": (7.0, 0.0),
            "11": (0.8, 0.0),
            "12": (0.09, 0.0),
        }))

        self.assertTrue(approx(hook.get("C10"), 100000.0))
        self.assertTrue(approx(hook.get("C11"), 20000.0))
        self.assertTrue(approx(hook.get("C12"), 3000.0))
        self.assertTrue(approx(hook.get("C14"), 40000.0))
        self.assertTrue(approx(hook.get("C21"), 500.0))
        self.assertTrue(approx(hook.get("C22"), 60.0))
        self.assertTrue(approx(hook.get("C23"), 7.0))
        self.assertTrue(approx(hook.get("C24"), 0.8))
        self.assertTrue(approx(hook.get("C25"), 0.09))

class TestPrevPeriod(unittest.TestCase):
    """D 列取上一期，规则与 D10 相同。"""

    def test_d_column_is_prev_period(self):
        """五项同时给上期数据，各落各 D 格。"""
        cur = make_cash_rows({
            item: (0.0, 0.0) for _r, item, _n, _e, _k, _a in INVEST_INFLOWS
        }, month=8)
        prev = make_cash_rows({
            item: (amt, 0.0)
            for _r, item, _n, _e, _k, amt in INVEST_INFLOWS
        }, month=7)

        hook, _ = run_report(cur, month=8, prev_rows=prev)

        for req, _item, _n, excel_row, key, amt in INVEST_INFLOWS:
            cell = f"D{excel_row}"
            self.assertTrue(
                approx(hook.get(cell), amt),
                f"需求{req}：{cell}（{key}）上期应为 {amt}，"
                f"实际 {hook.get(cell)}")

    def test_d_january_cross_year(self):
        """★ 1 月时 D 列取上年 12 月，且查上一年度库。"""
        cur = make_cash_rows({"08": (100.0, 0.0)}, month=1)
        prev = make_cash_rows({"08": (900.0, 300.0)},
                              month=12, year=2025)

        hook, ctx = run_report(cur, month=1, prev_rows=prev,
                               prev_database="UFDATA_001_2025")

        self.assertTrue(approx(hook.get("D21"), 600.0),
                        f"D21 应为 900-300=600，实际 {hook.get('D21')}")

        pc = call_with_item(prev_calls(ctx), "08")
        self.assertEqual(pc["params"][0], 202512,
                         f"1 月上期 iYPeriod 应为 202512，"
                         f"实际 {pc['params'][0]}")
        self.assertEqual(pc["database"], "UFDATA_001_2025",
                         f"1 月上期应查上一年度库，实际 {pc['database']}")

    def test_d_no_data_blank(self):
        """上期无对应记录时 D 列留空。"""
        hook, _ = run_report(
            make_cash_rows({"08": (1000.0, 0.0)}, month=5),
            month=5, prev_rows=[])
        self.assertIsNone(hook.get("D21"),
                          f"上期无数据时 D21 应留空，实际 {hook.get('D21')}")


class TestQuery(unittest.TestCase):
    """SQL 参数与时点。"""

    def test_each_item_queried(self):
        """五个项目编码都应出现在查询参数里。"""
        _, ctx = run_report([])
        all_params = []
        for c in current_calls(ctx):
            all_params.extend(c["params"] or [])
        for _req, item, *_ in INVEST_INFLOWS:
            self.assertIn(item, all_params,
                          f"项目编码 {item!r} 未出现在任何查询参数中")

    def test_single_month_not_range(self):
        """每条查询都是单月精确匹配，不出现 BETWEEN。"""
        _, ctx = run_report([])
        for c in current_calls(ctx):
            self.assertIn("iYPeriod = ?", c["sql"])
            self.assertNotIn("BETWEEN", c["sql"])


class TestFormulaRowsUntouched(unittest.TestCase):
    """公式行仍不得写入。"""

    def test_formula_rows_still_untouched(self):
        """五项都有数据时，公式行仍不得被写入。

        【特别注意行次 18「投资活动现金流入小计」（Excel 第 26 行）】
            本需求新增的五项正是它的加总来源。
            若程序误写该行，会覆盖模板公式，且下月无法恢复。
        """
        rows = make_cash_rows({
            item: (amt, 0.0) for _r, item, _n, _e, _k, amt in INVEST_INFLOWS
        })
        hook, _ = run_report(rows)
        for row_no in FORMULA_ROWS:
            excel_row = EXCEL_ROW[row_no]
            for col in ("C", "D"):
                cell = f"{col}{excel_row}"
                self.assertIsNone(hook.get(cell),
                                  f"{cell}（行次{row_no}）是公式行，不得写入")


class TestTemplateConsistency(unittest.TestCase):
    """直接读模板，确认第 21~25 行的项目名与行次对应。"""

    def _open_lk03(self):
        try:
            import xlrd
        except ImportError:
            self.skipTest("未安装 xlrd")

        from config.settings import DEFAULT_TEMPLATE, PROJECT_ROOT

        tpl = PROJECT_ROOT / DEFAULT_TEMPLATE
        if not tpl.exists():
            self.skipTest(f"模板不存在：{tpl}")

        wb = xlrd.open_workbook(str(tpl))
        return wb.sheet_by_name("LK03")

    def test_row_labels_match(self):
        """第 21~25 行的 A 列应与配置的项目名对应。

        【为什么用关键词而不是全文比对？】
            模板 A 列可能有前导空格、全角符号差异。
            用关键词能稳定匹配，同时足够发现"行号写错一位"。
        """
        sh = self._open_lk03()
        expect_keywords = {
            21: "收回投资",
            22: "投资收益",
            23: "处置固定资产",
            24: "处置子公司",
            25: "收到其他与投资活动",
        }
        for excel_row, kw in expect_keywords.items():
            val = str(sh.cell_value(excel_row - 1, 0)).strip()
            self.assertIn(kw, val,
                          f"第 {excel_row} 行 A 列应含「{kw}」，"
                          f"实际 {val!r} —— 行号可能写错了")

    def test_rows_in_investing_section(self):
        """第 21~25 行必须落在投资活动区段内。

        【区段边界（行次）】
            12 = 投资活动标题（Excel 第 20 行）
            18 = 投资活动现金流入小计（Excel 第 26 行）
        本需求五行（行次 13~17）应严格落在两者之间。
        """
        self.assertGreater(EXCEL_ROW[13], EXCEL_ROW[12],
                           "行次 13 应在投资活动标题之后")
        self.assertLess(EXCEL_ROW[17], EXCEL_ROW[18],
                        "行次 17 应在投资活动流入小计之前")


if __name__ == "__main__":
    unittest.main(verbosity=2)