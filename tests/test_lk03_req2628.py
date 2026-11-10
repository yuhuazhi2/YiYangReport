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
FINANCE_OUTFLOWS = (
    (26, "20", 30, 38, "偿还债务支付的现金", 10000.0),
    (27, "21", 31, 39, "分配股利利润或偿付利息支付的现金", 2000.0),
    (28, "22", 32, 40, "支付其他与筹资活动有关的现金", 300.0),
)
"""LK03 需求第 26~28 条：筹资活动三项现金流出 C38~C40 / D38~D40。"""

class TestConfig(unittest.TestCase):
    """配置层面逐项核对。"""

    def test_each_row_config(self):
        """三项：项目编码、行次、方向逐条核对。"""
        for req, item, row_no, excel_row, key, _amt in FINANCE_OUTFLOWS:
            with self.subTest(req=req, item=item):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.cash_items, (item,),
                                 f"需求{req}：项目编码应为 {item!r}")
                self.assertEqual(spec.excel_row, EXCEL_ROW[row_no],
                                 f"需求{req}：应写 Excel 第 {excel_row} 行")
                self.assertEqual(spec.direction, "outflow",
                                 f"需求{req}：支付类现金应为 outflow")

    def test_items_unique_across_all_rows(self):
        """全表所有 ROW_SPECS 的项目编码不得重复。"""
        all_items = [c for s in ROW_SPECS for c in s.cash_items]
        self.assertEqual(len(all_items), len(set(all_items)),
                         f"全表项目编码有重复：{all_items}")

    def test_rows_not_pending(self):
        """行次 30~32 已实现，应从 PENDING_ROWS 移除。"""
        pending_rows = {row_no for _, _, row_no in PENDING_ROWS}
        for _req, _item, row_no, *_ in FINANCE_OUTFLOWS:
            self.assertNotIn(row_no, pending_rows,
                             f"行次 {row_no} 已实现，应从 PENDING_ROWS 移除")

    def test_cells_in_cell_map(self):
        """六格（C/D 各三）都必须登记进 cell_map。"""
        values = set(LK03Report.cell_map.values())
        for _req, _item, _row_no, excel_row, _key, _amt in FINANCE_OUTFLOWS:
            for col in ("C", "D"):
                cell = f"{col}{excel_row}"
                self.assertIn(cell, values, f"{cell} 未登记 cell_map")


class TestNetDirection(unittest.TestCase):
    """净额 = 贷方 - 借方（outflow 方向）。"""

    def test_each_row_own_value(self):
        """三项同时给数，各落各格。

        【为什么数量级是 10000 / 2000 / 300？】
            任一处串位，相邻两行的断言必然有一个不成立。
        """
        rows = make_cash_rows({
            # 支付类：贷方为增项
            item: (0.0, amt) for _r, item, _n, _e, _k, amt in FINANCE_OUTFLOWS
        })
        hook, _ = run_report(rows)

        for req, _item, _row_no, excel_row, key, amt in FINANCE_OUTFLOWS:
            cell = f"C{excel_row}"
            self.assertTrue(
                approx(hook.get(cell), amt),
                f"需求{req}：{cell}（{key}）应为 {amt}，实际 {hook.get(cell)}")

    def test_net_is_credit_minus_debit(self):
        """净额 = 贷方 - 借方（与需求 2 条反向）。"""
        hook, _ = run_report(make_cash_rows({"20": (1200.0, 5000.0)}))
        self.assertTrue(approx(hook.get("C38"), 3800.0),
                        f"C38 应为 5000-1200=3800，实际 {hook.get('C38')}")
        self.assertNotEqual(hook.get("C38"), -3800.0,
                            "C38 不应是按借方-贷方算出的 -3800")

    def test_md_only_negative(self):
        """只有借方时净额为负（还款冲回）。"""
        hook, _ = run_report(make_cash_rows({"21": (400.0, 0.0)}))
        self.assertTrue(approx(hook.get("C39"), -400.0),
                        f"C39 应为 -400，实际 {hook.get('C39')}")

    def test_net_zero_written(self):
        """有记录且净额为 0 时写 0（不是留空）。"""
        hook, _ = run_report(make_cash_rows({"22": (600.0, 600.0)}))
        self.assertEqual(hook.get("C40"), 0.0,
                         f"C40 有记录且净额为 0 时应写 0，"
                         f"实际 {hook.get('C40')}")

    def test_no_data_blank(self):
        """无记录时六格全部保持空白。"""
        hook, _ = run_report([])
        for _req, _item, _n, excel_row, _k, _a in FINANCE_OUTFLOWS:
            for col in ("C", "D"):
                cell = f"{col}{excel_row}"
                self.assertIsNone(hook.get(cell),
                                  f"{cell} 无数据时应空白，"
                                  f"实际 {hook.get(cell)}")

    def test_fetch_failure_not_crash(self):
        """取数失败时不崩溃，且不写这六格。"""
        hook, _ = run_report([], fail=True)
        for _req, _item, _n, excel_row, _k, _a in FINANCE_OUTFLOWS:
            self.assertIsNone(hook.get(f"C{excel_row}"))
            self.assertIsNone(hook.get(f"D{excel_row}"))

    def test_other_items_ignored(self):
        """未实现的项目编码不得混入这六格。"""
        hook, _ = run_report(make_cash_rows({
            "20": (0.0, 1000.0),
            "19": (9999.0, 0.0),   # 属筹资流入（行次 28）
            "23": (8888.0, 0.0),   # 未实现
        }))
        self.assertTrue(approx(hook.get("C38"), 1000.0),
                        f"C38 应只含 '20'=1000，实际 {hook.get('C38')}")

    def test_finance_inflow_outflow_not_crossed(self):
        """★ 筹资流入（'17'~'19'）与筹资流出（'20'~'22'）不串用。

        【为什么这条最关键？】
            两组编码紧邻（17~19 与 20~22），方向相反。
            若串位，会出现"流入格显示流出数"或"方向套用错"，
            而报表上每格都有值，肉眼无法分辨。
        """
        hook, _ = run_report(make_cash_rows({
            # 筹资流入（借方为增项）
            "17": (50000.0, 0.0),
            "18": (5000.0, 0.0),
            "19": (500.0, 0.0),
            # 筹资流出（贷方为增项）
            "20": (0.0, 30000.0),
            "21": (0.0, 3000.0),
            "22": (0.0, 300.0),
        }))

        self.assertTrue(approx(hook.get("C34"), 50000.0))
        self.assertTrue(approx(hook.get("C35"), 5000.0))
        self.assertTrue(approx(hook.get("C36"), 500.0))
        self.assertTrue(approx(hook.get("C38"), 30000.0))
        self.assertTrue(approx(hook.get("C39"), 3000.0))
        self.assertTrue(approx(hook.get("C40"), 300.0))

class TestPrevPeriod(unittest.TestCase):
    """D 列取上一期，规则与 D14 相同。"""

    def test_d_column_is_prev_period(self):
        """三项同时给上期数据，各落各 D 格。"""
        cur = make_cash_rows({
            item: (0.0, 0.0) for _r, item, _n, _e, _k, _a in FINANCE_OUTFLOWS
        }, month=8)
        prev = make_cash_rows({
            item: (0.0, amt)
            for _r, item, _n, _e, _k, amt in FINANCE_OUTFLOWS
        }, month=7)

        hook, _ = run_report(cur, month=8, prev_rows=prev)

        for req, _item, _n, excel_row, key, amt in FINANCE_OUTFLOWS:
            cell = f"D{excel_row}"
            self.assertTrue(
                approx(hook.get(cell), amt),
                f"需求{req}：{cell}（{key}）上期应为 {amt}，"
                f"实际 {hook.get(cell)}")

    def test_d_january_cross_year(self):
        """★ 1 月时 D 列取上年 12 月，且查上一年度库。"""
        cur = make_cash_rows({"20": (0.0, 100.0)}, month=1)
        prev = make_cash_rows({"20": (300.0, 900.0)},
                              month=12, year=2025)

        hook, ctx = run_report(cur, month=1, prev_rows=prev,
                               prev_database="UFDATA_001_2025")

        self.assertTrue(approx(hook.get("D38"), 600.0),
                        f"D38 应为 900-300=600，实际 {hook.get('D38')}")

        pc = call_with_item(prev_calls(ctx), "20")
        self.assertEqual(pc["params"][0], 202512,
                         f"1 月上期 iYPeriod 应为 202512，"
                         f"实际 {pc['params'][0]}")
        self.assertEqual(pc["database"], "UFDATA_001_2025",
                         f"1 月上期应查上一年度库，实际 {pc['database']}")

    def test_d_no_data_blank(self):
        """上期无对应记录时 D 列留空。"""
        hook, _ = run_report(
            make_cash_rows({"20": (0.0, 1000.0)}, month=5),
            month=5, prev_rows=[])
        self.assertIsNone(hook.get("D38"),
                          f"上期无数据时 D38 应留空，实际 {hook.get('D38')}")


class TestQuery(unittest.TestCase):
    """SQL 参数与时点。"""

    def test_each_item_queried(self):
        """三个项目编码都应出现在查询参数里。"""
        _, ctx = run_report([])
        all_params = []
        for c in current_calls(ctx):
            all_params.extend(c["params"] or [])
        for _req, item, *_ in FINANCE_OUTFLOWS:
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
        """三项都有数据时，公式行仍不得被写入。

        【特别注意行次 33「筹资活动现金流出小计」（Excel 第 41 行）】
            本需求新增的三项正是它的加总来源，该行已按需求 25
            登记为"保持原样"。若程序误写，会覆盖模板公式。
        """
        rows = make_cash_rows({
            item: (0.0, amt) for _r, item, _n, _e, _k, amt in FINANCE_OUTFLOWS
        })
        hook, _ = run_report(rows)
        for row_no in FORMULA_ROWS:
            excel_row = EXCEL_ROW[row_no]
            for col in ("C", "D"):
                cell = f"{col}{excel_row}"
                self.assertIsNone(hook.get(cell),
                                  f"{cell}（行次{row_no}）是公式行，不得写入")


class TestTemplateConsistency(unittest.TestCase):
    """直接读模板，确认第 38~40 行的项目名。"""

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
        """第 38~40 行的 A 列应与配置的项目名对应。"""
        sh = self._open_lk03()
        expect_keywords = {
            38: "偿还债务",
            39: "分配股利",
            40: "支付其他与筹资活动",
        }
        for excel_row, kw in expect_keywords.items():
            val = str(sh.cell_value(excel_row - 1, 0)).strip()
            self.assertIn(kw, val,
                          f"第 {excel_row} 行 A 列应含「{kw}」，"
                          f"实际 {val!r} —— 行号可能写错了")

    def test_rows_in_financing_outflow_section(self):
        """第 38~40 行必须落在筹资流出区段内。

        【区段边界（行次）】
            34 = 筹资活动产生的现金流量净额（Excel 第 42 行）
            33 = 筹资活动现金流出小计（Excel 第 41 行）
        本需求三项（行次 30~32）应在流入小计(29)之后、流出小计(33)之前。
        """
        self.assertGreater(EXCEL_ROW[30], EXCEL_ROW[29],
                           "行次 30 应在筹资流入小计之后")
        self.assertLess(EXCEL_ROW[32], EXCEL_ROW[33],
                        "行次 32 应在筹资流出小计之前")


if __name__ == "__main__":
    unittest.main(verbosity=2)