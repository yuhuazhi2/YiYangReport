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
    COL_CURRENT,
    COL_PREV,
    EXCEL_ROW,
    FORMULA_ROWS,
    PENDING_ROWS,
    ROW_SPECS,
    LK03Report,
)


SPEC_KEY = "购买商品接受劳务支付的现金"


def _spec():
    return next(s for s in ROW_SPECS if s.key == SPEC_KEY)
"""LK03 需求第 6 条测试：C14 / D14（购买商品、接受劳务支付的现金）。"""

class TestConfig(unittest.TestCase):
    """配置层面的固化。"""

    def test_config_item_04_outflow(self):
        """本行只取项目 '04'，方向为 outflow，写第 14 行。"""
        spec = _spec()
        self.assertEqual(spec.cash_items, ("04",))
        self.assertEqual(spec.direction, "outflow",
                         f"本行是支付类，direction 必须是 outflow，"
                         f"实际 {spec.direction!r}")
        self.assertEqual(spec.excel_row, EXCEL_ROW[6], "应写第 14 行")
        self.assertTrue(spec.fill_current, "应填 C 列")
        self.assertTrue(spec.fill_prev, "应填 D 列")

    def test_direction_values_valid(self):
        """所有行的 direction 只能是 inflow / outflow。

        【为什么需要？】
            direction 拼错（如 "outFlow"、"output"）不会被 Python 报错，
            只会静默走进 else 分支当成 inflow —— 净额方向反掉，
            而报表上只是个数字，看不出正负应该是哪个方向。
        """
        allowed = {"inflow", "outflow"}
        for spec in ROW_SPECS:
            self.assertIn(spec.direction, allowed,
                          f"「{spec.key}」的 direction={spec.direction!r} 非法")

    def test_inflow_rows_unchanged(self):
        """行次 2/3/4 仍必须是 inflow（防止误改）。

        【为什么与上一条配对？】
            若有人为了"统一"把所有行都改成 outflow，
            前三行会全部变成负数 —— 而每行单独看都"有数字"。
        """
        for key in ("销售商品提供劳务收到的现金", "收到的税费返还",
                    "收到其他与经营活动有关的现金"):
            with self.subTest(key=key):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.direction, "inflow",
                                 f"「{key}」是收到类，应为 inflow")

    def test_row6_not_pending(self):
        """行次 6 已实现，应从 PENDING_ROWS 移除。"""
        pending_rows = {row_no for _, _, row_no in PENDING_ROWS}
        self.assertNotIn(6, pending_rows,
                         "行次 6 已实现，应从 PENDING_ROWS 移除")

    def test_cells_in_cell_map(self):
        """C14/D14 必须登记进 cell_map（供勾稽核对读回）。"""
        values = set(LK03Report.cell_map.values())
        for cell in ("C14", "D14"):
            self.assertIn(cell, values,
                          f"{cell} 未登记 cell_map，实际 {LK03Report.cell_map}")

class TestNetDirection(unittest.TestCase):
    """★ 本需求的核心：净额 = 贷方 - 借方（与前三行相反）。"""

    def test_c14_is_credit_minus_debit(self):
        """C14 = 贷方 mc － 借方 md。"""
        hook, _ = run_report(make_cash_rows({"04": (2000.0, 9000.0)}))
        self.assertTrue(
            approx(hook.get("C14"), 7000.0),
            f"C14 应为 9000-2000=7000（贷-借），实际 {hook.get('C14')}")
        # 若实现方向写反，结果会是 -7000 而不是空白。
        self.assertNotEqual(hook.get("C14"), -7000.0,
                            "C14 不应是按借方-贷方算出的 -7000")

    def test_mc_only_positive(self):
        """只有贷方（正常支付）时，C14 为正数。"""
        hook, _ = run_report(make_cash_rows({"04": (0.0, 5000.0)}))
        self.assertTrue(approx(hook.get("C14"), 5000.0),
                        f"C14 应为 5000，实际 {hook.get('C14')}")

    def test_md_only_negative(self):
        """只有借方时，C14 为负数（支付被冲回/退款）。"""
        hook, _ = run_report(make_cash_rows({"04": (800.0, 0.0)}))
        self.assertTrue(approx(hook.get("C14"), -800.0),
                        f"C14 应为 -800，实际 {hook.get('C14')}")

    def test_net_zero_written(self):
        """借贷相等时净额为 0，必须写 0（不是留空）。"""
        hook, _ = run_report(make_cash_rows({"04": (1200.0, 1200.0)}))
        self.assertEqual(hook.get("C14"), 0.0,
                         f"有记录且净额为 0 时应写 0，实际 {hook.get('C14')}")

    def test_no_data_blank(self):
        """无记录时 C14/D14 保持空白（不写 0）。"""
        hook, _ = run_report([])
        for cell in ("C14", "D14"):
            self.assertIsNone(hook.get(cell),
                              f"{cell} 无数据时应空白，实际 {hook.get(cell)}")

    def test_fetch_failure_not_crash(self):
        """取数失败时不崩溃，且不写 C14。"""
        hook, _ = run_report([], fail=True)
        self.assertIsNone(hook.get("C14"))

    def test_cross_month_records_summed(self):
        """同月多条记录按方向累加。"""
        rows = [
            ("04", 202605, 0.0, 1000.0),
            ("04", 202605, 200.0, 500.0),   # 同月第二条：付500退200
        ]
        hook, _ = run_report(rows)
        self.assertTrue(approx(hook.get("C14"), 1300.0),
                        f"C14 应为 (1000)+(500-200)=1300，实际 {hook.get('C14')}")


class TestPrevPeriod(unittest.TestCase):
    """D14 取上一期，规则与 D10 相同。"""

    def test_d14_is_prev_period(self):
        """选 8 月 → D14 取 7 月的 '04' 净额（同样贷-借）。"""
        hook, ctx = run_report(
            make_cash_rows({"04": (0.0, 8000.0)}, month=8),
            month=8,
            prev_rows=make_cash_rows({"04": (500.0, 3000.0)}, month=7))

        self.assertTrue(approx(hook.get("C14"), 8000.0),
                        f"C14 应为 8000，实际 {hook.get('C14')}")
        self.assertTrue(approx(hook.get("D14"), 2500.0),
                        f"D14 应为 3000-500=2500（贷-借），实际 {hook.get('D14')}")

        pc = call_with_item(prev_calls(ctx), "04")
        self.assertEqual(pc["params"][0], 202607,
                         f"上期 iYPeriod 应为 202607，实际 {pc['params'][0]}")

    def test_d14_january_cross_year(self):
        """★ 1 月时 D14 取上年 12 月，且查上一年度库。"""
        hook, ctx = run_report(
            make_cash_rows({"04": (0.0, 100.0)}, month=1),
            month=1,
            prev_rows=make_cash_rows({"04": (100.0, 700.0)},
                                     month=12, year=2025),
            prev_database="UFDATA_001_2025")

        self.assertTrue(approx(hook.get("D14"), 600.0),
                        f"D14 应为 700-100=600，实际 {hook.get('D14')}")

        pc = call_with_item(prev_calls(ctx), "04")
        self.assertEqual(pc["params"][0], 202512,
                         f"1 月上期 iYPeriod 应为 202512，"
                         f"实际 {pc['params'][0]}")
        self.assertEqual(pc["database"], "UFDATA_001_2025",
                         f"1 月上期应查上一年度库，实际 {pc['database']}")

    def test_d14_no_data_blank(self):
        """上期无 '04' 记录时 D14 留空。"""
        hook, _ = run_report(
            make_cash_rows({"04": (0.0, 1000.0)}, month=5),
            month=5, prev_rows=[])

        self.assertTrue(approx(hook.get("C14"), 1000.0))
        self.assertIsNone(hook.get("D14"),
                          f"上期无数据时 D14 应留空，实际 {hook.get('D14')}")

    def test_d14_first_year_january_blank(self):
        """首年 1 月：无上年库，D14 留空且不发起上期查询。"""
        hook, ctx = run_report(
            make_cash_rows({"04": (0.0, 500.0)}, month=1),
            month=1, prev_database="")

        self.assertTrue(approx(hook.get("C14"), 500.0))
        self.assertIsNone(hook.get("D14"))
        self.assertEqual(len(prev_calls(ctx)), 0,
                         "首年 1 月不应发起上期查询")

class TestNotCrossed(unittest.TestCase):
    """四行互不串用 —— 本类是整个文件最重要的护栏。"""

    def test_four_rows_not_crossed(self):
        """★ '01'/'02'/'03'/'04' 四行各归各行，且方向各自正确。

        【为什么数量级刻意拉开？】
            100000 / 1000 / 7 / 50000，任一处串位立刻让某个断言失败。
            若用相近的数，串位后可能"看起来差不多"而被忽略。
        """
        hook, _ = run_report(make_cash_rows({
            "01": (100000.0, 0.0),      # 流入：+100000
            "02": (1000.0, 0.0),        # 流入：+1000
            "03": (7.0, 0.0),           # 流入：+7
            "04": (1000.0, 50000.0),    # 流出：贷50000-借1000 = +49000
        }))

        self.assertTrue(approx(hook.get("C10"), 100000.0),
                        f"C10 应含 '01'=100000，实际 {hook.get('C10')}")
        self.assertTrue(approx(hook.get("C11"), 1000.0),
                        f"C11 应含 '02'=1000，实际 {hook.get('C11')}")
        self.assertTrue(approx(hook.get("C12"), 7.0),
                        f"C12 应含 '03'=7，实际 {hook.get('C12')}")
        self.assertTrue(approx(hook.get("C14"), 49000.0),
                        f"C14 应为 50000-1000=49000（贷-借），"
                        f"实际 {hook.get('C14')}")

    def test_four_rows_not_crossed_in_prev(self):
        """上期四行同样不串用、方向同样正确。"""
        hook, _ = run_report(
            make_cash_rows({"01": (0.0, 0.0), "02": (0.0, 0.0),
                            "03": (0.0, 0.0), "04": (0.0, 0.0)}, month=5),
            month=5,
            prev_rows=make_cash_rows({
                "01": (90000.0, 0.0),
                "02": (500.0, 0.0),
                "03": (9.0, 0.0),
                "04": (300.0, 20000.0),
            }, month=4))

        self.assertTrue(approx(hook.get("D10"), 90000.0),
                        f"D10 应含 '01'=90000，实际 {hook.get('D10')}")
        self.assertTrue(approx(hook.get("D11"), 500.0),
                        f"D11 应含 '02'=500，实际 {hook.get('D11')}")
        self.assertTrue(approx(hook.get("D12"), 9.0),
                        f"D12 应含 '03'=9，实际 {hook.get('D12')}")
        self.assertTrue(approx(hook.get("D14"), 19700.0),
                        f"D14 应为 20000-300=19700，实际 {hook.get('D14')}")

    def test_other_items_ignored(self):
        """非 '04' 的项目不得混入 C14。"""
        hook, _ = run_report(make_cash_rows({
            "01": (8888.0, 0.0),    # 属行次 2
            "02": (7777.0, 0.0),    # 属行次 3
            "03": (6666.0, 0.0),    # 属行次 4
            "04": (0.0, 1000.0),
            "05": (5555.0, 0.0),    # 未实现
        }))
        self.assertTrue(approx(hook.get("C14"), 1000.0),
                        f"C14 应只含 '04'=1000，实际 {hook.get('C14')}")

    def test_outflow_direction_isolated(self):
        """★ 方向隔离：流入行与流出行各自按自己的方向算。

        【这条防什么？】
            若 _aggregate 把 net 在循环外统一算好（而不是按 spec 逐行判断），
            同一批数据里既有流入行又有流出行时，必有一方方向错。
            本条给"借方有值、贷方为零"的数据：
              · 流入行 '01' 正常（借-贷 = 正）
              · 流出行 '04' 若也按借-贷算，会得到正数而非负数
            两条断言同时成立，才说明方向是按行分别判断的。
        """
        hook, _ = run_report(make_cash_rows({
            "01": (5000.0, 0.0),
            "04": (5000.0, 0.0),
        }))

        self.assertTrue(approx(hook.get("C10"), 5000.0),
                        f"C10（流入）应为 +5000，实际 {hook.get('C10')}")
        self.assertTrue(approx(hook.get("C14"), -5000.0),
                        f"C14（流出）应为 -5000，实际 {hook.get('C14')}")


class TestQuery(unittest.TestCase):
    """SQL 参数与时点。"""

    def test_query_uses_item_04_single_month(self):
        """SQL 参数必须带 '04'，且是单月精确匹配（不是区间）。"""
        _, ctx = run_report(make_cash_rows({"04": (0.0, 1.0)}), month=5)

        call = call_with_item(current_calls(ctx), "04")
        self.assertIn("gl_cashtable", call["sql"])
        self.assertIn("iYPeriod = ?", call["sql"])
        self.assertNotIn("BETWEEN", call["sql"],
                         "本表是单月口径，不应出现 BETWEEN 区间")
        self.assertEqual(call["params"][0], 202605)

    def test_query_month_follows_selection(self):
        """选择的年月变化时，参数跟着变。"""
        for month, expect in ((1, 202601), (8, 202608), (12, 202612)):
            with self.subTest(month=month):
                _, ctx = run_report(make_cash_rows({"04": (0.0, 1.0)}),
                                    month=month)
                call = call_with_item(current_calls(ctx), "04")
                self.assertEqual(call["params"][0], expect)

    def test_formula_rows_still_untouched(self):
        """有 '04' 数据时，公式行仍不得被写入。"""
        hook, _ = run_report(make_cash_rows({"04": (1000.0, 5000.0)}))
        for row_no in FORMULA_ROWS:
            excel_row = EXCEL_ROW[row_no]
            for col in (COL_CURRENT, COL_PREV):
                cell = f"{col}{excel_row}"
                self.assertIsNone(hook.get(cell),
                                  f"{cell}（行次{row_no}）是公式行，不得写入")


class TestTemplateConsistency(unittest.TestCase):
    """直接读模板，确认第 14 行确实是「购买商品、接受劳务支付的现金」。"""

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

    def test_row14_is_purchase_payment(self):
        """Excel 第 14 行 A 列应为「购买商品、接受劳务支付的现金」。"""
        sh = self._open_lk03()
        val = str(sh.cell_value(EXCEL_ROW[6] - 1, 0)).strip()
        self.assertIn("购买商品", val,
                      f"行次 6 的 A 列应为「购买商品、接受劳务支付的现金」，"
                      f"实际 {val!r} —— C14 可能填错行")

    def test_row14_is_between_inflow_and_outflow_subtotal(self):
        """第 14 行必须在第 13 行（流入小计）与第 18 行（流出小计）之间。

        【为什么验证位置关系？】
            本行的方向是 outflow，业务上它必须落在"经营活动现金流出"
            区段（行次 5 流入小计之后、行次 10 流出小计之前）。
            若模板调整导致它跑到流入区，方向就配错了 ——
            而数字本身依然"有值"，看不出异常。
        """
        self.assertGreater(EXCEL_ROW[6], EXCEL_ROW[5],
                           "行次 6 应在行次 5（流入小计）之后")
        self.assertLess(EXCEL_ROW[6], EXCEL_ROW[10],
                        "行次 6 应在行次 10（流出小计）之前")


if __name__ == "__main__":
    unittest.main(verbosity=2)