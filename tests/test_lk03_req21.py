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
    cell_row,
    make_cash_rows,
    run_report,
)

from reports.lk03_cashflow import (        # noqa: E402
    EXCEL_ROW,
    FORMULA_ROWS,
    INTENTIONALLY_UNTOUCHED,
    PENDING_ROWS,
    ROW_SPECS,
    SECTION_ROWS,
    LK03Report,
)


# (格位, 模板行次, Excel 行号, 项目名关键词)
UNTOUCHED = (
    ("C31", 23, 31, "投资活动现金流出小计"),
    ("D31", 23, 31, "投资活动现金流出小计"),
    ("C32", 24, 32, "投资活动产生的现金流量净额"),
    ("D32", 24, 32, "投资活动产生的现金流量净额"),
)
"""LK03 需求第 21 条：C31/D31、C32/D32 不做处理。"""

class TestConfig(unittest.TestCase):
    """配置层面：4 个格子已登记且位置正确。"""

    def test_all_four_registered(self):
        """4 个格子都必须出现在 INTENTIONALLY_UNTOUCHED 里。"""
        for cell, *_ in UNTOUCHED:
            self.assertIn(cell, INTENTIONALLY_UNTOUCHED,
                          f"{cell} 应登记为不做处理，"
                          f"实际 {INTENTIONALLY_UNTOUCHED}")

    def test_cells_match_excel_row(self):
        """格子行号必须与 EXCEL_ROW 对应。

        【为什么锁这条？】
            若把 C32 写成 C33（行次 25「三、筹资活动产生的现金流量：」），
            就会变成"跳过真实净额行 + 点名标题行"——
            标题本就不写、看着像生效，而真实净额行却会被当成可写行。
            两处都错，报表上看不出来。
        """
        for cell, row_no, excel_row, _kw in UNTOUCHED:
            self.assertEqual(EXCEL_ROW[row_no], excel_row,
                             f"行次 {row_no} 应对应 Excel 第 {excel_row} 行")
            self.assertEqual(cell_row(cell), excel_row,
                             f"{cell} 的行号与行次 {row_no} 不符")

    def test_cells_are_formula_rows(self):
        """4 个格子所在行都应是公式行（模板自带合计公式）。"""
        formula_cells = {
            f"{col}{EXCEL_ROW[n]}"
            for n in FORMULA_ROWS for col in ("C", "D")
        }
        for cell, *_ in UNTOUCHED:
            self.assertIn(cell, formula_cells,
                          f"{cell} 不在 FORMULA_ROWS 覆盖范围内 —— "
                          f"该行若没有模板公式，需重新评估是否要填数")

    def test_not_in_cell_map(self):
        """4 个格子不得登记进 cell_map。"""
        cm_values = set(LK03Report.cell_map.values())
        for cell, *_ in UNTOUCHED:
            self.assertNotIn(cell, cm_values,
                             f"{cell} 不写数值，不应登记进 cell_map")

    def test_not_in_row_specs(self):
        """4 个格子不得被任何 ROW_SPECS 占用。"""
        spec_cells = set()
        for spec in ROW_SPECS:
            spec_cells.add(f"C{spec.excel_row}")
            spec_cells.add(f"D{spec.excel_row}")
        for cell, *_ in UNTOUCHED:
            self.assertNotIn(cell, spec_cells,
                             f"{cell} 已确认不做处理，"
                             f"却被某条 ROW_SPECS 占用")

    def test_not_in_pending_rows(self):
        """4 个格子不得出现在待定清单里。"""
        pending_cells = set()
        for _name, cells, _row_no in PENDING_ROWS:
            pending_cells.update(cells.split("/"))
        for cell, *_ in UNTOUCHED:
            self.assertNotIn(cell, pending_cells,
                             f"{cell} 是确定不做，不应登记为待定项")

    def test_rows_not_detail_rows(self):
        """这两行（23/24）都不是明细行。

        【为什么？】
            它们是合计/净额行，由模板公式算出。
            若误把某个明细行登记为"不做处理"，该明细会被静默跳过，
            其小计因失去数据来源而偏小 —— 报表上每格都有数，
            只有合计不对。
        """
        detail_rows = (set(EXCEL_ROW.keys())
                       - set(FORMULA_ROWS) - set(SECTION_ROWS))
        for _cell, row_no, _excel_row, _kw in UNTOUCHED:
            self.assertNotIn(row_no, detail_rows,
                             f"行次 {row_no} 是明细行，不应登记为不做处理")

    def test_row23_row24_relationship(self):
        """行次 23 与 24 的 Excel 行号必须相邻（31、32）。

        【为什么？】
            模板中「流出小计」与「净额」是紧邻的两行。
            若 EXCEL_ROW 里这两条被写错（例如 23→32、24→31 互换），
            登记的四格会落到对方行上 —— 而两行都是公式行、
            都"不该写"，测试若只查"是公式行"就抓不出来。
        """
        self.assertEqual(EXCEL_ROW[24] - EXCEL_ROW[23], 1,
                         f"行次 23/24 应相邻，实际 "
                         f"{EXCEL_ROW[23]}/{EXCEL_ROW[24]}")

class TestBehavior(unittest.TestCase):
    """实际运行时不写入这 4 格。"""

    def test_never_written_with_no_data(self):
        """无任何数据时，4 格保持空白。"""
        hook, _ = run_report([])
        for cell, *_ in UNTOUCHED:
            self.assertIsNone(hook.get(cell),
                              f"{cell} 已确认不做处理，不应被写入，"
                              f"实际 {hook.get(cell)}")

    def test_never_written_with_full_data(self):
        """★ 即使投资活动明细行数据充足，这 4 格也不得被写入。

        【本条最有价值的地方：给的是"能被加总出来"的数据】
            行次 23 = 行次 19~22 之和（需求 17~20 已全部实现）
            行次 24 = 行次 18 - 行次 23
            若程序误写这些合计行，会用"程序算的合计"覆盖
            "模板公式的合计"—— 表面数字还一样，但公式已经没了，
            下个月数据变化时不会自动重算。

            所以断言"必须为 None"，而不是"等于某个合计数"。
        """
        hook, _ = run_report(
            make_cash_rows({
                "13": (0.0, 10000.0),   # 行次 19
                "14": (0.0, 2000.0),    # 行次 20
                "15": (0.0, 300.0),     # 行次 21
                "16": (0.0, 40.0),      # 行次 22
                "08": (50000.0, 0.0),   # 行次 13（流入，用于算净额）
            }, month=5),
            month=5)

        # 【前置条件】先确认明细行确实有数，否则本测试等于没测
        self.assertTrue(approx(hook.get("C27"), 10000.0),
                        f"前置条件不成立：C27 应有数据，"
                        f"实际 {hook.get('C27')}")
        self.assertTrue(approx(hook.get("C30"), 40.0),
                        f"前置条件不成立：C30 应有数据，"
                        f"实际 {hook.get('C30')}")

        for cell, *_ in UNTOUCHED:
            self.assertIsNone(hook.get(cell),
                              f"{cell} 已确认不做处理，不得写入，"
                              f"实际 {hook.get(cell)}")

    def test_never_written_with_prev_data(self):
        """上期数据充足时，D 列两格同样不得被写入。"""
        hook, _ = run_report(
            make_cash_rows({"13": (0.0, 0.0)}, month=5),
            month=5,
            prev_rows=make_cash_rows({
                "13": (0.0, 8000.0),
                "14": (0.0, 900.0),
            }, month=4))

        for cell in ("D31", "D32"):
            self.assertIsNone(hook.get(cell),
                              f"{cell} 已确认不做处理，不得写入，"
                              f"实际 {hook.get(cell)}")

    def test_never_written_with_fetch_failure(self):
        """取数失败时，4 格同样保持空白。"""
        hook, _ = run_report([], fail=True)
        for cell, *_ in UNTOUCHED:
            self.assertIsNone(hook.get(cell),
                              f"取数失败时 {cell} 应保持空白，"
                              f"实际 {hook.get(cell)}")


class TestTemplateConsistency(unittest.TestCase):
    """直接读模板，确认两行的项目名。"""

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

    def test_row_labels(self):
        """两行的 A 列应与预期项目名对应。"""
        sh = self._open_lk03()
        checked = {31: "流出小计", 32: "净额"}
        for excel_row, kw in checked.items():
            val = str(sh.cell_value(excel_row - 1, 0)).strip()
            self.assertIn(kw, val,
                          f"第 {excel_row} 行 A 列应含「{kw}」，"
                          f"实际 {val!r} —— 行号可能写错了")

    def test_row31_is_investing_outflow_subtotal(self):
        """第 31 行应明确是「投资活动」的流出小计。"""
        sh = self._open_lk03()
        val = str(sh.cell_value(30, 0)).strip()
        self.assertIn("投资", val,
                      f"第 31 行应属投资活动区段，实际 {val!r}")

    def test_rows_have_formula_or_value(self):
        """两行 C 列在模板中都应有内容（公式或 0.0 占位）。

        【为什么查这个？】
            若模板该行是空单元格，说明它不是预置的合计行，
            那"不做处理"会让整个合计缺失。此时必须回头确认需求。
        """
        sh = self._open_lk03()
        for excel_row in (31, 32):
            cell = sh.cell(excel_row - 1, 2)   # C 列
            has_content = (cell.ctype != 0) or bool(str(cell.value).strip())
            self.assertTrue(
                has_content,
                f"C{excel_row} 在模板中应为预置的合计值/公式，"
                f"实际是空单元格 —— 「不做处理」的决定需重新确认")


if __name__ == "__main__":
    unittest.main(verbosity=2)