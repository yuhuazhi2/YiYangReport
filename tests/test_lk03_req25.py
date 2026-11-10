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
    ("C37", 29, 37, "筹资活动现金流入小计"),
    ("D37", 29, 37, "筹资活动现金流入小计"),
    ("C41", 33, 41, "筹资活动现金流出小计"),
    ("D41", 33, 41, "筹资活动现金流出小计"),
    ("C42", 34, 42, "筹资活动产生的现金流量净额"),
    ("D42", 34, 42, "筹资活动产生的现金流量净额"),
    ("C44", 36, 44, "现金及现金等价物净增加额"),
    ("D44", 36, 44, "现金及现金等价物净增加额"),
    ("C46", 38, 46, "期末现金及现金等价物余额"),
    ("D46", 38, 46, "期末现金及现金等价物余额"),
)
"""LK03 需求第 25 条：筹资区段与汇总行共 10 格不做处理。"""

class TestConfig(unittest.TestCase):
    """配置层面：10 个格子已登记且位置正确。"""

    def test_all_ten_registered(self):
        """10 个格子都必须出现在 INTENTIONALLY_UNTOUCHED 里。"""
        for cell, *_ in UNTOUCHED:
            self.assertIn(cell, INTENTIONALLY_UNTOUCHED,
                          f"{cell} 应登记为不做处理，"
                          f"实际 {INTENTIONALLY_UNTOUCHED}")

    def test_cells_match_excel_row(self):
        """格子行号必须与 EXCEL_ROW 对应。

        【为什么锁这条？】
            若把 C42 写成 C43（行次 35「四、汇率变动对现金的影响」），
            就会变成"跳过真实净额行 + 点名一个明细行"——
            而 C43 本属于待定项，登记成"不做处理"会与 PENDING_ROWS 冲突。
        """
        for cell, row_no, excel_row, _kw in UNTOUCHED:
            self.assertEqual(EXCEL_ROW[row_no], excel_row,
                             f"行次 {row_no} 应对应 Excel 第 {excel_row} 行")
            self.assertEqual(cell_row(cell), excel_row,
                             f"{cell} 的行号与行次 {row_no} 不符")

    def test_cells_are_formula_rows(self):
        """10 个格子所在行都应是公式行（模板自带合计公式）。"""
        formula_cells = {
            f"{col}{EXCEL_ROW[n]}"
            for n in FORMULA_ROWS for col in ("C", "D")
        }
        for cell, *_ in UNTOUCHED:
            self.assertIn(cell, formula_cells,
                          f"{cell} 不在 FORMULA_ROWS 覆盖范围内 —— "
                          f"该行若没有模板公式，需重新评估是否要填数")

    def test_not_in_cell_map(self):
        """10 个格子不得登记进 cell_map。"""
        cm_values = set(LK03Report.cell_map.values())
        for cell, *_ in UNTOUCHED:
            self.assertNotIn(cell, cm_values,
                             f"{cell} 不写数值，不应登记进 cell_map")

    def test_not_in_row_specs(self):
        """10 个格子不得被任何 ROW_SPECS 占用。"""
        spec_cells = set()
        for spec in ROW_SPECS:
            spec_cells.add(f"C{spec.excel_row}")
            spec_cells.add(f"D{spec.excel_row}")
        for cell, *_ in UNTOUCHED:
            self.assertNotIn(cell, spec_cells,
                             f"{cell} 已确认不做处理，"
                             f"却被某条 ROW_SPECS 占用")

    def test_not_in_pending_rows(self):
        """10 个格子不得出现在待定清单里。

        【这条对本批尤其重要】
            行次 35（C43/D43，汇率变动影响）确实在待定清单里，
            与本批的 C42/D42 只差一行。若把 C42 错写成 C43，
            既会让 C42 变成"可写"、又会与待定清单冲突 ——
            本条能同时抓住。
        """
        pending_cells = set()
        for _name, cells, _row_no in PENDING_ROWS:
            pending_cells.update(cells.split("/"))
        for cell, *_ in UNTOUCHED:
            self.assertNotIn(cell, pending_cells,
                             f"{cell} 是确定不做，不应登记为待定项")

    def test_rows_not_detail_rows(self):
        """这 5 行都不是明细行。"""
        detail_rows = (set(EXCEL_ROW.keys())
                       - set(FORMULA_ROWS) - set(SECTION_ROWS))
        for _cell, row_no, _excel_row, _kw in UNTOUCHED:
            self.assertNotIn(row_no, detail_rows,
                             f"行次 {row_no} 是明细行，不应登记为不做处理")


class TestFullFormulaCoverage(unittest.TestCase):
    """★ 长期护栏：11 个公式行必须全部登记为"不做处理"。

    【为什么这条值得单独一组？】
        需求分四批（5、16、21、25）陆续登记了 22 个格子。
        若某一批漏了一行，后果是：
          · 该行仍是公式行，程序也不会写它 —— 报表看起来完全正常；
          · 但它不在"业务上确定不写"的清单里。
        日后有人看代码时，会以为"这行还没处理"，
        可能顺手给它加一条 ROW_SPECS 配置从而覆盖模板公式。

        本条把"公式行集合 == 不做处理集合"钉死，
        任何新公式行未登记都会立刻失败。
    """

    def _formula_cells(self) -> set[str]:
        return {
            f"{col}{EXCEL_ROW[n]}"
            for n in FORMULA_ROWS for col in ("C", "D")
        }

    def test_all_formula_cells_registered(self):
        """公式行的全部格子都应在 INTENTIONALLY_UNTOUCHED 中。"""
        missing = self._formula_cells() - set(INTENTIONALLY_UNTOUCHED)
        self.assertFalse(
            missing,
            f"以下公式行格子未登记为「不做处理」：{sorted(missing)} —— "
            f"若它们确实应保持原样，请登记；"
            f"若某行其实要填数，请从 FORMULA_ROWS 移除")

    def test_no_extra_cells_registered(self):
        """登记为「不做处理」的格子必须都是公式行。

        【反向校验】
            若某个明细行被误登记为"不做处理"，该明细会被静默跳过，
            其所属的小计因缺数据而偏小 —— 报表上每格都有数、
            只有合计不对，极难发现。
        """
        extra = set(INTENTIONALLY_UNTOUCHED) - self._formula_cells()
        self.assertFalse(
            extra,
            f"以下格子被登记为「不做处理」，但它们不在公式行内："
            f"{sorted(extra)} —— 若这些是该填数的明细行，"
            f"登记为不做处理会导致漏数")

    def test_cell_count_matches(self):
        """两边格子数必须相等（22 = 11 行 × 2 列）。"""
        self.assertEqual(len(INTENTIONALLY_UNTOUCHED),
                         len(self._formula_cells()),
                         f"登记 {len(INTENTIONALLY_UNTOUCHED)} 格，"
                         f"公式行共 {len(self._formula_cells())} 格")
        self.assertEqual(len(FORMULA_ROWS) * 2,
                         len(self._formula_cells()))

class TestBehavior(unittest.TestCase):
    """实际运行时不写入这 10 格。"""

    def test_never_written_with_no_data(self):
        """无任何数据时，10 格保持空白。"""
        hook, _ = run_report([])
        for cell, *_ in UNTOUCHED:
            self.assertIsNone(hook.get(cell),
                              f"{cell} 已确认不做处理，不应被写入，"
                              f"实际 {hook.get(cell)}")

    def test_never_written_with_full_data(self):
        """★ 即使全表明细行数据充足，这 10 格也不得被写入。

        【本条最有价值的地方：给的是"能被加总出来"的数据】
            行次 29 = 行次 26~28 之和
            行次 33 = 行次 30~32 之和（尚未实现，故只给能给的）
            行次 34 = 行次 29 - 行次 33
            行次 36 = 11 + 24 + 34 + 35
            行次 38 = 36 + 37
            若程序误写，会用"程序算的合计"覆盖"模板公式"——
            表面数字一样，但公式已失，下月数据变化时不会重算。

            所以断言必须是 None，而非"等于某个合计数"。
        """
        hook, _ = run_report(
            make_cash_rows({
                # 筹资流入三项（行次 26~28），是行次 29 的数据来源
                "17": (10000.0, 0.0),
                "18": (2000.0, 0.0),
                "19": (300.0, 0.0),
                # 经营与投资也各给一些，凑出净增加额的输入
                "01": (50000.0, 0.0),
                "08": (8000.0, 0.0),
                "13": (0.0, 4000.0),
            }, month=5),
            month=5)

        # 【前置条件】确认明细行确实有数，否则本测试等于没测
        self.assertTrue(approx(hook.get("C34"), 10000.0),
                        f"前置条件不成立：C34 应有数据，"
                        f"实际 {hook.get('C34')}")
        self.assertTrue(approx(hook.get("C36"), 300.0),
                        f"前置条件不成立：C36 应有数据，"
                        f"实际 {hook.get('C36')}")

        for cell, *_ in UNTOUCHED:
            self.assertIsNone(hook.get(cell),
                              f"{cell} 已确认不做处理，不得写入，"
                              f"实际 {hook.get(cell)}")

    def test_never_written_with_prev_data(self):
        """上期数据充足时，D 列五格同样不得被写入。"""
        hook, _ = run_report(
            make_cash_rows({"17": (0.0, 0.0)}, month=5),
            month=5,
            prev_rows=make_cash_rows({
                "17": (9000.0, 0.0),
                "18": (800.0, 0.0),
            }, month=4))

        for cell in ("D37", "D41", "D42", "D44", "D46"):
            self.assertIsNone(hook.get(cell),
                              f"{cell} 已确认不做处理，不得写入，"
                              f"实际 {hook.get(cell)}")

    def test_never_written_with_fetch_failure(self):
        """取数失败时，10 格同样保持空白。"""
        hook, _ = run_report([], fail=True)
        for cell, *_ in UNTOUCHED:
            self.assertIsNone(hook.get(cell),
                              f"取数失败时 {cell} 应保持空白，"
                              f"实际 {hook.get(cell)}")


class TestTemplateConsistency(unittest.TestCase):
    """直接读模板，确认 5 行的项目名与位置。"""

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
        """5 行的 A 列应与预期项目名对应。"""
        sh = self._open_lk03()
        checked = {
            37: "筹资活动现金流入小计",
            41: "筹资活动现金流出小计",
            42: "筹资活动产生的现金流量净额",
            44: "现金及现金等价物净增加额",
            46: "期末现金及现金等价物余额",
        }
        for excel_row, kw in checked.items():
            val = str(sh.cell_value(excel_row - 1, 0)).strip()
            self.assertIn(kw, val,
                          f"第 {excel_row} 行 A 列应含「{kw}」，"
                          f"实际 {val!r} —— 行号可能写错了")

    def test_rows_have_formula_or_value(self):
        """5 行 C 列在模板中都应有内容（公式或 0.0 占位）。

        【为什么查这个？】
            若模板某行是空单元格，说明它不是预置的合计行，
            那"不做处理"会让整个合计缺失。此时必须回头确认需求。
        """
        sh = self._open_lk03()
        for excel_row in (37, 41, 42, 44, 46):
            cell = sh.cell(excel_row - 1, 2)   # C 列
            has_content = (cell.ctype != 0) or bool(str(cell.value).strip())
            self.assertTrue(
                has_content,
                f"C{excel_row} 在模板中应为预置的合计值/公式，"
                f"实际是空单元格 —— 「不做处理」的决定需重新确认")

    def test_rows_excel_order(self):
        """5 行的 Excel 行号必须严格递增（37 < 41 < 42 < 44 < 46）。

        【为什么？】
            它们分别是筹资流入小计、流出小计、净额、净增加额、期末余额，
            业务上必然按此顺序排列。若 EXCEL_ROW 里某条写错
            （例如 36↔38 互换），登记的格子会落到错的合计行上 ——
            而两行都"不该写"，只查"是公式行"抓不出来。
        """
        rows = [EXCEL_ROW[n] for n in (29, 33, 34, 36, 38)]
        self.assertEqual(rows, sorted(rows),
                         f"五行 Excel 行号应递增，实际 {rows}")
        self.assertEqual(rows, [37, 41, 42, 44, 46],
                         f"五行应落在 37/41/42/44/46，实际 {rows}")


if __name__ == "__main__":
    unittest.main(verbosity=2)