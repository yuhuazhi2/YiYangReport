# -*- coding: utf-8 -*-
"""LK03 现金流量表 —— 需求第 5 条：C13 / D13 不做处理。

运行方式（在项目根目录执行）：
    python -m unittest discover -s tests -v

【本文件测什么？】
    需求第 5 条（20261110 确认）：C13 / D13 保持原样，程序不写入。

    这两格对应行次 5「经营活动现金流入小计」（Excel 第 13 行）。
    该行由模板公式自行合计本期三类经营流入（行次 2、3、4）。

【为什么单独一个测试文件？】
    与 LK02 的 test_lk02_req12/13/14/15 同一惯例：
    按"需求条目"分文件，日后回溯"这条需求怎么验的"时能直接定位，
    也避免主测试文件随需求堆积而变得难以浏览。

【与 FORMULA_ROWS 的关系（两者重叠但语义不同）】
    FORMULA_ROWS            → 技术约束："写了会覆盖模板合计公式"
    INTENTIONALLY_UNTOUCHED → 业务决定："需求要求该格保持原样"

    当前 C13 同时属于两者。分开登记的价值：若将来模板改版、
    该行不再是公式行，FORMULA_ROWS 会调整，而"需求要求不写"
    这个决定不会变。只记在 FORMULA_ROWS 里，届时它会被当成
    "可以填数了"而误写。
"""
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
    COL_CURRENT,
    COL_PREV,
    EXCEL_ROW,
    FORMULA_ROWS,
    INTENTIONALLY_UNTOUCHED,
    PENDING_ROWS,
    ROW_SPECS,
    LK03Report,
)


class TestConfig(unittest.TestCase):
    """配置层面的固化。

    【注意：INTENTIONALLY_UNTOUCHED 之后又加入了需求 16 的 6 个格子】
        本类只验证"C13/D13 这两格本身"的正确性，
        不再断言"清单恰好只有这两格"——那是需求 5 当时的说法，
        需求 16 补充后在**同一个常量**里追加了新格子。
        需求 16 的格子由 tests/test_lk03_req16.py 单独验证。
    """

    def test_c13_d13_present(self):
        """C13 与 D13 必须登记在清单里。"""
        self.assertIn("C13", INTENTIONALLY_UNTOUCHED,
                      f"C13 应登记为不做处理，实际 {INTENTIONALLY_UNTOUCHED}")
        self.assertIn("D13", INTENTIONALLY_UNTOUCHED,
                      f"D13 应登记为不做处理，实际 {INTENTIONALLY_UNTOUCHED}")

    def test_cells_are_on_row_5(self):
        """C13/D13 所在行必须是行次 5（经营活动现金流入小计）。

        【为什么锁这条？】
            "不做处理"必须对应到正确的那一行。若把行号写错
            （例如写成 C14，那是行次 6「购买商品支付的现金」，
            一行真实需要填数的明细行），该行就会永远空白 ——
            而报表上只是个空格，看不出是"有意不填"还是"取数失败"。
        """
        self.assertEqual(EXCEL_ROW[5], 13, "行次 5 应对应 Excel 第 13 行")
        for cell in ("C13", "D13"):
            self.assertEqual(cell_row(cell), EXCEL_ROW[5],
                             f"{cell} 不在行次 5 上，可能写错了行号")

    def test_cells_are_formula_rows(self):
        """这两格所在行应同时属于公式行（模板自带合计公式）。

        【这条是交叉校验】
            若模板该行其实没有公式，那"不做处理"就值得重新评估 ——
            因为合计行若靠人工填，跳过它会漏一整个合计数。
            本条失败时提示的正是这个疑问。
        """
        formula_cells = {
            f"{col}{EXCEL_ROW[n]}"
            for n in FORMULA_ROWS for col in (COL_CURRENT, COL_PREV)
        }
        for cell in INTENTIONALLY_UNTOUCHED:
            self.assertIn(cell, formula_cells,
                          f"{cell} 不在 FORMULA_ROWS 覆盖范围内 —— "
                          f"该行若没有模板公式，需重新评估是否要填数")

    def test_not_in_cell_map(self):
        """这两格不得登记进 cell_map。

        cell_map 是"确实写了数值的格子"的清单。
        登记一个永远空白的格子，verify() 会读回 None，
        核对报告里该指标永远是"无法核对"，属于噪音，
        会掩盖真正的问题。
        """
        cm_values = set(LK03Report.cell_map.values())
        for cell in INTENTIONALLY_UNTOUCHED:
            self.assertNotIn(cell, cm_values,
                             f"{cell} 不写数值，不应登记进 cell_map")

    def test_not_in_row_specs(self):
        """这两格不得被任何 ROW_SPECS 行占用。"""
        spec_cells = set()
        for spec in ROW_SPECS:
            spec_cells.add(f"{COL_CURRENT}{spec.excel_row}")
            spec_cells.add(f"{COL_PREV}{spec.excel_row}")
        for cell in INTENTIONALLY_UNTOUCHED:
            self.assertNotIn(cell, spec_cells,
                             f"{cell} 已确认不做处理，"
                             f"却被某条 ROW_SPECS 占用")

    def test_not_in_pending_rows(self):
        """这两格不得出现在待定清单里。

        【为什么？】
            PENDING_ROWS 是"口径待定、等财务给条件就能做"。
            本需求是"确定不做"，混进去会让人误以为还有待办，
            也会破坏"待定 + 已实现 = 24 个明细行"的覆盖恒等式
            （行次 5 是合计行，不在那 24 个明细行里）。
        """
        pending_cells = set()
        for _name, cells, _row_no in PENDING_ROWS:
            pending_cells.update(cells.split("/"))
        for cell in INTENTIONALLY_UNTOUCHED:
            self.assertNotIn(cell, pending_cells,
                             f"{cell} 是确定不做，不应登记为待定项")


class TestBehavior(unittest.TestCase):
    """实际运行时不写入这两格。"""

    def test_never_written_with_no_data(self):
        """无任何数据时，这两格保持空白。"""
        hook, _ = run_report([])
        for cell in INTENTIONALLY_UNTOUCHED:
            self.assertIsNone(hook.get(cell),
                              f"{cell} 已确认不做处理，不应被写入，"
                              f"实际 {hook.get(cell)}")

    def test_never_written_with_full_data(self):
        """★ 即使上游明细行有充足数据，这两格也不得被写入。

        【与上一条互补：一条查空输入，一条查满输入】
            配置正确不代表行为正确。若别处逻辑误写这两格
            （例如把某行 spec 的 excel_row 配成 13），
            报表仍会多出数、覆盖模板合计公式。
            本条给足数据跑一遍，这两格必须仍为 None。
        """
        hook, _ = run_report(
            make_cash_rows({
                "01": (100000.0, 0.0),
                "02": (2000.0, 0.0),
                "03": (300.0, 0.0),
            }, month=5),
            month=5,
            prev_rows=make_cash_rows({
                "01": (90000.0, 0.0),
                "02": (1500.0, 0.0),
                "03": (200.0, 0.0),
            }, month=4))

        # 【前置条件】先确认上游明细行确实有数，否则本测试等于没测
        self.assertTrue(approx(hook.get("C12"), 300.0),
                        f"前置条件不成立：C12 应有数据，实际 {hook.get('C12')}")
        self.assertTrue(approx(hook.get("D12"), 200.0),
                        f"前置条件不成立：D12 应有数据，实际 {hook.get('D12')}")

        for cell in INTENTIONALLY_UNTOUCHED:
            self.assertIsNone(hook.get(cell),
                              f"{cell} 已确认不做处理，不得写入，"
                              f"实际 {hook.get(cell)}")

    def test_never_written_with_fetch_failure(self):
        """取数失败时，这两格同样保持空白。"""
        hook, _ = run_report([], fail=True)
        for cell in INTENTIONALLY_UNTOUCHED:
            self.assertIsNone(hook.get(cell),
                              f"取数失败时 {cell} 应保持空白，"
                              f"实际 {hook.get(cell)}")


class TestTemplateConsistency(unittest.TestCase):
    """直接读模板，确认第 13 行确实是「经营活动现金流入小计」。"""

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

    def test_row13_is_inflow_subtotal(self):
        """Excel 第 13 行 A 列应为「经营活动现金流入小计」。"""
        sh = self._open_lk03()
        val = str(sh.cell_value(EXCEL_ROW[5] - 1, 0)).strip()
        self.assertIn("流入小计", val,
                      f"第 13 行 A 列应为「经营活动现金流入小计」，"
                      f"实际 {val!r} —— 行号可能写错了")

    def test_row13_has_formula_or_value(self):
        """第 13 行 C 列在模板中应有内容（公式或 0.0 占位）。

        【为什么查这个？】
            模板该行若完全是空白，说明它并不是预置的合计行，
            那么"不做处理"会让整个经营流入小计缺失 ——
            这种情形下必须回头确认需求。
        """
        sh = self._open_lk03()
        cell = sh.cell(EXCEL_ROW[5] - 1, 2)   # C 列
        has_content = (cell.ctype != 0) or bool(str(cell.value).strip())
        self.assertTrue(
            has_content,
            f"C13 在模板中应为预置的合计值/公式，实际是空单元格 —— "
            f"该行若需人工填数，「不做处理」的决定需重新确认")


if __name__ == "__main__":
    unittest.main(verbosity=2)