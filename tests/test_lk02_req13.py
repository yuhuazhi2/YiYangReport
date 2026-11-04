# -*- coding: utf-8 -*-
"""LK02 需求第 13 条的独立测试文件：三行"暂时不予处理"的登记校验。

【为什么单独建文件？】
    与 test_lk02_req12.py 同样的原因：主测试文件 test_lk02_logic.py
    较长，编辑时多次出现插入位置偏移、类定义重复等误改。
    新需求单独成文件，结构简单、互不影响，unittest discover 都会跑到。

【本需求是什么？】
    需求第 13 条（20261104 现场确认）：
        第 26 行 公允价值变动收益
        第 27 行 信用减值损失
        第 28 行 资产减值损失
    这三行**暂时不予处理**。

【为什么要为"不做"写测试？】
    "不做"有两种状态，报表上看起来完全一样：
        A. 有意不做 → 已登记在 PENDING_ROWS，程序不写这些格
        B. 忘了做   → 既没实现也没登记，日后无法分辨
    本文件把 A 状态钉住，防止日后"整理代码"时把登记删掉、
    使状态退回不可分辨的 B。

【测什么】
    ① 结构：三行都在 PENDING_ROWS 中，行次/格位与模板一致
    ② 结构：三行都不在 ROW_SPECS（否则会写数）
    ③ 行为：给足可能被误配的科目数据，这三行仍保持空白
    ④ 行为：相邻行（24/25/29）不受影响
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from tests.test_lk02_logic import run_report  # noqa: E402

from reports.lk02_profit import (     # noqa: E402
    COL_OF_MONTH,
    COL_YEAR_TOTAL,
    EXCEL_ROW,
    PENDING_ROWS,
    ROW_SPECS,
)

# 需求第 13 条涉及的三行：(行次, Excel 行, 项目名, 登记格位)
PENDING_LINES = (
    (18, 26, "公允价值变动收益", "C26:N26"),
    (19, 27, "信用减值损失", "C27:N27"),
    (20, 28, "资产减值损失", "C28:N28"),
)


class TestPendingRowsRegistered(unittest.TestCase):
    """结构校验：三行必须显式登记在 PENDING_ROWS。"""

    def test_all_three_registered(self):
        """三行都必须在 PENDING_ROWS 中登记。

        【为什么这条最重要？】
            若有人"整理待定清单"时删掉其中某项，该行就退化为
            "既没实现、也没登记"。日后翻代码的人只看到那格是 0，
            无法判断是有意还是遗漏 —— 两种情况的处置方式完全不同。
        """
        names = {name for name, _cells, _row_no in PENDING_ROWS}
        for row_no, _excel_row, name, _cells in PENDING_LINES:
            self.assertIn(
                name, names,
                f"行次 {row_no}「{name}」应在 PENDING_ROWS 中登记，"
                f"当前清单：{sorted(names)}")

    def test_row_numbers_match(self):
        """登记的行次必须与模板一致（18/19/20）。

        【为什么核行次？】
            PENDING_ROWS 第三个字段是模板行次，用于人工对照模板。
            若填错，日后按行次去模板里找会对不上号，
            而格位（第二个字段）可能仍是对的 —— 矛盾却不易察觉。
        """
        mapping = {name: row_no for name, _c, row_no in PENDING_ROWS}
        for row_no, _excel_row, name, _cells in PENDING_LINES:
            self.assertEqual(
                mapping.get(name), row_no,
                f"「{name}」在 PENDING_ROWS 中登记的行次应为 {row_no}，"
                f"实际 {mapping.get(name)}")

    def test_cells_match_excel_row(self):
        """登记的格位必须与 Excel 行号自洽（格位里的数字 == 行号）。

        【为什么校验这个？】
            PENDING_ROWS 同时存了"格位字符串"和"行次"两个冗余信息。
            冗余的好处是可读性，风险是两处不同步：
            比如登记成 ("信用减值损失", "C28:N28", 19) ——
            格位写的是第 28 行、行次却是 19（模板里第 19 行是第 27 行格）。
            这种矛盾会让人不知道以哪个为准。
        """
        cells_map = {name: cells for name, cells, _row_no in PENDING_ROWS}
        for _row_no, excel_row, name, expected_cells in PENDING_LINES:
            self.assertEqual(
                cells_map.get(name), expected_cells,
                f"「{name}」登记的格位应为 {expected_cells}，"
                f"实际 {cells_map.get(name)}")
            # 格位内的行号必须等于 excel_row
            self.assertIn(
                str(excel_row), expected_cells,
                f"「{name}」格位 {expected_cells} 应指向 Excel 第 {excel_row} 行")

    def test_not_in_row_specs(self):
        """三行不得出现在 ROW_SPECS（否则程序会往它们写数）。

        【这是结构与行为的分界线】
            PENDING_ROWS 只是"声明"，不影响程序行为；
            真正决定写不写数的是 ROW_SPECS。
            若有人为了"顺手补上"往 ROW_SPECS 加了一条 excel_row=26 的配置，
            本测试会立刻失败，提示需先与财务确认口径。
        """
        specs_rows = {spec.excel_row for spec in ROW_SPECS}
        for row_no, excel_row, name, _cells in PENDING_LINES:
            self.assertNotIn(
                excel_row, specs_rows,
                f"行次 {row_no}「{name}」（Excel 第 {excel_row} 行）"
                f"已确认不做处理，不应出现在 ROW_SPECS 中")

    def test_excel_row_mapping(self):
        """模板行次 → Excel 行的映射必须是 18→26、19→27、20→28。"""
        for row_no, excel_row, _name, _cells in PENDING_LINES:
            self.assertEqual(
                EXCEL_ROW[row_no], excel_row,
                f"行次 {row_no} 应对应 Excel 第 {excel_row} 行，"
                f"实际 {EXCEL_ROW[row_no]}")


class TestPendingRowsNotWritten(unittest.TestCase):
    """行为校验：给足数据，三行仍必须保持空白。

    【与上一个类互补】
        上一个类查"配置对不对"，本类查"实际写不写"。
        配置正确不代表行为正确 —— 若写数逻辑误用行号
        （例如复制某行配置时漏改 excel_row），报表仍会多出数。
    """

    def test_blank_with_full_data(self):
        """给入大量科目数据，三行仍为空白。

        故意包含三行可能涉及的科目（6101/6701/6702 等），
        以及已实现的行（6111/6115/6117）——
        后者的数据必须只落到自己的行，不能溢出到 26/27/28。
        """
        rows = [
            # 已实现的三行 signed（不得溢出到 26/27/28）
            ("6117", 3, 1000.0, 0.0),
            ("6111", 3, 2000.0, 0.0),
            ("6115", 3, 3000.0, 0.0),
            # 三行"暂时不做"可能涉及的科目（不得被写入任何格）
            ("6101", 3, 9999.0, 0.0),
            ("6701", 3, 8888.0, 0.0),
            ("6702", 3, 7777.0, 0.0),
        ]
        hook, _ = run_report(rows, month=3)

        for row_no, excel_row, name, _cells in PENDING_LINES:
            for month, col in COL_OF_MONTH.items():
                cell = f"{col}{excel_row}"
                self.assertIsNone(
                    hook.get(cell),
                    f"{cell}（行次 {row_no}「{name}」，{month}月）"
                    f"已确认不做处理，不应被写入，实际 {hook.get(cell)}")
            total_cell = f"{COL_YEAR_TOTAL}{excel_row}"
            self.assertIsNone(
                hook.get(total_cell),
                f"{total_cell}（行次 {row_no}「{name}」本年累计）"
                f"不应被写入，实际 {hook.get(total_cell)}")

    def test_blank_with_no_data(self):
        """完全无数据时，三行也必须空白（不写 0）。"""
        hook, _ = run_report([], month=1)

        for row_no, excel_row, name, _cells in PENDING_LINES:
            for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
                cell = f"{col}{excel_row}"
                self.assertIsNone(
                    hook.get(cell),
                    f"{cell}（行次 {row_no}「{name}」）不应被写入，"
                    f"实际 {hook.get(cell)}")

    def test_neighbor_rows_still_work(self):
        """相邻行（24 其他收益 / 25 投资收益 / 29 资产处置收益）不受影响。

        【为什么测这个？】
            三行"不做"夹在已实现的行之间（24/25 与 29）。
            若将来有人为实现这三行而改动了行号映射，
            可能连带把相邻行也弄错。本条确认邻居正常，
            使"新增三行"的改动范围被限定在 26/27/28 内。
        """
        rows = [
            ("6117", 3, 1000.0, 0.0),
            ("6111", 3, 2000.0, 0.0),
            ("6115", 3, 3000.0, 0.0),
        ]
        hook, _ = run_report(rows, month=3)

        self.assertIsNotNone(hook.get("E24"), "E24（其他收益）应有数")
        self.assertIsNotNone(hook.get("E25"), "E25（投资收益）应有数")
        self.assertIsNotNone(hook.get("E29"), "E29（资产处置收益）应有数")


if __name__ == "__main__":
    unittest.main(verbosity=2)