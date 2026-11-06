# -*- coding: utf-8 -*-
"""LK02 需求第 12 条的独立测试文件。

【为什么单独建一个文件，而不是加进 test_lk02_logic.py？】
    需求第 12 条（投资收益 6111、资产处置收益 6115）与第 11 条
    （其他收益 6117）口径完全相同，都是 direction="signed"。
    本可并入 test_lk02_logic.py 的 TestOtherIncome 类旁。

    但编辑过程中该文件多次被工具误覆盖（插入位置偏移、类定义重复、
    新增类被删除）。为避免继续反复修补、耽误开发，
    把本次需求单独放一个文件：
        · 结构简单，不易被误改；
        · 与主测试文件互不影响（unittest discover 两个都会跑）；
        · 复用主文件的 run_report 等工具函数，不重复实现。

【测什么】
    ① 投资收益 6111 → C25:N25 + O25，signed 模式
    ② 资产处置收益 6115 → C29:N29 + O29，signed 模式
    ③ 三行 signed（6117/6111/6115）互不串用、合并为一次查询
    ④ 第 26/27/28 行（需求未提及）必须保持空白
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from tests.fakes import approx  # noqa: E402
from tests.test_lk02_logic import run_report  # noqa: E402

from reports.lk02_profit import (     # noqa: E402
    COL_OF_MONTH,
    COL_YEAR_TOTAL,
    EXCEL_ROW,
    ROW_SPECS,
)


class TestInvestmentGain(unittest.TestCase):
    """需求第 12 条第 1 项：行次 17「投资收益」6111 → C25:N25 + O25。"""

    def test_debit_side_positive(self):
        """6111 出现在借方 → 正数（投资收益）。"""
        hook, _ = run_report([("6111", 1, 800.0, 0.0)], month=1)

        self.assertTrue(approx(hook.get("C25"), 800.0),
                        f"C25（投资收益）应为 +800，实际 {hook.get('C25')}")
        self.assertTrue(approx(hook.get("O25"), 800.0))

    def test_credit_side_negative(self):
        """6111 出现在贷方 → 负数（投资损失）。

        【为什么这条最关键？】
            若固定取 md，此场景取到 0，C25 空白，报表上看不出异常。
            本条用贷方数据强制暴露该错误。
        """
        hook, _ = run_report([("6111", 1, 0.0, 200.0)], month=1)

        self.assertTrue(approx(hook.get("C25"), -200.0),
                        f"C25（投资损失）应为 -200，实际 {hook.get('C25')}")
        self.assertTrue(approx(hook.get("O25"), -200.0))

    def test_config_is_signed(self):
        """配置层面：投资收益必须是 signed、科目 6111、第 25 行。"""
        spec = next(s for s in ROW_SPECS if s.key == "投资收益")

        self.assertEqual(spec.direction, "signed")
        self.assertEqual(spec.codes, ("6111",))
        self.assertEqual(spec.excel_row, 25)
        self.assertEqual(spec.excel_row, EXCEL_ROW[17])


class TestDisposalGain(unittest.TestCase):
    """需求第 12 条第 2 项：行次 21「资产处置收益」6115 → C29:N29 + O29。"""

    def test_debit_side_positive(self):
        """6115 出现在借方 → 正数（处置收益）。"""
        hook, _ = run_report([("6115", 1, 400.0, 0.0)], month=1)

        self.assertTrue(approx(hook.get("C29"), 400.0),
                        f"C29（处置收益）应为 +400，实际 {hook.get('C29')}")
        self.assertTrue(approx(hook.get("O29"), 400.0))

    def test_credit_side_negative(self):
        """6115 出现在贷方 → 负数（处置损失）。"""
        hook, _ = run_report([("6115", 1, 0.0, 150.0)], month=1)

        self.assertTrue(approx(hook.get("C29"), -150.0),
                        f"C29（处置损失）应为 -150，实际 {hook.get('C29')}")
        self.assertTrue(approx(hook.get("O29"), -150.0))

    def test_config_is_signed(self):
        """配置层面：资产处置收益必须是 signed、科目 6115、第 29 行。"""
        spec = next(s for s in ROW_SPECS if s.key == "资产处置收益")

        self.assertEqual(spec.direction, "signed")
        self.assertEqual(spec.codes, ("6115",))
        self.assertEqual(spec.excel_row, 29)
        self.assertEqual(spec.excel_row, EXCEL_ROW[21])


class TestSignedRowsTogether(unittest.TestCase):
    """三行 signed（6117/6111/6115）的相互关系。"""

    def test_three_rows_not_crossed(self):
        """三行互不串用（给数量级差异明显的金额）。"""
        rows = [
            ("6117", 3, 1000.0, 0.0),   # 其他收益
            ("6111", 3, 100.0, 0.0),    # 投资收益
            ("6115", 3, 1.0, 0.0),      # 资产处置收益
        ]
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("E24"), 1000.0),
                        f"E24（其他收益）应为 1000，实际 {hook.get('E24')}")
        self.assertTrue(approx(hook.get("E25"), 100.0),
                        f"E25（投资收益）应为 100，实际 {hook.get('E25')}")
        self.assertTrue(approx(hook.get("E29"), 1.0),
                        f"E29（资产处置收益）应为 1，实际 {hook.get('E29')}")

    def test_merged_into_one_query(self):
        """三行口径相同，应合并为 1 条双列查询。"""
        _, ctx = run_report([])

        # 【只统计本年】20261106 起会额外查一次上年（SQL 结构相同）
        signed_calls = [
            c for c in ctx.fetch_calls
            if not c["report_code"].endswith("-上年")
            and "debit_amount" in c["sql"] and "credit_amount" in c["sql"]
        ]
        self.assertEqual(
            len(signed_calls), 1,
            f"三行 signed 应合并为 1 条双列查询，"
            f"实际 {len(signed_calls)} 条")

        params = signed_calls[0]["params"]
        for code in ("6111", "6115", "6117"):
            self.assertIn(code, params, f"查询应包含科目 {code}")

    def test_twelve_months_mixed_signs(self):
        """两行各 12 月，混合借贷符号正确、O 列累计正确。

        奇数月借方 100、偶数月贷方 50 → 累计 6*100 - 6*50 = 300
        """
        rows = []
        for m in range(1, 13):
            if m % 2 == 1:
                rows.append(("6111", m, 100.0, 0.0))
                rows.append(("6115", m, 100.0, 0.0))
            else:
                rows.append(("6111", m, 0.0, 50.0))
                rows.append(("6115", m, 0.0, 50.0))

        hook, _ = run_report(rows, month=12)

        for month, col in COL_OF_MONTH.items():
            expect = 100.0 if month % 2 == 1 else -50.0
            self.assertTrue(approx(hook.get(f"{col}25"), expect),
                            f"{col}25（{month}月）应为 {expect}，"
                            f"实际 {hook.get(f'{col}25')}")
            self.assertTrue(approx(hook.get(f"{col}29"), expect),
                            f"{col}29（{month}月）应为 {expect}，"
                            f"实际 {hook.get(f'{col}29')}")

        self.assertTrue(approx(hook.get("O25"), 300.0))
        self.assertTrue(approx(hook.get("O29"), 300.0))

    def test_blank_when_no_data(self):
        """无数据时两行整行空白（不写 0）。"""
        hook, _ = run_report([], month=1)

        for row in (25, 29):
            for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
                self.assertIsNone(hook.get(f"{col}{row}"),
                                  f"{col}{row} 无数据时应空白，"
                                  f"实际 {hook.get(f'{col}{row}')}")


class TestSkippedRows(unittest.TestCase):
    """第 26/27/28 行（需求未提及）必须保持空白。"""

    def test_rows_26_27_28_blank(self):
        """即使给了可能被误配的科目数据，这三行仍必须空白。

        【为什么锁这条？】
            资产处置收益在第 29 行，中间跳过 26/27/28 三行：
                26 公允价值变动收益
                27 信用减值损失
                28 资产减值损失
            这三行本次需求未提及，按"未提及即不做"处理。
            若日后有人"顺手补上"却没确认口径，会静默产出错数。
        """
        rows = [
            ("6111", 3, 100.0, 0.0),
            ("6115", 3, 100.0, 0.0),
            ("6117", 3, 100.0, 0.0),
            # 下面这些是 26/27/28 行可能涉及的科目，不应被写入
            ("6101", 3, 999.0, 0.0),   # 公允价值变动损益（推测）
            ("6702", 3, 999.0, 0.0),   # 资产减值损失（推测）
            ("6701", 3, 999.0, 0.0),   # 营业外收入（推测）
        ]
        hook, _ = run_report(rows, month=3)

        for row in (26, 27, 28):
            for col in COL_OF_MONTH.values():
                self.assertIsNone(hook.get(f"{col}{row}"),
                                  f"{col}{row} 需求未提及，应保持空白，"
                                  f"实际 {hook.get(f'{col}{row}')}")


if __name__ == "__main__":
    unittest.main(verbosity=2)