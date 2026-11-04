# -*- coding: utf-8 -*-
"""LK02 需求第 15 条的独立测试文件：两行费用类项目。

【需求原文（20261104 现场确认）】
    15、C32 营业外支出 6711 和 C34 所得税 6801
        取数逻辑与 C16 其他业务成本的类同

【"类同"具体指什么？】
    C16（第 7 条 其他业务成本 6402）的口径是：
        ① ccode = 指定科目
        ② 摘要含「损益结转」/「结转损益」
        ③ 对方科目 = 4103（本年利润）
        ④ 取 mc（贷方金额）合计
        ⑤ 填 12 个月列 + O 列本年累计
    这两行全部照搬，只有科目与格位不同。

【方向为什么是贷方？（与第 14 条对比）】
    营业外收入 6301 是收入类 → 借方转出 → 取 md
    营业外支出 6711 / 所得税费用 6801 是费用类 → 贷方转出 → 取 mc
    对方科目同为 4103，方向却相反。本文件用测试把这一对钉死，
    防止"income/expense 一律取 md"这类错误。

【测什么】
    ① 配置：科目、方向 credit、格位、摘要/对方科目
    ② 与 C16 逐字段比对（把"类同"变成机器可验证的约束）
    ③ 与第 14 条营业外收入的方向必须相反
    ④ 行为：贷方金额落格、12 月累计、缺月留空
    ⑤ 第 34 行不是公式行（必须真的被写入）
    ⑥ 边界：不越界写第 33/35 行（公式行）、不串行
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
    DIGEST_SETTLE,
    EQUAL_PROFIT,
    EXCEL_ROW,
    FORMULA_ROWS,
    ROW_SPECS,
)


def _spec(key: str):
    """按 key 取配置，并断言只有一条（防复制粘贴漏改 key）。"""
    specs = [s for s in ROW_SPECS if s.key == key]
    assert len(specs) == 1, f"{key} 应只配置 1 条，实际 {len(specs)} 条"
    return specs[0]


# 本需求涉及的两行：(key, 科目, Excel 行, 模板行次)
CASES = (
    ("营业外支出", "6711", 32, 24),
    ("所得税费用", "6801", 34, 26),
)


class TestConfig(unittest.TestCase):
    """配置校验：两行的科目/方向/格位/口径。"""

    def test_codes(self):
        """科目分别为 6711 与 6801。"""
        for key, code, _row, _no in CASES:
            self.assertEqual(_spec(key).codes, (code,),
                             f"{key} 科目应为 {code}")

    def test_direction_credit(self):
        """方向必须是 credit（与 C16 相同，取贷方 mc）。

        【为什么这条最关键？】
            这两行是费用类，直觉上像"借方科目"，
            但本表取的是结转发生额，费用在贷方转出。
            若写成 debit，C32/C34 永远空白且不报任何错 ——
            利润总额虚高、净利润虚高，报表上看不出异常。
        """
        for key, _code, _row, _no in CASES:
            self.assertEqual(_spec(key).direction, "credit",
                             f"{key} 应取贷方 mc（与 C16 类同）")

    def test_excel_row(self):
        """格位分别是第 32 行与第 34 行。"""
        for key, _code, row, no in CASES:
            spec = _spec(key)
            self.assertEqual(spec.excel_row, row, f"{key} 应在第 {row} 行")
            self.assertEqual(EXCEL_ROW[no], row,
                             f"EXCEL_ROW[{no}] 应为 {row}")

    def test_digest_and_equal_codes(self):
        """摘要关键字与对方科目必须与 C16 完全一致。"""
        for key, _code, _row, _no in CASES:
            spec = _spec(key)
            self.assertEqual(spec.digest_keywords, DIGEST_SETTLE)
            self.assertEqual(spec.equal_codes, (EQUAL_PROFIT,))

    def test_no_exclude_no_fallback_not_optional(self):
        """无减项、无备用科目、非 optional、填 O 列。"""
        for key, _code, _row, _no in CASES:
            spec = _spec(key)
            self.assertEqual(spec.exclude_codes, ())
            self.assertEqual(spec.fallback_codes, ())
            self.assertFalse(spec.optional,
                             f"{key} 需求未提暂未启用，应按普通行处理")
            self.assertTrue(spec.fill_total)

    def test_matches_c16_except_code_and_row(self):
        """与 C16（其他业务成本）逐字段比对，除科目与行号必须全同。

        【这是"类同"的机器化验证】
            以后若有人改了 C16 的摘要口径或方向却忘了同步这两行，
            本测试会立即失败。
        """
        c16 = _spec("其他业务成本")
        for key, _code, _row, _no in CASES:
            cur = _spec(key)
            self.assertEqual(cur.direction, c16.direction,
                             f"{key} 方向应与 C16 一致")
            self.assertEqual(cur.digest_keywords, c16.digest_keywords,
                             f"{key} 摘要筛选应与 C16 一致")
            self.assertEqual(cur.equal_codes, c16.equal_codes,
                             f"{key} 对方科目应与 C16 一致")
            self.assertEqual(cur.fill_total, c16.fill_total)
            self.assertEqual(cur.optional, c16.optional)
            self.assertEqual(cur.exclude_codes, c16.exclude_codes)
            self.assertEqual(cur.fallback_codes, c16.fallback_codes)
            self.assertNotEqual(cur.codes, c16.codes, "科目必须不同")
            self.assertNotEqual(cur.excel_row, c16.excel_row, "格位必须不同")

    def test_opposite_direction_to_req14(self):
        """与第 14 条营业外收入的方向必须相反。

        【为什么把这条单列？】
            营业外收入/支出在模板上紧挨着（31/32 行），
            且都涉及"营业外"，最容易被复制粘贴成同一方向。
            事实是：#14 收入取 md、#15 支出取 mc。
            若两行方向相同，必有一行取不到数（且不报错）。
        """
        income = _spec("营业外收入")
        expense = _spec("营业外支出")

        self.assertEqual(income.direction, "debit")
        self.assertEqual(expense.direction, "credit")
        self.assertNotEqual(income.direction, expense.direction,
                            "营业外收入取借方、支出取贷方，方向必须相反")
        self.assertEqual(income.equal_codes, expense.equal_codes,
                         "两行对方科目同为 4103")


class TestIncomeTaxNotFormulaRow(unittest.TestCase):
    """第 34 行（所得税费用）不是公式行，必须真的被写入。"""

    def test_not_in_formula_rows(self):
        """行次 26 不得出现在 FORMULA_ROWS 中。

        【为什么单独测？】
            第 34 行紧邻第 33（利润总额，公式）与第 35（净利润，公式），
            很容易被"顺手"一并当作公式行登记。
            一旦如此，所得税永远不会被写，而模板公式照常计算
            净利润 = 33行 － 34行，34 行为空 → 净利润虚高，
            且报表上无任何异常提示。
        """
        self.assertNotIn(26, FORMULA_ROWS,
                         "行次 26（所得税费用）是明细行，不应登记为公式行")

    def test_formula_rows_unchanged(self):
        """模板公式行仍只有 22/25/27。"""
        self.assertEqual(set(FORMULA_ROWS), {22, 25, 27})

    def test_income_tax_actually_written(self):
        """给 6801 数据，第 34 行必须被写入。"""
        hook, _ = run_report([("6801", 3, 2500.0)], month=3)

        self.assertTrue(approx(hook.get("E34"), 2500.0),
                        f"E34 应为 2500，实际 {hook.get('E34')}")
        self.assertTrue(approx(hook.get("O34"), 2500.0))


class TestValues(unittest.TestCase):
    """行为校验：贷方金额落格。"""

    def test_expense_single_month(self):
        """1 月 6711 贷方 300 → C32=300、O32=300。"""
        hook, _ = run_report([("6711", 1, 300.0)], month=1)

        self.assertTrue(approx(hook.get("C32"), 300.0),
                        f"C32 应为 300，实际 {hook.get('C32')}")
        self.assertTrue(approx(hook.get("O32"), 300.0))

    def test_twelve_months_and_total(self):
        """两行各 12 个月，每月 100 → O32=O34=1200。"""
        rows = []
        for m in range(1, 13):
            rows.append(("6711", m, 100.0))
            rows.append(("6801", m, 100.0))
        hook, _ = run_report(rows, month=12)

        for month, col in COL_OF_MONTH.items():
            self.assertTrue(approx(hook.get(f"{col}32"), 100.0),
                            f"{col}32（{month}月）应为 100")
            self.assertTrue(approx(hook.get(f"{col}34"), 100.0),
                            f"{col}34（{month}月）应为 100")

        self.assertTrue(approx(hook.get("O32"), 1200.0))
        self.assertTrue(approx(hook.get("O34"), 1200.0))

    def test_gap_months_left_blank(self):
        """1、3 月有数、2 月无 → 2 月留空，累计只加有数月份。"""
        rows = [("6711", 1, 100.0), ("6711", 3, 200.0)]
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("C32"), 100.0))
        self.assertIsNone(hook.get("D32"), "2 月无数据应留空")
        self.assertTrue(approx(hook.get("E32"), 200.0))
        self.assertTrue(approx(hook.get("O32"), 300.0))

    def test_negative_amount_reported_as_is(self):
        """credit 模式对红字（负数）原样填列，不做符号翻转。

        【为什么用负数代替原来的"只出现在借方"用例？】
            credit 模式走单列函数，返回的 amount 就是 SUM(mc)。
            "该科目只出现在借方"这件事发生在 SQL 层（SUM(mc) 取不到），
            测试层造数时无法表达 —— 假数据里 amount 就是最终结果，
            它不区分来自哪一列。

            所以改为验证本层能决定的事：若贷方合计为负
            （红字冲回/调账），程序必须原样填入负数，
            不能"顺手取绝对值"或"当成异常丢弃"。
            若被静默改动，报表面上的支出会由负变正，方向失真。
        """
        hook, _ = run_report([("6711", 3, -500.0)], month=3)

        self.assertTrue(approx(hook.get("E32"), -500.0),
                        f"E32 应原样填 -500，实际 {hook.get('E32')}")
        self.assertTrue(approx(hook.get("O32"), -500.0))

class TestQuery(unittest.TestCase):
    """取数路径校验。"""

    def test_uses_credit_sum(self):
        """两行所在查询应取 SUM(mc)，与 C16 同组。"""
        _, ctx = run_report([])

        credit_calls = [c for c in ctx.fetch_calls
                        if "SUM(mc) AS amount" in c["sql"]]
        self.assertTrue(credit_calls, "应存在取贷方的查询")

        hit = [c for c in credit_calls
               if "6711" in (c["params"] or ())]
        self.assertTrue(hit, "6711 应出现在取贷方的查询参数中")

        hit_tax = [c for c in credit_calls
                   if "6801" in (c["params"] or ())]
        self.assertTrue(hit_tax, "6801 应出现在取贷方的查询参数中")

    def test_merged_with_cost_rows(self):
        """与 C16/C17/C18 等费用行合并为同一次 credit 查询。"""
        _, ctx = run_report([])

        credit_calls = [c for c in ctx.fetch_calls
                        if "SUM(mc) AS amount" in c["sql"]]
        hit = [c for c in credit_calls
               if "6711" in (c["params"] or ())]
        self.assertEqual(len(hit), 1, "6711 应恰好出现在 1 次 credit 查询中")

        params = tuple(hit[0]["params"] or ())
        for code in ("6402", "6403", "6601", "6711", "6801"):
            self.assertIn(code, params,
                          f"费用类科目 {code} 应与 6711 同组查出")

    def test_not_in_debit_group(self):
        """两行不得混入取借方的查询（那是收入行的组）。"""
        _, ctx = run_report([])

        debit_calls = [c for c in ctx.fetch_calls
                       if "SUM(md) AS amount" in c["sql"]]
        for c in debit_calls:
            params = tuple(c["params"] or ())
            self.assertNotIn("6711", params, "6711 不应出现在借方查询里")
            self.assertNotIn("6801", params, "6801 不应出现在借方查询里")


class TestBoundary(unittest.TestCase):
    """边界：不越界写入公式行与其他行。"""

    def test_no_data_keeps_blank(self):
        """整年无数据 → 整行空白（不写 0）。"""
        hook, _ = run_report([], month=1)

        for _key, _code, row, _no in CASES:
            for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
                cell = f"{col}{row}"
                self.assertIsNone(hook.get(cell),
                                  f"{cell} 无数据时应空白，"
                                  f"实际 {hook.get(cell)}")

    def test_formula_rows_not_written(self):
        """第 30/33/35 行是公式行，程序不得写入。"""
        rows = [("6711", 3, 0.0, 100.0), ("6801", 3, 0.0, 50.0)]
        hook, _ = run_report(rows, month=3)

        for row in (30, 33, 35):
            for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
                cell = f"{col}{row}"
                self.assertIsNone(hook.get(cell),
                                  f"{cell} 属公式行，不应被写入，"
                                  f"实际 {hook.get(cell)}")

    def test_rows_not_crossed(self):
        """6711 与 6801 的数字不得互相串行，也不得串到第 31 行。"""
        rows = [("6711", 3, 111.0), ("6801", 3, 222.0)]
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("E32"), 111.0),
                        f"E32 应为 111，实际 {hook.get('E32')}")
        self.assertTrue(approx(hook.get("E34"), 222.0),
                        f"E34 应为 222，实际 {hook.get('E34')}")
        self.assertIsNone(hook.get("E31"), "E31（营业外收入）不应被写入")


if __name__ == "__main__":
    unittest.main(verbosity=2)