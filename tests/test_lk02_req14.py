# -*- coding: utf-8 -*-
"""LK02 需求第 14 条的独立测试文件：营业外收入 6301 → C31:N31 + O31。

【需求原文（20261104 现场确认）】
    14、C31 营业外收入 6301 取数逻辑与 C10 主营业务收入的类同

【"类同"到底类同在哪？】
    C10（第 2 条 出口主营业务收入 600102）的口径是：
        ① ccode = 指定科目
        ② 摘要含「损益结转」/「结转损益」
        ③ 对方科目 = 4103（本年利润）
        ④ 取 md（借方金额）合计
        ⑤ 填 12 个月列 + O 列本年累计
    本行这五点全部照搬，只有科目（6301）与格位（第 31 行）不同。
    所以本文件除了验 6301/第31行，还要验这几条口径确实一致 ——
    否则"类同"就只是口头说说。

【测什么】
    ① 配置：科目 6301、方向 debit、格位第 31 行、摘要/对方科目口径
    ② 行为：借方金额 → C31 正数；本年累计 O31 正确
    ③ 取数：走 SUM(md) 单列查询（不是 signed 双列）
    ④ 合并：与其它收入行同组，只发一次 debit 查询
    ⑤ 空数：整年无数据时整行空白（不写 0）
    ⑥ 边界：不得越界写入第 32 行（营业外支出，需求未给）
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
    ROW_SPECS,
)


def _spec():
    """取出营业外收入的配置项，顺带断言它只配置了一条。

    【为什么断言条数？】
        若日后有人复制粘贴时漏改 key，会出现两条同名配置，
        next() 只会拿到第一条，第二条的错配置就测不到。
    """
    specs = [s for s in ROW_SPECS if s.key == "营业外收入"]
    assert len(specs) == 1, f"营业外收入应只配置 1 条，实际 {len(specs)} 条"
    return specs[0]


class TestConfig(unittest.TestCase):
    """配置校验：口径必须与 C10（出口主营业务收入）一致。"""

    def test_codes(self):
        """科目必须是 6301。"""
        self.assertEqual(_spec().codes, ("6301",))

    def test_direction_debit(self):
        """方向必须是 debit（与 C10 相同，取借方 md）。

        【为什么这条最关键？】
            营业外收入是收入类，直觉上像"贷方科目"，
            但本表取的是结转发生额，收入在借方转出。
            若写成 credit，C31 永远空白，且不报任何错 ——
            报表看起来像"整年没有营业外收入"。
        """
        self.assertEqual(_spec().direction, "debit",
                         "营业外收入应取借方 md（与 C10 类同）")

    def test_excel_row(self):
        """格位必须是 Excel 第 31 行（模板行次 23）。"""
        self.assertEqual(_spec().excel_row, 31)
        self.assertEqual(_spec().excel_row, EXCEL_ROW[23],
                         "EXCEL_ROW[23] 应为 31")

    def test_digest_and_equal_codes(self):
        """摘要关键字与对方科目必须与 C10 完全一致。"""
        spec = _spec()
        self.assertEqual(spec.digest_keywords, DIGEST_SETTLE,
                         "摘要筛选应与 C10 一致")
        self.assertEqual(spec.equal_codes, (EQUAL_PROFIT,),
                         "对方科目应限定为 4103（本年利润）")

    def test_no_exclude_no_fallback_no_optional(self):
        """需求没有减项、没有备用科目、没有"暂未启用"的说明。"""
        spec = _spec()
        self.assertEqual(spec.exclude_codes, ())
        self.assertEqual(spec.fallback_codes, ())
        self.assertFalse(spec.optional,
                         "需求未提暂未启用，应按普通行处理")
        self.assertTrue(spec.fill_total, "应填 O 列本年累计")

    def test_matches_c10_except_code_and_row(self):
        """与 C10 的口径字段逐一比对 —— 除科目与行号外必须全同。

        【这是"类同"的机器化验证】
            逐字段比较而非人工读注释：
            以后若有人改了 C10 的摘要口径却忘了同步本行，
            本测试会立即失败。
        """
        c10 = next(s for s in ROW_SPECS if s.key == "出口主营业务收入")
        cur = _spec()

        self.assertEqual(cur.direction, c10.direction)
        self.assertEqual(cur.digest_keywords, c10.digest_keywords)
        self.assertEqual(cur.equal_codes, c10.equal_codes)
        self.assertEqual(cur.fill_total, c10.fill_total)
        self.assertNotEqual(cur.codes, c10.codes, "科目必须不同")
        self.assertNotEqual(cur.excel_row, c10.excel_row, "格位必须不同")


class TestValues(unittest.TestCase):
    """行为校验：借方金额落到 C31/N31 + O31。"""

    def test_single_month(self):
        """1 月借方 6301 = 500 → C31=500、O31=500。"""
        hook, _ = run_report([("6301", 1, 500.0, 0.0)], month=1)

        self.assertTrue(approx(hook.get("C31"), 500.0),
                        f"C31 应为 500，实际 {hook.get('C31')}")
        self.assertTrue(approx(hook.get("O31"), 500.0),
                        f"O31 应为 500，实际 {hook.get('O31')}")

    def test_twelve_months_and_total(self):
        """12 个月各 100 → 每月格 100、O31 = 1200。"""
        rows = [("6301", m, 100.0, 0.0) for m in range(1, 13)]
        hook, _ = run_report(rows, month=12)

        for month, col in COL_OF_MONTH.items():
            self.assertTrue(approx(hook.get(f"{col}31"), 100.0),
                            f"{col}31（{month}月）应为 100，"
                            f"实际 {hook.get(f'{col}31')}")

        self.assertTrue(approx(hook.get("O31"), 1200.0),
                        f"O31 应为 1200，实际 {hook.get('O31')}")

    def test_gap_months_left_blank(self):
        """有 1、3 月数据、2 月无 → 2 月格留空，累计只加有数月份。

        对应开发口径⑤"月份覆盖不假设连续"。
        """
        rows = [("6301", 1, 100.0, 0.0), ("6301", 3, 200.0, 0.0)]
        hook, _ = run_report(rows, month=3)

        self.assertTrue(approx(hook.get("C31"), 100.0))
        self.assertIsNone(hook.get("D31"), "2 月无数据应留空，不写 0")
        self.assertTrue(approx(hook.get("E31"), 200.0))
        self.assertTrue(approx(hook.get("O31"), 300.0),
                        f"O31 应为 300（只加有数月份），实际 {hook.get('O31')}")


class TestQuery(unittest.TestCase):
    """取数路径校验：应是单列 SUM(md)，且与其它收入行合并。"""

    def test_uses_debit_sum(self):
        """营业外收入所在查询应取 SUM(md)，不是 SUM(mc)、也不是双列。"""
        _, ctx = run_report([])

        debit_calls = [c for c in ctx.fetch_calls
                       if "SUM(md) AS amount" in c["sql"]]
        self.assertTrue(debit_calls, "应存在取借方的查询")

        found = any("6301" in (c["params"] or ()) for c in debit_calls)
        self.assertTrue(found, "6301 应出现在取借方的查询参数中")

    def test_merged_with_other_income_rows(self):
        """与 C10/C11/C12 同口径 → 合并进同一次 debit 查询。

        【为什么要测合并？】
            合并本身是性能优化，但也是"口径一致"的证据：
            若某天本行被误改了摘要或对方科目，分组键就会变，
            本行会被拆到另一次查询里 —— 本测试会失败，
            提示"这一行的口径已经和其他收入行不同了"。
        """
        _, ctx = run_report([])

        debit_calls = [c for c in ctx.fetch_calls
                       if "SUM(md) AS amount" in c["sql"]]
        hit = [c for c in debit_calls
               if "6301" in (c["params"] or ())]
        self.assertEqual(len(hit), 1, "6301 应恰好出现在 1 次 debit 查询中")

        params = tuple(hit[0]["params"] or ())
        for code in ("600101", "600102", "6051", "6301"):
            self.assertIn(code, params,
                          f"收入类科目 {code} 应与 6301 同组查出")

    def test_not_signed_query(self):
        """本行不得走 signed 双列查询（它不属于方向不定类）。"""
        _, ctx = run_report([])

        signed_calls = [c for c in ctx.fetch_calls
                        if "debit_amount" in c["sql"]]
        for c in signed_calls:
            self.assertNotIn("6301", (c["params"] or ()),
                             "6301 不应出现在 signed 查询里")


class TestBoundary(unittest.TestCase):
    """边界：不得越界写入未实现的行。"""

    def test_no_data_keeps_blank(self):
        """整年无数据 → 整行空白（不写 0）。"""
        hook, _ = run_report([], month=1)

        for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
            cell = f"{col}31"
            self.assertIsNone(hook.get(cell),
                              f"{cell} 无数据时应空白，实际 {hook.get(cell)}")

    def test_row_32_untouched(self):
        """第 32 行（营业外支出）需求尚未给出，必须保持空白。

        【为什么测这个？】
            营业外收入/支出在模板里紧挨着（31/32 行），
            复制配置时最容易把 excel_row 抄成 32，
            而报表上第 32 行出现数字看起来也"合理"。
        """
        rows = [("6301", 3, 500.0, 0.0)]
        hook, _ = run_report(rows, month=3)

        for col in (*COL_OF_MONTH.values(), COL_YEAR_TOTAL):
            cell = f"{col}32"
            self.assertIsNone(hook.get(cell),
                              f"{cell} 需求未给出，应保持空白，"
                              f"实际 {hook.get(cell)}")

    def test_does_not_leak_to_other_income_rows(self):
        """6301 的数据不得串到其他收入行（600102/600101/6051）。"""
        rows = [("6301", 3, 777.0, 0.0)]
        hook, _ = run_report(rows, month=3)

        for row in (10, 11, 12):
            self.assertIsNone(hook.get(f"E{row}"),
                              f"E{row} 不应因 6301 有数而被写入，"
                              f"实际 {hook.get(f'E{row}')}")
        self.assertTrue(approx(hook.get("E31"), 777.0))


if __name__ == "__main__":
    unittest.main(verbosity=2)