# -*- coding: utf-8 -*-
"""LK02 上年同期功能测试（20261106 新增需求）。

【本需求要解决的问题】
    本年利润表的 P 列表头是"上年同期累计数"，必须真的去查上一年度的账，
    不能靠本年数据推算。同时要生成一张"上年对比页"供维护核查，
    平时隐藏、用快捷键切换。

【本文件测什么】
    ① 本年页 P 列 = 上年同行的 O 列（1 月至所选月份的合计）
    ② P 列与 O 列不串（O 用本年、P 用上年），这是最易犯的错
    ③ 上年对比页：C4 写上年年月、月列填上年数、**不填 P 列**
    ④ 该账套首年（无上年库）时：P 列与对比页都留空，且不报错
    ⑤ 取数层面：上年那次查询用 year-1 与 prev_database

【为什么这些必须单独测？】
    本年与上年两组数据都是 float、都填进同样形状的格子。
    若 P 列误取成本年数据，报表上看不出任何异常 ——
    P 列照样"有数字"，只是它与 O 列一模一样而已。
    而这种"两列相同"在人工核查时极易被忽略（看着像是巧合）。
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from tests.fakes import PREV_DB, TARGET_DB, FakeHook, approx  # noqa: E402
from tests.test_lk02_logic import run_report  # noqa: E402

from reports.lk02_profit import (     # noqa: E402
    COL_LAST_YEAR_TOTAL,
    COL_OF_MONTH,
    COL_YEAR_TOTAL,
    EXCEL_ROW,
    LK02Report,
    PERIOD_CELL,
)


def run_with_prev(cur_rows, prev_rows, month: int = 5, year: int = 2026):
    """同时给本年与上年数据，跑一遍 fetch + write。"""
    return run_report(cur_rows, month=month, year=year,
                      prev_rows=prev_rows, prev_database=PREV_DB)


class TestPrevYearColumn(unittest.TestCase):
    """本年页 P 列（上年同期累计数）。"""

    def test_p_equals_prev_year_total(self):
        """P 列 = 上年 1 月至所选月份的合计。

        上年 1~5 月各 100 → P10 = 500（本月选 5 月）。
        """
        prev = [("600102", m, 100.0) for m in range(1, 6)]
        hook, _ = run_with_prev([], prev, month=5)

        self.assertTrue(approx(hook.get("P10"), 500.0),
                        f"P10（上年同期累计）应为 500，"
                        f"实际 {hook.get('P10')}")

    def test_p_not_equal_to_current_year(self):
        """P 列必须取上年，不能取成本年（这是本功能最大的风险点）。

        本年 1~5 月各 1000、上年 1~5 月各 7：
            正确：O10 = 5000（本年）、P10 = 35（上年）
            错误：P10 = 5000（误用本年数据）

        这两列在报表上都"有数字"，肉眼分不出对错，
        所以必须用数量级明显不同的假数据强制区分。
        """
        cur = [("600102", m, 1000.0) for m in range(1, 6)]
        prev = [("600102", m, 7.0) for m in range(1, 6)]
        hook, _ = run_with_prev(cur, prev, month=5)

        self.assertTrue(approx(hook.get("O10"), 5000.0),
                        f"O10（本年累计）应为 5000，实际 {hook.get('O10')}")
        self.assertTrue(approx(hook.get("P10"), 35.0),
                        f"P10（上年同期）应为 35，实际 {hook.get('P10')} —— "
                        f"若等于 5000 说明误用了本年数据")

    def test_p_blank_when_prev_row_has_no_data(self):
        """上年该行无数据时，P 列留空（不写 0）。

        【为什么不写 0？】
            0 会被读成"上年该项确为零"，而留空表示"上年没有这项数据"。
            两者在做同比分析时结论完全不同。
        """
        # 上年只有内销(600101)，没有出口(600102)
        prev = [("600101", m, 50.0) for m in range(1, 6)]
        hook, _ = run_with_prev([], prev, month=5)

        self.assertIsNone(hook.get("P10"),
                          f"出口行上年无数据，P10 应留空，"
                          f"实际 {hook.get('P10')}")
        self.assertTrue(approx(hook.get("P11"), 250.0),
                        f"内销行 P11 应为 250，实际 {hook.get('P11')}")

    def test_p_filled_even_when_current_year_blank(self):
        """本年无数据但上年有 → P 列仍须填。

        【场景】本年还没做结转，但上年同期有数。
            这正是"同比"最有价值的时刻：让财务看到"去年这时候有收入，
            今年还没有"。若因为本行无数据就跳过 P 列，
            这份对比信息就丢了。
        """
        prev = [("600102", 1, 888.0)]
        hook, _ = run_with_prev([], prev, month=1)

        self.assertIsNone(hook.get("C10"), "本年无数据，C10 应留空")
        self.assertTrue(approx(hook.get("P10"), 888.0),
                        f"上年有数据，P10 应为 888，实际 {hook.get('P10')}")

    def test_p_respects_selected_month(self):
        """P 列只累计到所选月份，不能把上年全年都算进来。

        上年 1~12 月各 100：
            选 3 月 → P10 = 300
            选 12 月 → P10 = 1200
        """
        prev = [("600102", m, 100.0) for m in range(1, 13)]

        hook3, _ = run_with_prev([], prev, month=3)
        self.assertTrue(approx(hook3.get("P10"), 300.0),
                        f"选 3 月时 P10 应为 300，实际 {hook3.get('P10')}")

        hook12, _ = run_with_prev([], prev, month=12)
        self.assertTrue(approx(hook12.get("P10"), 1200.0),
                        f"选 12 月时 P10 应为 1200，实际 {hook12.get('P10')}")


class TestNoPrevDatabase(unittest.TestCase):
    """该账套首年（无上年库）时的降级行为。"""

    def test_p_blank_when_no_prev_database(self):
        """prev_database 为空 → P 列全部留空。"""
        hook, ctx = run_report([("600102", 1, 100.0)], month=1,
                               prev_database="")

        for spec_row in {s.excel_row for s in _all_rows()}:
            self.assertIsNone(
                hook.get(f"{COL_LAST_YEAR_TOTAL}{spec_row}"),
                f"无上年库时 P{spec_row} 应留空")

    def test_no_prev_query_when_no_prev_database(self):
        """无上年库时**不应发**上年查询（省一次无意义的往返）。"""
        _, ctx = run_report([], month=1, prev_database="")

        prev_calls = [c for c in ctx.fetch_calls
                      if c["report_code"].endswith("-上年")]
        self.assertEqual(len(prev_calls), 0,
                         f"无上年库时不应查上年，实际 {len(prev_calls)} 次")

    def test_current_year_unaffected(self):
        """无上年库不影响本年数据。"""
        hook, _ = run_report([("600102", 1, 777.0)], month=1,
                             prev_database="")

        self.assertTrue(approx(hook.get("C10"), 777.0))
        self.assertTrue(approx(hook.get("O10"), 777.0))


class TestQueryUsesPrevYear(unittest.TestCase):
    """取数层面：上年那次查询的年度与库必须是上年的。"""

    def test_query_year_is_previous(self):
        """上年查询的参数里，年度应为 year-1。"""
        _, ctx = run_report([], month=5, year=2026,
                            prev_rows=[], prev_database=PREV_DB)

        prev_calls = [c for c in ctx.fetch_calls
                      if c["report_code"].endswith("-上年")]
        self.assertTrue(prev_calls, "应有上年查询")

        params = prev_calls[0]["params"]
        self.assertEqual(params[0], 2025,
                         f"上年查询年度应为 2025，实际 {params[0]}")
        self.assertEqual(params[1], 202501, "上年区间下限应为 202501")
        self.assertEqual(params[2], 202505, "上年区间上限应为 202505")

    def test_query_uses_prev_database(self):
        """上年查询必须显式指定 prev_database。"""
        _, ctx = run_report([], month=5, prev_rows=[], prev_database=PREV_DB)

        prev_calls = [c for c in ctx.fetch_calls
                      if c["report_code"].endswith("-上年")]
        self.assertTrue(prev_calls)
        for c in prev_calls:
            self.assertEqual(c["database"], PREV_DB,
                             f"上年查询应指定库 {PREV_DB}，"
                             f"实际 {c['database']}")

    def test_current_query_uses_target_database(self):
        """本年查询必须显式指定 target_database（不能被上年带偏）。"""
        _, ctx = run_report([], month=5, prev_rows=[], prev_database=PREV_DB)

        cur_calls = [c for c in ctx.fetch_calls
                     if not c["report_code"].endswith("-上年")]
        self.assertTrue(cur_calls)
        for c in cur_calls:
            self.assertEqual(c["database"], TARGET_DB,
                             f"本年查询应指定库 {TARGET_DB}，"
                             f"实际 {c['database']}")


class TestPrevSheet(unittest.TestCase):
    """上年对比页的写入（由框架复制表页后调用 write_prev_sheet）。"""

    def _make(self, cur_rows, prev_rows, month=5, year=2026):
        # 复用 run_with_prev 的取数结果，再单独写一张"上年页"
        _, ctx = run_report(cur_rows, month=month, year=year,
                            prev_rows=prev_rows, prev_database=PREV_DB)
        rep = LK02Report()
        rep.fetch(ctx)
        hook = FakeHook()
        rep.write_prev_sheet(hook, ctx)
        return hook

    def test_period_cell_is_previous_year(self):
        """对比页 C4 必须写**上年**年月。

        【为什么这条重要？】
            对比页是复制本年页得到的，C4 里原本是本年年月。
            若忘记改写，打开这张页会看到"2026年5月"却填着 2025 年的数 ——
            维护者会以为数据错了，而实际上是标题没改。
        """
        hook = self._make([], [("600102", 1, 100.0)], month=5, year=2026)
        self.assertEqual(hook.get(PERIOD_CELL), "2025年5月",
                         f"对比页 C4 应为 2025年5月，"
                         f"实际 {hook.get(PERIOD_CELL)}")

    def test_month_columns_use_prev_data(self):
        """对比页的月列填上年数据。"""
        prev = [("600102", m, 100.0) for m in range(1, 6)]
        hook = self._make([("600102", m, 9999.0) for m in range(1, 6)],
                          prev, month=5)

        for month, col in COL_OF_MONTH.items():
            if month > 5:
                continue
            self.assertTrue(approx(hook.get(f"{col}10"), 100.0),
                            f"{col}10（{month}月）应填上年 100，"
                            f"实际 {hook.get(f'{col}10')}")

    def test_year_total_is_prev_total(self):
        """对比页 O 列 = 上年累计。"""
        prev = [("600102", m, 100.0) for m in range(1, 6)]
        hook = self._make([], prev, month=5)

        self.assertTrue(approx(hook.get(f"{COL_YEAR_TOTAL}10"), 500.0),
                        f"对比页 O10 应为 500，"
                        f"实际 {hook.get(f'{COL_YEAR_TOTAL}10')}")

    def test_p_column_cleared(self):
        """对比页 **不得**有 P 列数字（它没有"上年同期"可比）。

        【为什么必须主动清空？】
            对比页是复制本年页得来的，P 列里还留着本年算出的
            "上年同期累计数"。若不清理，展开对比页会看到：
                上年页的 P 列 = 本年的上年同期数
            而这一页本身就是上年 —— "上年的上年同期"是前年，
            值却显示着今年的计算结果，自相矛盾。
        """
        prev = [("600102", m, 100.0) for m in range(1, 6)]
        hook = self._make([("600102", m, 1000.0) for m in range(1, 6)],
                          prev, month=5)

        for row in {s.excel_row for s in _all_rows()}:
            cell = f"{COL_LAST_YEAR_TOTAL}{row}"
            val = hook.get(cell)
            self.assertIn(val, (None, ""),
                          f"对比页 {cell} 应被清空，实际 {val!r}")

    def test_no_prev_data_still_generates_with_marker(self):
        """上年无数据时，仍生成对比页，并在 C4 明确标注。

        【为什么必须生成，而不是跳过？（20261106 线上反馈修正）】
            早期实现是"无数据就跳过生成"，理由是别给用户空白页。
            但实际使用后维护者分不清两种情况：
              · 功能没生效（页根本没生成）
              · 上年确实没数据（页生成了但是空的）
            这两者的处理方式完全不同：前者查程序，后者查 U8 账。

            所以改为总是生成，并在 C4 明确标注。一张标注清楚的空页
            就是"程序查过了、确实没有"的证据 —— 这比一张不存在的页
            信息量大得多。

        【为什么金额格必须留空而不是写 0】
            0 表示"上年该项确为零"，留空表示"没有这项数据"。
            做同比分析时这是两个完全不同的结论。
        """
        _, ctx = run_report([], month=5, prev_rows=None,
                            prev_database=PREV_DB)
        rep = LK02Report()
        rep.fetch(ctx)
        hook = FakeHook()
        rep.write_prev_sheet(hook, ctx)

        self.assertEqual(hook.get(PERIOD_CELL), "2025年5月（上年同期无数据）",
                         f"无上年数据时 C4 应带标注，"
                         f"实际 {hook.get(PERIOD_CELL)!r}")
        for row in {s.excel_row for s in _all_rows()}:
            for col in COL_OF_MONTH.values():
                self.assertIsNone(hook.get(f"{col}{row}"),
                                  f"{col}{row} 无上年数据时应留空（不写 0）")


class TestPrevSheetDeclared(unittest.TestCase):
    """能力声明本身。"""

    def test_suffix_declared(self):
        """LK02 必须声明 prev_sheet_suffix，否则框架不会生成对比页。"""
        self.assertEqual(LK02Report.prev_sheet_suffix, "(上年)")

    def test_other_reports_default_empty(self):
        """基类默认为空，即"不生成对比页"（其他表不受影响）。"""
        from reports.base import BaseReport
        self.assertEqual(BaseReport.prev_sheet_suffix, "")


def _all_rows():
    """本表所有 ROW_SPECS（延迟导入，避免模块顶层循环）。"""
    from reports.lk02_profit import ROW_SPECS
    return ROW_SPECS


if __name__ == "__main__":
    unittest.main(verbosity=2)