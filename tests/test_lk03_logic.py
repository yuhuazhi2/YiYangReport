# -*- coding: utf-8 -*-
"""LK03 现金流量表的逻辑测试。

运行方式（在项目根目录执行）：
    python -m unittest discover -s tests -v

【本文件测什么？】
    1. 会计期间是否写入 **B4**（不是 C4），且必须用 set_text
    2. B4 是否只写一次（防止通用页眉叠加写第二处年月）
    3. 【结构护栏】C4 绝不能写 —— 模板里 C4 是「复核：」标签
    4. 【结构护栏】公式行（小计/净额/余额）任何输入下都不得写入
    5. 【结构护栏】分段标题行（行次 1/12/25）不得写入
    6. 【结构护栏】EXCEL_ROW 覆盖 1~38，且与模板 B 列行次逐行相符
    7. 【结构护栏】待定项不占用已实现行、彼此不重复、不落在公式行
    8. 【结构护栏】ROW_SPECS 与 cell_map 一致
    9. 需求第 2 条：行次 2「销售商品、提供劳务收到的现金」→ C10

    第 3~8 条是**长期护栏**：以后逐条追加明细行时若写错，
    测试会立刻失败，不必等生成报表才发现。

【不测什么？】
    真实 Excel 的写入效果（格式、日期误判）—— 那必须用真 Excel，
    属手工验证范畴，见 docs/报表开发规范.md 的说明。
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

# 让测试既能通过 discover 运行，也能单独 python tests/xxx.py 运行
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from tests.fakes import (  # noqa: E402
    TARGET_DB,
    FakeCtx,
    FakeHook,
    approx,
)

from reports.lk03_cashflow import (        # noqa: E402
    BALANCE_SPECS,
    COL_CURRENT,
    COL_PREV,
    EXCEL_ROW,
    FORMULA_ROWS,
    INTENTIONALLY_UNTOUCHED,
    PENDING_ROWS,
    PERIOD_CELL,
    ROW_SPECS,
    SECTION_ROWS,
    LK03Report,
)


def cell_row(cell: str) -> int:
    """把 "C10" 拆出 Excel 行号 10。"""
    return int("".join(ch for ch in cell if ch.isdigit()))


def cell_col(cell: str) -> str:
    """把 "C10" 拆出列字母 "C"。"""
    return cell[0]


def make_cash_rows(items_map: dict, month: int = 5,
                   year: int = 2026) -> list:
    """造 gl_cashtable 的查询结果：[(cCashItem, iYPeriod, md, mc), ...]。

    参数 items_map : {项目编码: (借方md, 贷方mc)}
         year, month : 该批数据所属的年月（决定 iYPeriod）

    【为什么不复用 tests/fakes.py 的 make_rows()？】
        那个造的是 `(ccode, iperiod, mb, me)` —— 余额格式。
        gl_cashtable 返回的是 `(cCashItem, iYPeriod, md, mc)`，
        列数与语义都不同，强行复用会产生"看起来跑通了、
        实际上字段错位"的测试假阳性。
    """
    yperiod = int(f"{year}{month:02d}")
    return [(item, yperiod, md, mc) for item, (md, mc) in items_map.items()]


def make_balance_rows(balances: dict, month: int = 5,
                      year: int = 2026) -> list:
    """造 GL_accsum 余额查询结果：[(ccode, iYPeriod, mb), ...]。

    参数 balances : {科目编码: 期初余额 mb}

    【为什么不与 make_cash_rows 合并？】
        两者的列数与语义完全不同（发生额 vs 余额），
        且来自不同的表。合并会造出"看起来跑通了、
        实际字段错位"的测试假阳性。
    """
    yperiod = int(f"{year}{month:02d}")
    return [(code, yperiod, mb) for code, mb in balances.items()]


def run_report(rows=None, month: int = 5, year: int = 2026,
               fail: bool = False, prev_rows=None,
               prev_database: str = "",
               balance_rows=None, balance_prev_rows=None,
               formula_values=None):
    """跑一遍 LK03 的 fetch + write，返回 (hook, ctx)。

    参数：
        rows          本期现金流量项目查询结果（gl_cashtable）
        prev_rows     上期现金流量项目查询结果；None = 上期无数据
        prev_database 上一年度业务库。传 "" 表示"该账套首年"，
                      此时 LK03 在 1 月不会去查上期（也查不到）。
        balance_rows      本期科目余额查询结果（GL_accsum，需求31）
        balance_prev_rows 上期科目余额查询结果
    """
    ctx = FakeCtx(rows or [], year=year, month=month, fail=fail,
                  prev_rows=prev_rows, prev_database=prev_database,
                  balance_rows=balance_rows,
                  balance_prev_rows=balance_prev_rows)
    hook = FakeHook(formula_values=formula_values)
    rep = LK03Report()
    rep.fetch(ctx)
    rep.write(hook, ctx)
    return hook, ctx


def prev_calls(ctx) -> list[dict]:
    """只取"上期"那批查询（report_code 以 -上期 结尾）。"""
    return [c for c in ctx.fetch_calls
            if c["report_code"].endswith("-上期")]


def current_calls(ctx) -> list[dict]:
    """只取"本期"那批查询。"""
    return [c for c in ctx.fetch_calls
            if not c["report_code"].endswith("-上期")]


def call_with_item(calls: list[dict], item: str) -> dict:
    """从查询列表里挑出"参数中含某个项目编码"的那一条。

    【为什么需要它？】
        每行的项目集合不同（'01'、'02'…）→ 报表按集合分组，
        一行一组、各发一次查询。所以"本期有几次查询"会随
        已实现行数增长，不能再按下标取第 0 条。
        按项目编码找，既不受行数影响，也不受分组顺序影响。
    """
    for c in calls:
        if item in (c["params"] or []):
            return c
    raise AssertionError(f"未找到包含项目编码 {item!r} 的查询："
                         f"{[c['params'] for c in calls]}")


class TestPeriodCell(unittest.TestCase):
    """验证会计期间写入 B4（需求第 1 条）。"""

    def test_period_written_as_text(self):
        """B4 应等于"2026年5月"，且必须用 set_text 写入。"""
        hook, _ = run_report(month=5, year=2026)
        self.assertEqual(hook.get(PERIOD_CELL), "2026年5月")

        kinds = [e[0] for e in hook.events if e[1] == PERIOD_CELL]
        self.assertIn("set_text", kinds,
                      f"B4 未用 set_text 写入，实际事件：{hook.events}")
        self.assertNotIn("set", kinds,
                         "B4 不应使用普通 set 写入（可能被识别成日期）")

    def test_period_follows_selected_month(self):
        """换个月份生成时，B4 必须跟着变（不能写死）。"""
        for year, month, expect in (
            (2026, 5, "2026年5月"),
            (2026, 8, "2026年8月"),
            (2025, 11, "2025年11月"),
            (2026, 1, "2026年1月"),
            (2027, 12, "2027年12月"),
        ):
            with self.subTest(year=year, month=month):
                hook, _ = run_report(month=month, year=year)
                self.assertEqual(hook.get(PERIOD_CELL), expect)

    def test_no_duplicate_period_written(self):
        """同一页不应出现两处年月（fill_period=False 的意义）。"""
        hook, _ = run_report()
        writes = [e for e in hook.events if e[1] == PERIOD_CELL]
        self.assertEqual(len(writes), 1,
                         f"B4 应只写一次，实际 {len(writes)} 次：{writes}")

    def test_c4_never_written(self):
        """C4 绝不能写 —— 模板里 C4 是「复核：」标签。

        LK01、LK02 的会计期间都写在 **C4**，本表却是 **B4**。
        模板 LK03 第 4 行：A4='会计期间：'  B4=(空)  C4='复核：'

        若照抄 LK01 的写法去写 C4：
            · B4 仍是空白 → 报表上看不到会计期间；
            · C4 的"复核："被覆盖成"2026年5月" → 表头被破坏。
        而 C 列同时也是金额列，写入年月文字还有被识别成日期
        序列号的风险。整张表数字都正常，只有表头错了。
        """
        hook, _ = run_report()
        self.assertIsNone(hook.get("C4"),
                          f"C4 不应被写入（模板中它是「复核：」标签），"
                          f"实际 {hook.get('C4')}")

    def test_period_cell_is_b4(self):
        """会计期间常量必须是 B4（防止有人改回 C4）。"""
        self.assertEqual(PERIOD_CELL, "B4",
                         f"LK03 的会计期间应写 B4，实际 {PERIOD_CELL!r}")


class TestSalesCashReceived(unittest.TestCase):
    """需求第 2 条：行次 2「销售商品、提供劳务收到的现金」→ C10。

    口径：gl_cashtable 中 cCashItem='01' 的记录，
          取所选年月的单月数据，净额 = SUM(md) - SUM(mc)。
    """

    def test_c10_is_net_of_debit_minus_credit(self):
        """C10 = 借方 md － 贷方 mc。"""
        hook, _ = run_report(make_cash_rows({"01": (10000.0, 3000.0)}))
        self.assertTrue(approx(hook.get("C10"), 7000.0),
                        f"C10 应为 10000-3000=7000，实际 {hook.get('C10')}")

    def test_md_only(self):
        """只有借方时，C10 = md。"""
        hook, _ = run_report(make_cash_rows({"01": (8000.0, 0.0)}))
        self.assertTrue(approx(hook.get("C10"), 8000.0),
                        f"C10 应为 8000，实际 {hook.get('C10')}")

    def test_mc_only_gives_negative(self):
        """只有贷方时，C10 为负数（退款/冲回）。

        【为什么允许负数？】
            现金流量项目既可能有收款（md），也可能有退款（mc）。
            若某月只有退款，净额为负是业务的真实反映。
            强行取绝对值或写 0，都会掩盖业务事实。
        """
        hook, _ = run_report(make_cash_rows({"01": (0.0, 2000.0)}))
        self.assertTrue(approx(hook.get("C10"), -2000.0),
                        f"C10 应为 -2000，实际 {hook.get('C10')}")

    def test_net_zero_still_written(self):
        """md == mc 时净额为 0，必须写 0（不是留空）。

        【为什么这条重要？】
            "有记录且净额为 0"与"完全没有记录"是两件事：
              · 前者：本期确实收了又退了同样多，报表应显示 0；
              · 后者：没取到数，应留空让人去查。
            若用"净额是否为 0"判断是否写入，两者会被混为一谈。
            实现上用 row_has（是否取到记录）而不是净额判断。
        """
        hook, _ = run_report(make_cash_rows({"01": (5000.0, 5000.0)}))
        self.assertEqual(hook.get("C10"), 0.0,
                         f"有记录且净额为 0 时应写 0，实际 {hook.get('C10')}")

    def test_no_data_stays_blank(self):
        """完全没有记录时，C10 保持空白（不写 0）。"""
        hook, _ = run_report([])
        self.assertIsNone(hook.get("C10"),
                          f"C10 无数据时应空白，实际 {hook.get('C10')}")

    def test_fetch_failure_not_crash(self):
        """取数抛异常时不崩溃，且不写 C10。"""
        hook, _ = run_report([], fail=True)
        self.assertIsNone(hook.get("C10"))

    def test_other_items_ignored(self):
        """非 '01' 的项目不得混入 C10。

        【为什么重要？】
            若取数漏了 cCashItem 过滤，或报表层按项目分流写错，
            其它项目的金额会静默并入本行 —— 报表上只是个偏大的数，
            看不出多了什么。

        【干扰项为什么用 '04'/'05'？】
            '02'、'03' 已分别是行次 3、4 的合法项目。
            拿已实现的项目当"不该进来的项"，测试通过也不能说明
            分流逻辑正确 —— 它可能只是碰巧没写进 C10。
            用尚未实现的 '04'/'05' 才真正验证分流。
        """
        hook, _ = run_report(make_cash_rows({
            "01": (1000.0, 0.0),
            "04": (9999.0, 0.0),   # 未实现
            "05": (8888.0, 0.0),   # 未实现
        }))
        self.assertTrue(approx(hook.get("C10"), 1000.0),
                        f"C10 应只含 '01'=1000，实际 {hook.get('C10')}")

    def test_cross_month_records_summed(self):
        """同月多条记录应累加（GROUP BY 之外的兜底）。"""
        rows = [
            ("01", 202605, 1000.0, 0.0),
            ("01", 202605, 500.0, 200.0),   # 同月第二条
        ]
        hook, _ = run_report(rows)
        self.assertTrue(approx(hook.get("C10"), 1300.0),
                        f"C10 应为 1000+300=1300，实际 {hook.get('C10')}")

    def test_config_single_item_01(self):
        """配置层面固化：本行只取项目 '01'，C/D 两列都要填。"""
        spec = next(s for s in ROW_SPECS
                    if s.key == "销售商品提供劳务收到的现金")
        self.assertEqual(spec.cash_items, ("01",))
        self.assertEqual(spec.exclude_items, ())
        self.assertEqual(spec.excel_row, EXCEL_ROW[2], "应写第 10 行")
        self.assertTrue(spec.fill_current, "应填 C 列")
        self.assertTrue(spec.fill_prev, "应填 D 列（上期金额）")

    def test_query_uses_cashtable_and_single_month(self):
        """SQL 必须查 gl_cashtable、按 cCashItem 过滤、且是单月精确匹配。

        【为什么必须断言"单月"而不是区间？】
            本表 C 列是『本期金额』（单月），与 LK02 的"1~N 月累计"
            口径不同。若误用 BETWEEN 区间，C10 会变成累计数 ——
            报表上只是个数字，看不出它是单月还是累计，事后极难发现。
        """
        _, ctx = run_report(make_cash_rows({"01": (1.0, 0.0)}), month=5)

        call = call_with_item(current_calls(ctx), "01")
        sql = call["sql"]
        params = call["params"]

        self.assertIn("gl_cashtable", sql,
                      f"数据源应为 gl_cashtable，实际 SQL：{sql}")
        self.assertIn("cCashItem IN (", sql,
                      f"应按项目过滤，实际 SQL：{sql}")
        self.assertIn("iYPeriod = ?", sql,
                      f"应按单月精确匹配（不是 BETWEEN 区间），实际 SQL：{sql}")
        self.assertNotIn("BETWEEN", sql,
                         f"本表是单月口径，不应出现 BETWEEN 区间，实际 SQL：{sql}")

        # 参数顺序：① 年月整数  ②… 项目编码
        self.assertEqual(params[0], 202605,
                         f"第1个参数应为所选年月 202605，实际 {params[0]}")
        self.assertIn("01", params, "参数应包含项目编码 01")

    def test_query_month_follows_selection(self):
        """选择的年月变化时，查询参数必须跟着变（不能写死）。"""
        for month, expect in ((1, 202601), (5, 202605), (12, 202612)):
            with self.subTest(month=month):
                _, ctx = run_report(make_cash_rows({"01": (1.0, 0.0)}),
                                    month=month)
                for c in current_calls(ctx):
                    self.assertEqual(c["params"][0], expect,
                                     f"本期查询 iYPeriod 应为 {expect}，"
                                     f"实际 {c['params'][0]}")

    def test_formula_rows_still_untouched_with_data(self):
        """有数据时，公式行仍不得被写入。"""
        hook, _ = run_report(
            make_cash_rows({"01": (1000.0, 200.0)}),
            prev_rows=make_cash_rows({"01": (900.0, 100.0)}, month=4))
        for row_no in FORMULA_ROWS:
            excel_row = EXCEL_ROW[row_no]
            for col in (COL_CURRENT, COL_PREV):
                cell = f"{col}{excel_row}"
                self.assertIsNone(hook.get(cell),
                                  f"{cell}（行次{row_no}）是公式行，不得写入")


class TestTaxRefund(unittest.TestCase):
    """需求第 3 条：行次 3「收到的税费返还」→ C11 / D11。

    口径与需求第 2 条完全相同，**只有项目编码不同**：'02'。
    """

    def test_c11_is_net_of_debit_minus_credit(self):
        """C11 = 借方 md － 贷方 mc。"""
        hook, _ = run_report(make_cash_rows({"02": (3000.0, 1000.0)}))
        self.assertTrue(approx(hook.get("C11"), 2000.0),
                        f"C11 应为 3000-1000=2000，实际 {hook.get('C11')}")

    def test_c11_mc_only_negative(self):
        """只有贷方时 C11 为负数（退税被退回/冲回）。"""
        hook, _ = run_report(make_cash_rows({"02": (0.0, 800.0)}))
        self.assertTrue(approx(hook.get("C11"), -800.0),
                        f"C11 应为 -800，实际 {hook.get('C11')}")

    def test_c11_no_data_blank(self):
        """无记录时 C11 保持空白（不写 0）。"""
        hook, _ = run_report([])
        self.assertIsNone(hook.get("C11"),
                          f"C11 无数据时应空白，实际 {hook.get('C11')}")

    def test_c11_net_zero_written(self):
        """有记录且净额为 0 时写 0（不是留空）。"""
        hook, _ = run_report(make_cash_rows({"02": (500.0, 500.0)}))
        self.assertEqual(hook.get("C11"), 0.0,
                         f"有记录且净额为 0 时应写 0，实际 {hook.get('C11')}")

    def test_d11_is_prev_period(self):
        """D11 取上一期（选 8 月 → 7 月）的 '02' 净额。"""
        hook, ctx = run_report(
            make_cash_rows({"02": (3000.0, 0.0)}, month=8),
            month=8,
            prev_rows=make_cash_rows({"02": (1200.0, 200.0)}, month=7))

        self.assertTrue(approx(hook.get("C11"), 3000.0),
                        f"C11 应为 3000，实际 {hook.get('C11')}")
        self.assertTrue(approx(hook.get("D11"), 1000.0),
                        f"D11 应为 1200-200=1000，实际 {hook.get('D11')}")

        pc = call_with_item(prev_calls(ctx), "02")
        self.assertEqual(pc["params"][0], 202607,
                         f"上期 iYPeriod 应为 202607，实际 {pc['params'][0]}")

    def test_d11_january_cross_year(self):
        """1 月时 D11 取上年 12 月，且查上一年度库（与 D10 同规则）。"""
        hook, ctx = run_report(
            make_cash_rows({"02": (100.0, 0.0)}, month=1),
            month=1,
            prev_rows=make_cash_rows({"02": (600.0, 100.0)},
                                     month=12, year=2025),
            prev_database="UFDATA_001_2025")

        self.assertTrue(approx(hook.get("D11"), 500.0),
                        f"D11 应为 600-100=500，实际 {hook.get('D11')}")

        pc = call_with_item(prev_calls(ctx), "02")
        self.assertEqual(pc["params"][0], 202512,
                         f"1 月上期 iYPeriod 应为 202512，"
                         f"实际 {pc['params'][0]}")
        self.assertEqual(pc["database"], "UFDATA_001_2025",
                         f"1 月上期应查上一年度库，实际 {pc['database']}")

    def test_d11_no_data_blank(self):
        """上期无 '02' 记录时 D11 留空。"""
        hook, _ = run_report(
            make_cash_rows({"02": (1000.0, 0.0)}, month=5),
            month=5, prev_rows=[])

        self.assertTrue(approx(hook.get("C11"), 1000.0))
        self.assertIsNone(hook.get("D11"),
                          f"上期无数据时 D11 应留空，实际 {hook.get('D11')}")

    def test_two_rows_not_crossed(self):
        """★ 行次2（'01'）与行次3（'02'）绝不串用。

        【为什么这条最关键？】
            两行口径完全相同，只有项目编码差一位（'01'/'02'）、
            Excel 行号差一行（10/11）。若分流映射写错，
            会出现"两行显示同一个数"或"某行空着"，
            而报表上每格都"有数字"，肉眼看不出任何异常。

            本条给数量级差异明显的金额，任一处串位立刻暴露。
        """
        hook, _ = run_report(make_cash_rows({
            "01": (100000.0, 0.0),
            "02": (7.0, 0.0),
        }))
        self.assertTrue(approx(hook.get("C10"), 100000.0),
                        f"C10 应含 '01'=100000，实际 {hook.get('C10')}")
        self.assertTrue(approx(hook.get("C11"), 7.0),
                        f"C11 应含 '02'=7，实际 {hook.get('C11')}")

    def test_two_rows_not_crossed_in_prev(self):
        """上期两行同样不串用。"""
        hook, _ = run_report(
            make_cash_rows({"01": (0.0, 0.0), "02": (0.0, 0.0)}, month=5),
            month=5,
            prev_rows=make_cash_rows({
                "01": (50000.0, 0.0),
                "02": (9.0, 0.0),
            }, month=4))

        self.assertTrue(approx(hook.get("D10"), 50000.0),
                        f"D10 应含 '01'=50000，实际 {hook.get('D10')}")
        self.assertTrue(approx(hook.get("D11"), 9.0),
                        f"D11 应含 '02'=9，实际 {hook.get('D11')}")

    def test_other_items_ignored(self):
        """非 '02' 的项目不得混入 C11。"""
        hook, _ = run_report(make_cash_rows({
            "01": (8888.0, 0.0),   # 属行次2，不应入 C11
            "02": (1000.0, 0.0),
            "04": (7777.0, 0.0),
        }))
        self.assertTrue(approx(hook.get("C11"), 1000.0),
                        f"C11 应只含 '02'=1000，实际 {hook.get('C11')}")

    def test_config_item_02(self):
        """配置层面固化：本行只取项目 '02'，C/D 两列都要填。"""
        spec = next(s for s in ROW_SPECS if s.key == "收到的税费返还")
        self.assertEqual(spec.cash_items, ("02",))
        self.assertEqual(spec.exclude_items, ())
        self.assertEqual(spec.excel_row, EXCEL_ROW[3], "应写第 11 行")
        self.assertTrue(spec.fill_current, "应填 C 列")
        self.assertTrue(spec.fill_prev, "应填 D 列")

    def test_query_uses_item_02(self):
        """SQL 参数必须带 '02'，且同样是单月精确匹配。"""
        _, ctx = run_report(make_cash_rows({"02": (1.0, 0.0)}), month=5)

        call = call_with_item(current_calls(ctx), "02")
        self.assertIn("gl_cashtable", call["sql"])
        self.assertIn("iYPeriod = ?", call["sql"])
        self.assertNotIn("BETWEEN", call["sql"])
        self.assertEqual(call["params"][0], 202605)

    def test_formula_rows_still_untouched(self):
        """有 '02' 数据时，公式行仍不得被写入。"""
        hook, _ = run_report(make_cash_rows({"02": (1000.0, 200.0)}))
        for row_no in FORMULA_ROWS:
            excel_row = EXCEL_ROW[row_no]
            for col in (COL_CURRENT, COL_PREV):
                cell = f"{col}{excel_row}"
                self.assertIsNone(hook.get(cell),
                                  f"{cell}（行次{row_no}）是公式行，不得写入")

    def test_cell_map_has_c11_d11(self):
        """C11/D11 必须登记进 cell_map（供勾稽核对读回）。"""
        values = set(LK03Report.cell_map.values())
        for cell in ("C11", "D11"):
            self.assertIn(cell, values,
                          f"{cell} 未登记 cell_map，实际 {LK03Report.cell_map}")


class TestOtherOperatingCashIn(unittest.TestCase):
    """需求第 4 条：行次 4「收到其他与经营活动有关的现金」→ C12 / D12。

    口径与需求第 2、3 条完全相同，项目编码为 '03'。

    【项目名与需求原文的差异（已在代码注释中记录）】
        需求该条写的是「收到的税费返还」，与上一条重复；
        但给定格位是 C12/D12，模板该行 A 列实为
        「收到其他与经营活动有关的现金」。这里按模板实现。
    """

    def test_c12_is_net_of_debit_minus_credit(self):
        """C12 = 借方 md － 贷方 mc。"""
        hook, _ = run_report(make_cash_rows({"03": (6000.0, 1500.0)}))
        self.assertTrue(approx(hook.get("C12"), 4500.0),
                        f"C12 应为 6000-1500=4500，实际 {hook.get('C12')}")

    def test_c12_mc_only_negative(self):
        """只有贷方时 C12 为负数（退款/冲回）。"""
        hook, _ = run_report(make_cash_rows({"03": (0.0, 700.0)}))
        self.assertTrue(approx(hook.get("C12"), -700.0),
                        f"C12 应为 -700，实际 {hook.get('C12')}")

    def test_c12_no_data_blank(self):
        """无记录时 C12 保持空白（不写 0）。"""
        hook, _ = run_report([])
        self.assertIsNone(hook.get("C12"),
                          f"C12 无数据时应空白，实际 {hook.get('C12')}")

    def test_c12_net_zero_written(self):
        """有记录且净额为 0 时写 0（不是留空）。"""
        hook, _ = run_report(make_cash_rows({"03": (900.0, 900.0)}))
        self.assertEqual(hook.get("C12"), 0.0,
                         f"有记录且净额为 0 时应写 0，实际 {hook.get('C12')}")

    def test_d12_is_prev_period(self):
        """D12 取上一期（选 8 月 → 7 月）的 '03' 净额。"""
        hook, ctx = run_report(
            make_cash_rows({"03": (6000.0, 0.0)}, month=8),
            month=8,
            prev_rows=make_cash_rows({"03": (2500.0, 500.0)}, month=7))

        self.assertTrue(approx(hook.get("C12"), 6000.0),
                        f"C12 应为 6000，实际 {hook.get('C12')}")
        self.assertTrue(approx(hook.get("D12"), 2000.0),
                        f"D12 应为 2500-500=2000，实际 {hook.get('D12')}")

        pc = call_with_item(prev_calls(ctx), "03")
        self.assertEqual(pc["params"][0], 202607,
                         f"上期 iYPeriod 应为 202607，实际 {pc['params'][0]}")

    def test_d12_january_cross_year(self):
        """1 月时 D12 取上年 12 月，且查上一年度库。"""
        hook, ctx = run_report(
            make_cash_rows({"03": (100.0, 0.0)}, month=1),
            month=1,
            prev_rows=make_cash_rows({"03": (900.0, 300.0)},
                                     month=12, year=2025),
            prev_database="UFDATA_001_2025")

        self.assertTrue(approx(hook.get("D12"), 600.0),
                        f"D12 应为 900-300=600，实际 {hook.get('D12')}")

        pc = call_with_item(prev_calls(ctx), "03")
        self.assertEqual(pc["params"][0], 202512,
                         f"1 月上期 iYPeriod 应为 202512，"
                         f"实际 {pc['params'][0]}")
        self.assertEqual(pc["database"], "UFDATA_001_2025",
                         f"1 月上期应查上一年度库，实际 {pc['database']}")

    def test_d12_no_data_blank(self):
        """上期无 '03' 记录时 D12 留空。"""
        hook, _ = run_report(
            make_cash_rows({"03": (1000.0, 0.0)}, month=5),
            month=5, prev_rows=[])

        self.assertTrue(approx(hook.get("C12"), 1000.0))
        self.assertIsNone(hook.get("D12"),
                          f"上期无数据时 D12 应留空，实际 {hook.get('D12')}")

    def test_three_rows_not_crossed(self):
        """★ 行次 2/3/4（'01'/'02'/'03'）三行互不串用。

        【为什么这条比两行互斥更重要？】
            现在有三行口径完全相同，只有项目编码与 Excel 行号不同：
                '01' → 第 10 行
                '02' → 第 11 行
                '03' → 第 12 行
            一旦分流映射写错（例如字典 key 重复、或行号抄错），
            可能出现"某两行显示同一个数"或"某行空着"，
            而报表上每格都"有数字"，完全看不出异常。

            本条给三个数量级明显不同的金额，任一处串位立刻暴露。
        """
        hook, _ = run_report(make_cash_rows({
            "01": (100000.0, 0.0),
            "02": (1000.0, 0.0),
            "03": (7.0, 0.0),
        }))
        self.assertTrue(approx(hook.get("C10"), 100000.0),
                        f"C10 应含 '01'=100000，实际 {hook.get('C10')}")
        self.assertTrue(approx(hook.get("C11"), 1000.0),
                        f"C11 应含 '02'=1000，实际 {hook.get('C11')}")
        self.assertTrue(approx(hook.get("C12"), 7.0),
                        f"C12 应含 '03'=7，实际 {hook.get('C12')}")

    def test_three_rows_not_crossed_in_prev(self):
        """上期三行同样不串用。"""
        hook, _ = run_report(
            make_cash_rows({"01": (0.0, 0.0), "02": (0.0, 0.0),
                            "03": (0.0, 0.0)}, month=5),
            month=5,
            prev_rows=make_cash_rows({
                "01": (50000.0, 0.0),
                "02": (500.0, 0.0),
                "03": (9.0, 0.0),
            }, month=4))

        self.assertTrue(approx(hook.get("D10"), 50000.0),
                        f"D10 应含 '01'=50000，实际 {hook.get('D10')}")
        self.assertTrue(approx(hook.get("D11"), 500.0),
                        f"D11 应含 '02'=500，实际 {hook.get('D11')}")
        self.assertTrue(approx(hook.get("D12"), 9.0),
                        f"D12 应含 '03'=9，实际 {hook.get('D12')}")

    def test_other_items_ignored(self):
        """非 '03' 的项目不得混入 C12。"""
        hook, _ = run_report(make_cash_rows({
            "01": (8888.0, 0.0),   # 属行次2
            "02": (7777.0, 0.0),   # 属行次3
            "03": (1000.0, 0.0),
            "05": (6666.0, 0.0),   # 未实现
        }))
        self.assertTrue(approx(hook.get("C12"), 1000.0),
                        f"C12 应只含 '03'=1000，实际 {hook.get('C12')}")

    def test_config_item_03(self):
        """配置层面固化：本行只取项目 '03'，C/D 两列都要填。"""
        spec = next(s for s in ROW_SPECS
                    if s.key == "收到其他与经营活动有关的现金")
        self.assertEqual(spec.cash_items, ("03",))
        self.assertEqual(spec.exclude_items, ())
        self.assertEqual(spec.excel_row, EXCEL_ROW[4], "应写第 12 行")
        self.assertTrue(spec.fill_current, "应填 C 列")
        self.assertTrue(spec.fill_prev, "应填 D 列")

    def test_label_matches_template_not_request(self):
        """label 必须用模板原文，不能照抄需求里的重复名称。

        【为什么要锁这条？】
            需求第 4 条原文写的是「收到的税费返还」，与第 3 条重复。
            若 label 也照抄，cell_map 里会出现两个同名指标：
                "收到的税费返还_本期金额" → C11
                "收到的税费返还_本期金额" → C12   ← 后者覆盖前者
            结果 C11 从核对清单里消失，而"核对通过"看起来正常 ——
            这正是规范里"指标名必须唯一、可区分"所要防的问题。
        """
        keys = [s.key for s in ROW_SPECS]
        self.assertEqual(len(keys), len(set(keys)),
                         f"ROW_SPECS 的 key 有重复：{keys}")

        spec = next(s for s in ROW_SPECS
                    if s.key == "收到其他与经营活动有关的现金")
        self.assertIn("其他与经营活动", spec.label,
                      f"label 应采用模板原文，实际 {spec.label!r}")
        # 与行次 3 的标签必须不同
        tax = next(s for s in ROW_SPECS if s.key == "收到的税费返还")
        self.assertNotEqual(spec.label, tax.label,
                            "两行标签不得相同，否则 cell_map 指标名会冲突")

    def test_query_uses_item_03(self):
        """SQL 参数必须带 '03'，且同样是单月精确匹配。"""
        _, ctx = run_report(make_cash_rows({"03": (1.0, 0.0)}), month=5)

        call = call_with_item(current_calls(ctx), "03")
        self.assertIn("gl_cashtable", call["sql"])
        self.assertIn("iYPeriod = ?", call["sql"])
        self.assertNotIn("BETWEEN", call["sql"])
        self.assertEqual(call["params"][0], 202605)

    def test_formula_rows_still_untouched(self):
        """有 '03' 数据时，公式行仍不得被写入。"""
        hook, _ = run_report(make_cash_rows({"03": (1000.0, 200.0)}))
        for row_no in FORMULA_ROWS:
            excel_row = EXCEL_ROW[row_no]
            for col in (COL_CURRENT, COL_PREV):
                cell = f"{col}{excel_row}"
                self.assertIsNone(hook.get(cell),
                                  f"{cell}（行次{row_no}）是公式行，不得写入")

    def test_cell_map_has_c12_d12(self):
        """C12/D12 必须登记进 cell_map（供勾稽核对读回）。"""
        values = set(LK03Report.cell_map.values())
        for cell in ("C12", "D12"):
            self.assertIn(cell, values,
                          f"{cell} 未登记 cell_map，实际 {LK03Report.cell_map}")


class TestPrevPeriodAmount(unittest.TestCase):
    """需求第 2 条补充：D10 = 上一期（iYPeriod 减一）的同一项目净额。

    规则（现场 20261106 确认）：
        · 所选月份 > 1 → 上期 = 同年上月，仍在**本年度库**；
        · 所选月份 = 1 → 上期 = 上年 12 月，在**上一年度库**。
    """

    def test_non_january_prev_is_previous_month(self):
        """选 8 月 → 上期是 7 月（同一库）。"""
        hook, ctx = run_report(
            make_cash_rows({"01": (10000.0, 0.0)}, month=8),
            month=8,
            prev_rows=make_cash_rows({"01": (8000.0, 1000.0)}, month=7))

        self.assertTrue(approx(hook.get("C10"), 10000.0),
                        f"C10 应为 10000，实际 {hook.get('C10')}")
        self.assertTrue(approx(hook.get("D10"), 7000.0),
                        f"D10 应为 8000-1000=7000，实际 {hook.get('D10')}")

        # 上期查询必须是 202607
        pc = prev_calls(ctx)
        self.assertTrue(pc, "应发起上期查询")
        for c in pc:
            self.assertEqual(c["params"][0], 202607,
                             f"上期 iYPeriod 应为 202607，"
                             f"实际 {c['params'][0]}")

    def test_non_january_prev_uses_target_database(self):
        """非 1 月时，上期仍查**本年度库**（不切到上年库）。"""
        _, ctx = run_report(make_cash_rows({"01": (1.0, 0.0)}, month=6),
                            month=6,
                            prev_rows=make_cash_rows({"01": (2.0, 0.0)},
                                                     month=5))
        pc = prev_calls(ctx)
        self.assertEqual(pc[0]["database"], TARGET_DB,
                         f"上期应查本年度库 {TARGET_DB}，"
                         f"实际 {pc[0]['database']}")

    def test_january_prev_is_last_december(self):
        """选 2026年1月 → 上期是 2025年12月（iYPeriod=202512）。"""
        hook, ctx = run_report(
            make_cash_rows({"01": (500.0, 0.0)}, month=1, year=2026),
            month=1, year=2026,
            prev_rows=make_cash_rows({"01": (700.0, 200.0)},
                                     month=12, year=2025),
            prev_database="UFDATA_001_2025")

        self.assertTrue(approx(hook.get("C10"), 500.0),
                        f"C10 应为 500，实际 {hook.get('C10')}")
        self.assertTrue(approx(hook.get("D10"), 500.0),
                        f"D10 应为 700-200=500，实际 {hook.get('D10')}")

        pc = prev_calls(ctx)
        self.assertTrue(pc, "应发起上期查询")
        for c in pc:
            self.assertEqual(c["params"][0], 202512,
                             f"1 月的上期 iYPeriod 应为 202512，"
                             f"实际 {c['params'][0]}")

    def test_january_prev_uses_prev_database(self):
        """★ 关键：1 月的上期必须查**上一年度库**，不能查本年度库。

        【为什么这条最重要？】
            分年度多库的账套里，202512 的记录物理上不在 2026 库。
            若仍查本年度库，SQL 能跑通、只是查不到行，D 列静默留空 ——
            报表上看起来像"上月确实没有流量"，与"确实没有数据"
            完全无法区分，且不报任何错。
        """
        _, ctx = run_report(
            make_cash_rows({"01": (1.0, 0.0)}, month=1),
            month=1,
            prev_rows=make_cash_rows({"01": (2.0, 0.0)},
                                     month=12, year=2025),
            prev_database="UFDATA_001_2025")

        pc = prev_calls(ctx)
        self.assertEqual(pc[0]["database"], "UFDATA_001_2025",
                         f"1 月的上期应查上一年度库 UFDATA_001_2025，"
                         f"实际 {pc[0]['database']}")

    def test_january_single_database_also_works(self):
        """跨年度单库：1 月的上期与本期间同库，仍应正常取数。

        【为什么单测？】
            prev_database 与 target_database 相同是常见形态
            （U8 推荐的跨年度单库）。此时"换库"动作实际上没换，
            但解析路径不同，要确认不会因此漏查。
        """
        hook, ctx = run_report(
            make_cash_rows({"01": (1.0, 0.0)}, month=1),
            month=1,
            prev_rows=make_cash_rows({"01": (300.0, 0.0)},
                                     month=12, year=2025),
            prev_database="UFDATA_001_2026")

        self.assertTrue(approx(hook.get("D10"), 300.0),
                        f"D10 应为 300，实际 {hook.get('D10')}")
        self.assertEqual(prev_calls(ctx)[0]["database"], "UFDATA_001_2026")

    def test_first_year_january_skips_prev_query(self):
        """账套首年的 1 月：无上年库，不查上期，D10 留空。

        【为什么要"不查"而不是"查了为空"？】
            查也查不到，只会白刷一条"未取到记录"的日志，
            把真正的问题淹没。而且首年没有上年数据是**确定的事实**，
            不是"可能有也可能没有"。
        """
        hook, ctx = run_report(
            make_cash_rows({"01": (500.0, 0.0)}, month=1),
            month=1, prev_database="")

        self.assertTrue(approx(hook.get("C10"), 500.0))
        self.assertIsNone(hook.get("D10"),
                          f"首年无上年库时 D10 应留空，实际 {hook.get('D10')}")
        self.assertEqual(len(prev_calls(ctx)), 0,
                         "首年 1 月不应发起上期查询")

    def test_prev_no_data_leaves_blank(self):
        """上期查不到记录（可能有、可能无）时，D10 留空不写 0。"""
        hook, _ = run_report(
            make_cash_rows({"01": (5000.0, 0.0)}, month=5),
            month=5, prev_rows=[])

        self.assertTrue(approx(hook.get("C10"), 5000.0))
        self.assertIsNone(hook.get("D10"),
                          f"上期无数据时 D10 应留空，实际 {hook.get('D10')}")

    def test_prev_only_other_items_leaves_blank(self):
        """上期只有别的项目（无 '01'）时，D10 留空。

        【干扰项为什么用 '04' 而不是 '02'/'03'？】
            '02'、'03' 已是行次 3、4 的合法项目。
            若仍拿它们当"别的项目"，测试通过也不能说明
            "报表层把 '01' 与其它项目区分开了" ——
            它可能只是碰巧没写进 D10。
            用尚未实现的 '04' 才真正验证分流逻辑。
        """
        hook, _ = run_report(
            make_cash_rows({"01": (5000.0, 0.0)}, month=5),
            month=5,
            prev_rows=make_cash_rows({"04": (999.0, 0.0)}, month=4))

        self.assertTrue(approx(hook.get("C10"), 5000.0))
        self.assertIsNone(hook.get("D10"),
                          f"上期无 '01' 时 D10 应留空，实际 {hook.get('D10')}")

    def test_prev_net_negative(self):
        """上期只有贷方时，D10 为负数（退款/冲回）。"""
        hook, _ = run_report(
            make_cash_rows({"01": (0.0, 0.0)}, month=5),
            month=5,
            prev_rows=make_cash_rows({"01": (0.0, 1500.0)}, month=4))

        self.assertTrue(approx(hook.get("D10"), -1500.0),
                        f"D10 应为 -1500，实际 {hook.get('D10')}")

    def test_prev_net_zero_written(self):
        """上期有记录且净额为 0 时，D10 写 0（不是留空）。"""
        hook, _ = run_report(
            make_cash_rows({"01": (0.0, 0.0)}, month=5),
            month=5,
            prev_rows=make_cash_rows({"01": (2000.0, 2000.0)}, month=4))

        self.assertEqual(hook.get("D10"), 0.0,
                         f"上期有记录且净额为 0 时应写 0，"
                         f"实际 {hook.get('D10')}")

    def test_current_and_prev_not_crossed(self):
        """本期与上期的数绝不互相混入（给数量级差异明显的金额）。"""
        hook, _ = run_report(
            make_cash_rows({"01": (100000.0, 0.0)}, month=5),
            month=5,
            prev_rows=make_cash_rows({"01": (7.0, 0.0)}, month=4))

        self.assertTrue(approx(hook.get("C10"), 100000.0),
                        f"C10 应为本期 100000，实际 {hook.get('C10')}")
        self.assertTrue(approx(hook.get("D10"), 7.0),
                        f"D10 应为上期 7，实际 {hook.get('D10')}")

    def test_rows_with_wrong_period_skipped(self):
        """纵深防御：iYPeriod 与期望月份不符的记录必须跳过。

        【为什么需要？】
            取数 SQL 已按单月精确过滤，正常不会出现别的月份。
            但本期与上期是两次查询，若将来有人把 SQL 改成区间，
            两期的数就会互相混入 —— 而报表上只是个数字，
            看不出它是哪个月的。
        """
        # 本期查询返回了 7 月的数据（本应是 5 月）
        hook, _ = run_report(
            [("01", 202607, 9999.0, 0.0)],
            month=5, prev_database="")

        self.assertIsNone(hook.get("C10"),
                          f"月份不符的记录应被跳过，C10 应留空，"
                          f"实际 {hook.get('C10')}")

    def test_prev_cell_in_cell_map(self):
        """D10 现在会写数值，必须登记进 cell_map。

        【为什么锁这条？】
            需求第 2 条最初只给了 C10，D10 当时不写，故未登记。
            现在 D 列口径已确认并开始填写，若仍不登记，
            该格不进勾稽核对清单，日后"核对通过"可能是假象。
        """
        self.assertIn("D10", LK03Report.cell_map.values(),
                      f"D10 已开始填写，应登记进 cell_map，"
                      f"实际 {LK03Report.cell_map}")


class TestFormulaRows(unittest.TestCase):
    """公式行（小计/净额/余额）永不得写入。"""

    def test_formula_rows_never_written(self):
        """任何情况下，公式行的 C/D 两列都不得被写入。"""
        hook, _ = run_report()
        for row_no in FORMULA_ROWS:
            excel_row = EXCEL_ROW[row_no]
            for col in (COL_CURRENT, COL_PREV):
                cell = f"{col}{excel_row}"
                self.assertIsNone(
                    hook.get(cell),
                    f"{cell}（行次{row_no}）是公式行，程序不得写入")

    def test_section_rows_never_written(self):
        """分段标题行（行次 1/12/25）只有文字，不得写入金额。"""
        hook, _ = run_report()
        for row_no in SECTION_ROWS:
            excel_row = EXCEL_ROW[row_no]
            for col in (COL_CURRENT, COL_PREV):
                cell = f"{col}{excel_row}"
                self.assertIsNone(
                    hook.get(cell),
                    f"{cell}（行次{row_no}）是分段标题行，不得写入")

    def test_formula_rows_not_in_specs(self):
        """公式行不得出现在 ROW_SPECS 中。"""
        formula_excel_rows = {EXCEL_ROW[n] for n in FORMULA_ROWS}
        for spec in ROW_SPECS:
            self.assertNotIn(
                spec.excel_row, formula_excel_rows,
                f"「{spec.label}」的 Excel 行 {spec.excel_row} 属公式行，"
                f"程序不应写入（会覆盖模板公式）")

    def test_formula_rows_values(self):
        """公式行清单本身必须是 11 个，且含关键的净额行。"""
        self.assertEqual(
            set(FORMULA_ROWS),
            {5, 10, 11, 18, 23, 24, 29, 33, 34, 36, 38},
            "公式行清单与模板核对结果不一致")
        for row_no in (11, 24, 34, 36, 38):
            self.assertIn(row_no, FORMULA_ROWS)


class TestStructuralConstants(unittest.TestCase):
    """常量本身的正确性（坐标写错会静默填到别的项目上）。"""

    def test_excel_row_covers_all_periods(self):
        """EXCEL_ROW 必须覆盖行次 1~38，且不重复、不越界。"""
        self.assertEqual(sorted(EXCEL_ROW.keys()), list(range(1, 39)))

        excel_rows = sorted(EXCEL_ROW.values())
        self.assertEqual(excel_rows, list(range(9, 47)),
                         "Excel 行号应为 9~46 且连续")
        self.assertEqual(len(set(excel_rows)), len(excel_rows),
                         "Excel 行号有重复")

    def test_excel_row_anchors(self):
        """关键锚点：明细行、公式行、末行的 Excel 行号。"""
        self.assertEqual(EXCEL_ROW[1], 9, "行次1 经营活动标题应在第9行")
        self.assertEqual(EXCEL_ROW[2], 10, "行次2 销售商品收到现金应在第10行")
        self.assertEqual(EXCEL_ROW[5], 13, "行次5 经营活动流入小计应在第13行")
        self.assertEqual(EXCEL_ROW[11], 19, "行次11 经营活动净额应在第19行")
        self.assertEqual(EXCEL_ROW[12], 20, "行次12 投资活动标题应在第20行")
        self.assertEqual(EXCEL_ROW[25], 33, "行次25 筹资活动标题应在第33行")
        self.assertEqual(EXCEL_ROW[35], 43, "行次35 汇率变动影响应在第43行")
        self.assertEqual(EXCEL_ROW[37], 45, "行次37 期初余额应在第45行")
        self.assertEqual(EXCEL_ROW[38], 46, "行次38 期末余额应在第46行")

    def test_section_rows(self):
        """分段标题行应为 1/12/25（一、二、三）。"""
        self.assertEqual(set(SECTION_ROWS), {1, 12, 25})

    def test_columns(self):
        """本期金额 C 列、上期金额 D 列。"""
        self.assertEqual(COL_CURRENT, "C")
        self.assertEqual(COL_PREV, "D")

    def test_no_row_overlap_between_categories(self):
        """明细行、公式行、分段标题行三者不得重叠。"""
        details = set(range(1, 39)) - set(FORMULA_ROWS) - set(SECTION_ROWS)
        self.assertFalse(details & set(FORMULA_ROWS))
        self.assertFalse(details & set(SECTION_ROWS))
        self.assertFalse(set(FORMULA_ROWS) & set(SECTION_ROWS))
        self.assertEqual(len(details), 24,
                         f"明细行应为 24 行，实际 {len(details)}：{sorted(details)}")


class TestRowSpecs(unittest.TestCase):
    """已实现行的配置校验。"""

    def test_specs_excel_row_exists(self):
        """spec 的 excel_row 必须来自 EXCEL_ROW 的值（防手写行号）。"""
        allowed = set(EXCEL_ROW.values())
        for spec in ROW_SPECS:
            self.assertIn(spec.excel_row, allowed,
                          f"「{spec.key}」的 excel_row={spec.excel_row} "
                          f"不在 EXCEL_ROW 对照表中")

    def test_no_duplicate_excel_rows(self):
        """不同行项目不得占用同一 Excel 行。"""
        seen: dict[int, str] = {}
        for spec in ROW_SPECS:
            self.assertNotIn(
                spec.excel_row, seen,
                f"Excel 行 {spec.excel_row} 被「{seen.get(spec.excel_row)}」"
                f"和「{spec.key}」同时占用")
            seen[spec.excel_row] = spec.key

    def test_cash_items_are_strings(self):
        """项目编码必须是字符串（'01' 写成 1 会匹配不上）。"""
        for spec in ROW_SPECS:
            for c in (*spec.cash_items, *spec.exclude_items):
                self.assertIsInstance(
                    c, str,
                    f"「{spec.key}」的项目编码 {c!r} 不是字符串")

    def test_cash_items_not_empty(self):
        """已配置的行必须有项目编码（否则填不出数，且原因不明）。"""
        for spec in ROW_SPECS:
            self.assertTrue(spec.cash_items,
                            f"「{spec.key}」的 cash_items 为空")


class TestPendingRows(unittest.TestCase):
    """待定项清单的结构正确性。"""

    def _implemented_period_rows(self) -> set[int]:
        """把已实现行的 excel_row 反查成模板行次。

        【★ 必须同时包含 BALANCE_SPECS（需求31 起）★】
            行次 37「期初现金及现金等价物余额」走 GL_accsum 余额线，
            登记在 BALANCE_SPECS 而不是 ROW_SPECS。
            若这里只看 ROW_SPECS，行次 37 就既不在"已实现"
            也不在"待定"，被误判成"缺口" ——
            而下方的覆盖恒等式会失败，指向一个并不存在的问题。
        """
        excel_rows = {s.excel_row for s in ROW_SPECS}
        excel_rows |= {b.excel_row for b in BALANCE_SPECS}
        return {row_no for row_no, er in EXCEL_ROW.items() if er in excel_rows}

    def test_pending_plus_implemented_covers_all_detail_rows(self):
        """待定项 + 已实现行 = 全部 24 个明细行（没有空洞）。"""
        detail_rows = (set(EXCEL_ROW.keys())
                       - set(FORMULA_ROWS) - set(SECTION_ROWS))
        pending_rows = {row_no for _, _, row_no in PENDING_ROWS}
        implemented = self._implemented_period_rows()

        self.assertEqual(
            pending_rows | implemented, detail_rows,
            f"缺口：{sorted(detail_rows - pending_rows - implemented)}")
        self.assertFalse(pending_rows & implemented,
                         f"重叠：{sorted(pending_rows & implemented)}")
        self.assertEqual(len(detail_rows), 24)

    def test_pending_cells_match_excel_row(self):
        """待定项登记的格位，其行号必须与 EXCEL_ROW 一致。"""
        for name, cells, row_no in PENDING_ROWS:
            expect_row = EXCEL_ROW[row_no]
            for cell in cells.split("/"):
                self.assertEqual(
                    cell_row(cell), expect_row,
                    f"待定项「{name}」（行次{row_no}）的 {cell} "
                    f"行号应为 {expect_row}")
                self.assertIn(cell_col(cell), (COL_CURRENT, COL_PREV),
                              f"待定项「{name}」的 {cell} 列位非法")

    def test_pending_not_conflict_with_implemented(self):
        """待定项不得与已实现行占用同一格子。"""
        implemented = set()
        for spec in ROW_SPECS:
            implemented.add(f"{COL_CURRENT}{spec.excel_row}")
            implemented.add(f"{COL_PREV}{spec.excel_row}")

        for name, cells, row_no in PENDING_ROWS:
            for cell in cells.split("/"):
                self.assertNotIn(
                    cell, implemented,
                    f"待定项「{name}」的 {cell} 与已实现行冲突（行次 {row_no}）")

    def test_pending_unique(self):
        """待定项之间不得重复。"""
        seen: dict[str, str] = {}
        for name, cells, _ in PENDING_ROWS:
            for cell in cells.split("/"):
                self.assertNotIn(
                    cell, seen,
                    f"待定项 {cell} 被「{seen.get(cell)}」和「{name}」同时占用")
                seen[cell] = name

    def test_pending_not_in_formula_or_section(self):
        """待定项不得落在公式行或分段标题行上。"""
        forbidden = set(FORMULA_ROWS) | set(SECTION_ROWS)
        for name, _cells, row_no in PENDING_ROWS:
            self.assertNotIn(
                row_no, forbidden,
                f"待定项「{name}」（行次{row_no}）落在公式行/标题行上")

    def test_pending_cells_not_written(self):
        """待定行的格子必须保持空白（确认"暂不实现"真的没写数）。"""
        hook, _ = run_report([])
        for _name, cells, _row_no in PENDING_ROWS:
            for cell in cells.split("/"):
                self.assertIsNone(hook.get(cell),
                                  f"待定行 {cell} 不应被写入，"
                                  f"实际 {hook.get(cell)}")

    def test_implemented_rows_not_pending(self):
        """已实现的行次不得再出现在待定清单里。

        【为什么？】
            若两处都有，读代码的人无法判断它到底做没做；
            且日后补实现时可能重复取数、写到同一格上。
        """
        implemented = self._implemented_period_rows()
        pending_rows = {row_no for _, _, row_no in PENDING_ROWS}
        for row_no in (2, 3, 4):
            self.assertIn(row_no, implemented,
                          f"行次 {row_no} 应已实现")
            self.assertNotIn(row_no, pending_rows,
                             f"行次 {row_no} 已实现，应从 PENDING_ROWS 移除")


class TestCellMap(unittest.TestCase):
    """cell_map 与 ROW_SPECS 的一致性。"""

    def test_period_registered(self):
        """会计期间必须登记进 cell_map。"""
        self.assertEqual(LK03Report.cell_map.get("会计期间"), PERIOD_CELL)

    def test_cell_map_matches_specs(self):
        """ROW_SPECS 中"确实会写数值"的格子应登记进 cell_map。

        【为什么按 fill_current / fill_prev 判断？】
            D 列（上期金额）当前口径未定、程序不写。
            若也要求登记，verify() 会读回 None，
            核对报告里该指标永远是"无法核对"，
            真正的核对失败反而被淹没。
        """
        cm_values = set(LK03Report.cell_map.values())
        for spec in ROW_SPECS:
            if spec.fill_current:
                cell = f"{COL_CURRENT}{spec.excel_row}"
                self.assertIn(cell, cm_values,
                              f"「{spec.key}」的 {cell} 未登记 cell_map")
            if spec.fill_prev:
                cell = f"{COL_PREV}{spec.excel_row}"
                self.assertIn(cell, cm_values,
                              f"「{spec.key}」的 {cell} 未登记 cell_map")

    def test_unfilled_cells_not_registered(self):
        """不写数值的格子不得登记进 cell_map。

        【为什么？】
            规范明确：cell_map 只登记"确实写了数值"的格子。
            登记一个永远空白的格子，核对时会得到 None（无法核对），
            属于噪音，会掩盖真正的问题。
        """
        cm_values = set(LK03Report.cell_map.values())
        for spec in ROW_SPECS:
            if not spec.fill_current:
                self.assertNotIn(f"{COL_CURRENT}{spec.excel_row}", cm_values,
                                 f"{spec.key} 不填 C 列，却登记进了 cell_map")
            if not spec.fill_prev:
                self.assertNotIn(f"{COL_PREV}{spec.excel_row}", cm_values,
                                 f"{spec.key} 不填 D 列，却登记进了 cell_map")

    def test_cell_map_values_point_to_value_cells(self):
        """cell_map 只能指向 C/D 列（防止误登记表头文字格）。"""
        for name, cell in LK03Report.cell_map.items():
            if name == "会计期间":
                self.assertEqual(cell, PERIOD_CELL)
                continue
            self.assertIn(cell_col(cell), (COL_CURRENT, COL_PREV),
                          f"cell_map「{name}」指向 {cell}，不在 C/D 金额列")


class TestReportMeta(unittest.TestCase):
    """报表类元信息。"""

    def test_code_and_name(self):
        self.assertEqual(LK03Report.code, "LK03")
        self.assertEqual(LK03Report.name, "现金流量表")

    def test_no_prev_sheet_suffix(self):
        """暂不声明上年对比页。

        【为什么锁这条？】
            「上期金额」D 列到底指"上年同期"还是"本年上月"，
            需求尚未给出。一旦确认是上年同期，才应设 prev_sheet_suffix
            并复用 ctx.prev_database（与 LK02 同机制）。
            在此之前若提前声明，框架会复制出一张空白对比页，
            增加维护负担且可能被误认为"已实现上期数"。
        """
        self.assertEqual(LK03Report.prev_sheet_suffix, "",
                         "上期口径未确认前，不应声明上年对比页")


class TestTemplateConsistency(unittest.TestCase):
    """直接读模板 .xls，校验坐标确实压在模板的实际内容上。"""

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

    def test_excel_rows_have_labels(self):
        """EXCEL_ROW 里每个 Excel 行，模板 A 列应有项目名。"""
        sh = self._open_lk03()
        for row_no, excel_row in EXCEL_ROW.items():
            val = str(sh.cell_value(excel_row - 1, 0)).strip()
            self.assertTrue(
                val,
                f"行次 {row_no} 应为 Excel 第 {excel_row} 行，"
                f"但该行 A 列是空的 —— 行号可能写错了")

    def test_template_row_numbers_match(self):
        """模板 B 列的行次必须与 EXCEL_ROW 逐行相符。

        【这条比"A 列非空"更严】
            A 列非空只能证明"没写到空行上"。
            若整体错一行（例如行次2 配成 Excel 13 行），
            A 列同样非空（那是"经营活动现金流入小计"），
            只有比对 B 列的行次数值才能发现。
        """
        sh = self._open_lk03()
        for row_no, excel_row in EXCEL_ROW.items():
            raw = sh.cell_value(excel_row - 1, 1)   # B 列
            try:
                actual = int(float(raw))
            except (TypeError, ValueError):
                self.fail(f"行次 {row_no}：Excel 第 {excel_row} 行的 B 列 "
                          f"不是数字，实际 {raw!r}")
            self.assertEqual(
                actual, row_no,
                f"Excel 第 {excel_row} 行的 B 列行次应为 {row_no}，"
                f"实际 {actual} —— EXCEL_ROW 对照表可能有偏移")

    def test_period_cell_position(self):
        """B4 是填写位：A4 为"会计期间"，B4 为空，C4 为"复核"。"""
        sh = self._open_lk03()
        a4 = str(sh.cell_value(3, 0)).strip()   # A4
        b4 = str(sh.cell_value(3, 1)).strip()   # B4
        c4 = str(sh.cell_value(3, 2)).strip()   # C4

        self.assertIn("会计期间", a4,
                      f"模板 A4 应为「会计期间」，实际 {a4!r}")
        self.assertEqual(b4, "",
                         f"模板 B4 应为空白填写位，实际 {b4!r} —— "
                         f"若模板已改成有内容，需重新确认会计期间写哪一格")
        self.assertIn("复核", c4,
                      f"模板 C4 应为「复核：」标签，实际 {c4!r} —— "
                      f"这正是不写 C4 的原因")

    def test_row10_is_sales_cash_received(self):
        """Excel 第 10 行（行次 2）A 列应是「销售商品、提供劳务收到的现金」。"""
        sh = self._open_lk03()
        val = str(sh.cell_value(EXCEL_ROW[2] - 1, 0)).strip()
        self.assertIn("销售商品", val,
                      f"行次 2 的 A 列应为「销售商品、提供劳务收到的现金」，"
                      f"实际 {val!r} —— C10 可能填错行")

    def test_formula_rows_are_summary_rows(self):
        """公式行在模板中应是"小计/净额/余额"类汇总项目。"""
        sh = self._open_lk03()
        for row_no in FORMULA_ROWS:
            excel_row = EXCEL_ROW[row_no]
            val = str(sh.cell_value(excel_row - 1, 0)).strip()
            self.assertTrue(
                any(k in val for k in ("小计", "净额", "余额", "净增加额")),
                f"行次 {row_no}（Excel 第 {excel_row} 行）应是汇总行，"
                f"实际 A 列文字为 {val!r}")

    def test_section_rows_are_headings(self):
        """分段标题行在模板中应是"一、""二、""三、"开头。"""
        sh = self._open_lk03()
        for row_no in SECTION_ROWS:
            excel_row = EXCEL_ROW[row_no]
            val = str(sh.cell_value(excel_row - 1, 0)).strip()
            ok = (val.startswith("一、") or val.startswith("二、")
                  or val.startswith("三、"))
            self.assertTrue(
                ok,
                f"行次 {row_no}（Excel 第 {excel_row} 行）应是分段标题，"
                f"实际 A 列文字为 {val!r}")


class TestIntentionallyUntouched(unittest.TestCase):
    """需求第 5 条：C13 / D13 不做处理，保持原样。

    这两格对应行次 5「经营活动现金流入小计」（Excel 第 13 行）。
    该行由模板公式自行合计本期三类经营流入（行次 2、3、4），程序不填。

    【为什么要单独登记，而不只是放在 FORMULA_ROWS 里？】
        FORMULA_ROWS 表达**技术约束**："写了会覆盖模板合计公式"。
        需求第 5 条表达**业务决定**："这几格保持原样"。

        两者当前重叠，但语义不同：若将来模板改版、该行不再是公式行，
        FORMULA_ROWS 会调整，而"需求要求不写"的决定不会变。
        只记在 FORMULA_ROWS 里，届时它会被当成"可以填数了"而误写。
    """

    def test_c13_d13_present(self):
        """C13/D13 必须登记（需求 16 后清单已扩充，不再断言"恰好两格"）。"""
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
        """这两格所在行应同时属于公式行（模板自带合计公式）。"""
        formula_cells = {
            f"{col}{EXCEL_ROW[n]}"
            for n in FORMULA_ROWS for col in (COL_CURRENT, COL_PREV)
        }
        for cell in INTENTIONALLY_UNTOUCHED:
            self.assertIn(cell, formula_cells,
                          f"{cell} 不在 FORMULA_ROWS 覆盖范围内")

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

        # 先确认上游明细行确实有数（否则本测试等于没测）
        self.assertTrue(approx(hook.get("C12"), 300.0),
                        f"前置条件不成立：C12 应有数据，实际 {hook.get('C12')}")

        for cell in INTENTIONALLY_UNTOUCHED:
            self.assertIsNone(hook.get(cell),
                              f"{cell} 已确认不做处理，不得写入，"
                              f"实际 {hook.get(cell)}")

    def test_not_in_cell_map(self):
        """这两格不得登记进 cell_map。

        cell_map 是"确实写了数值的格子"的清单。
        登记一个永远空白的格子，verify() 会读回 None，
        核对报告里该指标永远是"无法核对"，属于噪音。
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
                             f"{cell} 已确认不做处理，却被某条 ROW_SPECS 占用")

    def test_not_in_pending_rows(self):
        """这两格不得出现在待定清单里。

        PENDING_ROWS 是"口径待定、等财务给条件就能做"。
        本需求是"确定不做"，混进去会让人误以为还有待办，
        也会破坏"待定 + 已实现 = 24 个明细行"的覆盖恒等式。
        """
        pending_cells = set()
        for _name, cells, _row_no in PENDING_ROWS:
            pending_cells.update(cells.split("/"))
        for cell in INTENTIONALLY_UNTOUCHED:
            self.assertNotIn(cell, pending_cells,
                             f"{cell} 是确定不做，不应登记为待定项")


class TestFormulaRowNegativeFormat(unittest.TestCase):
    """公式行为负数时也要修正负号间距（20261110 现场反馈第二批）。

    【背景：为什么第一轮只改好了 D14，D18 还是老样子】
        D14 是行次 6「购买商品支付的现金」—— **明细行**，
        程序写入它时会顺带改格式，所以第一轮就修好了。

        D18 是行次 10「经营活动现金流出小计」—— **公式行**，
        模板自带 =SUM(D14:D17)。程序从不写入它的值
        （写了会覆盖公式），因此永远不会走到 _write_cell()，
        也就从未被改过格式 —— 负数继续按模板会计格式
        显示成 `-        1,036,332.84`。

        两行紧挨着，一个改了一个没改，看起来就像"改漏了"。

    【本组测什么】
        程序要**读取**公式行的值（由 Excel 算出），
        若为负则**只改格式、不改值**。
    """

    def test_negative_formula_row_formatted(self):
        """公式行为负时，格式被改为负号紧贴。"""
        hook, _ = run_report(
            make_cash_rows({"04": (1500.0, 1000.0)}),
            formula_values={"D18": -1036332.84})
        self.assertIn("D18", hook.number_formats,
                      "负数公式行 D18 未设置数字格式")
        self.assertIn("-#,##0.00", hook.number_formats["D18"])

    def test_positive_formula_row_not_formatted(self):
        """公式行为正时不得改格式。"""
        hook, _ = run_report(
            make_cash_rows({"04": (0.0, 1000.0)}),
            formula_values={"D18": 5000.0})
        self.assertNotIn("D18", hook.number_formats)

    def test_formula_value_never_written(self):
        """★ 公式行的值绝不能被程序写入（铁律）。

        【为什么必须锁死？】
            一旦写入数值，模板公式即被覆盖，
            下月再生成时公式已经没了，只剩本月算出的死数，
            且不报任何错。这比"格式难看"严重得多。
        """
        hook, _ = run_report(
            make_cash_rows({"04": (1500.0, 1000.0)}),
            formula_values={"D18": -1036332.84})
        self.assertNotIn("D18", hook.written,
                         "公式行 D18 不得被写入值")
        self.assertNotIn("C18", hook.written)

    def test_all_formula_rows_checked(self):
        """所有公式行都参与负数格式检查。"""
        values = {}
        for row_no in FORMULA_ROWS:
            values["C%d" % EXCEL_ROW[row_no]] = -100.0
        hook, _ = run_report(make_cash_rows({"04": (1.0, 0.0)}),
                             formula_values=values)
        for row_no in FORMULA_ROWS:
            cell = "C%d" % EXCEL_ROW[row_no]
            self.assertIn(cell, hook.number_formats,
                          "公式行 %s（行次%d）未参与格式检查"
                          % (cell, row_no))

    def test_missing_formula_value_skipped(self):
        """读不到公式值时跳过，不报错、不误设格式。"""
        hook, _ = run_report(make_cash_rows({"04": (1.0, 0.0)}))
        self.assertNotIn("D18", hook.number_formats)

    def test_non_numeric_formula_value_skipped(self):
        """公式返回错误值（如 #DIV/0!）时跳过，不崩溃。"""
        hook, _ = run_report(make_cash_rows({"04": (1.0, 0.0)}),
                             formula_values={"D18": "#DIV/0!"})
        self.assertNotIn("D18", hook.number_formats)


if __name__ == "__main__":
    unittest.main(verbosity=2)
