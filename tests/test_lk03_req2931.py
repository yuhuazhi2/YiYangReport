# -*- coding: utf-8 -*-
"""LK03 需求第 29~31 条。

    29  C43 / D43  行次 35  汇率变动影响            gl_cashtable '23'，md - mc
    30  C44 / D44  行次 36  现金及现金等价物净增加额  不做处理（需求25 已登记）
    31  C45 / D45  行次 37  期初现金及现金等价物余额  GL_accsum 1001/1002 的 mb

【本文件测三件容易搞混的事】
    ① 需求29 的方向必须是 md - mc（保留正负号），不能像"支付"类取反
    ② 需求31 的数据源不是 gl_cashtable，而是 GL_accsum 的余额
    ③ 需求30 不重复登记，登记格数仍等于公式行数 × 2
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
    make_balance_rows,
    make_cash_rows,
    run_report,
)

from reports.lk03_cashflow import (        # noqa: E402
    BALANCE_SPECS,
    NEGATIVE_NUMBER_FORMAT,
    EXCEL_ROW,
    FORMULA_ROWS,
    INTENTIONALLY_UNTOUCHED,
    PENDING_ROWS,
    ROW_SPECS,
    LK03Report,
)


def _cash_call_with_item(ctx, item):
    """从查询列表中挑出"现金流量项目参数中含 item"的那一条。"""
    for c in ctx.fetch_calls:
        if item in (c["params"] or []):
            return c
    raise AssertionError("未找到包含项目编码 %r 的查询" % (item,))


def _balance_call(ctx, prev=False):
    """挑出查 GL_accsum 的那一条（可指定本期/上期）。"""
    for c in ctx.fetch_calls:
        if "GL_accsum" not in c["sql"]:
            continue
        is_prev = c["report_code"].endswith(("-上年", "-上期"))
        if is_prev == prev:
            return c
    raise AssertionError("未找到 GL_accsum 查询（prev=%r）" % (prev,))


class TestReq29Config(unittest.TestCase):
    """需求29：行次 35 是普通明细行，走 gl_cashtable 的 '23'。"""

    def _spec(self):
        return next(s for s in ROW_SPECS
                    if s.key == "汇率变动对现金及现金等价物的影响")

    def test_item_is_23(self):
        self.assertEqual(self._spec().cash_items, ("23",))

    def test_excel_row_is_43(self):
        self.assertEqual(self._spec().excel_row, EXCEL_ROW[35])
        self.assertEqual(EXCEL_ROW[35], 43)

    def test_direction_is_inflow(self):
        """★ 方向必须是 md - mc。

        【为什么锁这条？】
            需求原文："汇总 md 值减去 mc 值的差额填入 C43（保留正负符号）"。
            若误配 outflow（mc - md），正负号整体翻转：
            本期汇率损失 -500 会变成 +500，
            而模板净增加额把它直接相加，合计数随之偏差 1000 ——
            两行都"有数字"，只有合计不对。
        """
        self.assertEqual(self._spec().direction, "inflow")

    def test_not_formula_row(self):
        self.assertNotIn(35, FORMULA_ROWS)


class TestReq29Value(unittest.TestCase):
    """需求29：净额 = md - mc，保留正负号。"""

    def test_positive_when_debit_larger(self):
        hook, _ = run_report(make_cash_rows({"23": (5000.0, 2000.0)}))
        self.assertTrue(approx(hook.get("C43"), 3000.0),
                        "C43 应为 5000-2000=3000，实际 %r" % (hook.get("C43"),))

    def test_negative_when_credit_larger(self):
        """★ 贷方大时必须是负数（汇率损失）。

        【这条是本需求的核心】
            若把方向配成 outflow，这里会得到 +500，
            表面上"有个数字"，实际上符号全错。
        """
        hook, _ = run_report(make_cash_rows({"23": (1000.0, 1500.0)}))
        self.assertTrue(approx(hook.get("C43"), -500.0),
                        "C43 应为 1000-1500=-500，实际 %r" % (hook.get("C43"),))

    def test_md_only_positive(self):
        hook, _ = run_report(make_cash_rows({"23": (800.0, 0.0)}))
        self.assertTrue(approx(hook.get("C43"), 800.0))

    def test_mc_only_negative(self):
        hook, _ = run_report(make_cash_rows({"23": (0.0, 800.0)}))
        self.assertTrue(approx(hook.get("C43"), -800.0))

    def test_net_zero_written(self):
        hook, _ = run_report(make_cash_rows({"23": (600.0, 600.0)}))
        self.assertEqual(hook.get("C43"), 0.0)

    def test_no_data_blank(self):
        hook, _ = run_report([])
        self.assertIsNone(hook.get("C43"))
        self.assertIsNone(hook.get("D43"))

    def test_fetch_failure_not_crash(self):
        hook, _ = run_report([], fail=True)
        self.assertIsNone(hook.get("C43"))

    def test_other_items_not_mixed(self):
        """其它已实现的项目不得混入 C43。"""
        hook, _ = run_report(make_cash_rows({
            "22": (0.0, 4000.0),   # 筹资流出，属 C40
            "23": (700.0, 0.0),
        }))
        self.assertTrue(approx(hook.get("C43"), 700.0),
                        "C43 应只含 '23'=700，实际 %r" % (hook.get("C43"),))


class TestReq29Prev(unittest.TestCase):
    """需求29：D43 取上一期，规则与 D10 相同。"""

    def test_d43_prev_month(self):
        hook, _ = run_report(
            make_cash_rows({"23": (1.0, 0.0)}, month=8),
            month=8,
            prev_rows=make_cash_rows({"23": (900.0, 400.0)}, month=7))
        self.assertTrue(approx(hook.get("D43"), 500.0),
                        "D43 应为 900-400=500，实际 %r" % (hook.get("D43"),))

    def test_d43_prev_negative(self):
        hook, _ = run_report(
            make_cash_rows({"23": (1.0, 0.0)}, month=8),
            month=8,
            prev_rows=make_cash_rows({"23": (100.0, 300.0)}, month=7))
        self.assertTrue(approx(hook.get("D43"), -200.0),
                        "D43 应为 100-300=-200，实际 %r" % (hook.get("D43"),))

    def test_d43_january_cross_year(self):
        hook, ctx = run_report(
            make_cash_rows({"23": (1.0, 0.0)}, month=1),
            month=1,
            prev_rows=make_cash_rows({"23": (300.0, 100.0)},
                                     month=12, year=2025),
            prev_database="UFDATA_001_2025")
        self.assertTrue(approx(hook.get("D43"), 200.0))

        pc = _cash_call_with_item(ctx, "23")
        self.assertEqual(pc["params"][0], 202601)
        prev = [c for c in ctx.fetch_calls
                if "23" in (c["params"] or [])
                and c["report_code"].endswith("-上期")]
        self.assertTrue(prev, "应发起 '23' 的上期查询")
        self.assertEqual(prev[0]["params"][0], 202512)
        self.assertEqual(prev[0]["database"], "UFDATA_001_2025")

    def test_d43_no_data_blank(self):
        hook, _ = run_report(make_cash_rows({"23": (100.0, 0.0)}),
                             month=5, prev_rows=[])
        self.assertTrue(approx(hook.get("C43"), 100.0))
        self.assertIsNone(hook.get("D43"))


class TestReq29Structure(unittest.TestCase):
    """需求29：登记与查询口径。"""

    def test_cells_in_cell_map(self):
        values = set(LK03Report.cell_map.values())
        for cell in ("C43", "D43"):
            self.assertIn(cell, values, "%s 未登记 cell_map" % (cell,))

    def test_not_pending(self):
        pending = {row_no for _, _, row_no in PENDING_ROWS}
        self.assertNotIn(35, pending, "行次 35 已实现，应从 PENDING_ROWS 移除")

    def test_query_uses_cashtable(self):
        """本行是现金流量项目，必须查 gl_cashtable，不是 GL_accsum。"""
        _, ctx = run_report(make_cash_rows({"23": (1.0, 0.0)}))
        call = _cash_call_with_item(ctx, "23")
        self.assertIn("gl_cashtable", call["sql"])
        self.assertNotIn("GL_accsum", call["sql"])
        self.assertIn("iYPeriod = ?", call["sql"])
        self.assertNotIn("BETWEEN", call["sql"])


class TestReq31Config(unittest.TestCase):
    """需求31：行次 37 走 GL_accsum 余额线。"""

    def _spec(self):
        return next(b for b in BALANCE_SPECS
                    if b.key == "期初现金及现金等价物余额")

    def test_account_codes(self):
        """科目必须是 1001、1002 两个，不含 1012。"""
        self.assertEqual(self._spec().account_codes, ("1001", "1002"))

    def test_1012_not_included(self):
        """★ 1012 其他货币资金不得纳入。

        LK01 已确认 1012 归入"交易性金融资产"。
        若这里也算进来，两张表的"现金"口径就不一致，
        而各自看起来都正常，只有交叉核对时才对不上。
        """
        self.assertNotIn("1012", self._spec().account_codes)

    def test_excel_row_is_45(self):
        self.assertEqual(self._spec().excel_row, EXCEL_ROW[37])
        self.assertEqual(EXCEL_ROW[37], 45)

    def test_not_formula_row(self):
        self.assertNotIn(37, FORMULA_ROWS)

    def test_not_pending(self):
        pending = {row_no for _, _, row_no in PENDING_ROWS}
        self.assertNotIn(37, pending, "行次 37 已实现，应从 PENDING_ROWS 移除")


class TestReq31Value(unittest.TestCase):
    """需求31：C45 = 1001 与 1002 的 mb 合计。"""

    def test_sum_of_two_accounts(self):
        hook, _ = run_report(
            make_cash_rows({"01": (1.0, 0.0)}),
            balance_rows=make_balance_rows({"1001": 3000.0,
                                            "1002": 120000.0}))
        self.assertTrue(approx(hook.get("C45"), 123000.0),
                        "C45 应为 3000+120000=123000，实际 %r" % (hook.get("C45"),))

    def test_single_account(self):
        hook, _ = run_report(
            make_cash_rows({"01": (1.0, 0.0)}),
            balance_rows=make_balance_rows({"1002": 5000.0}))
        self.assertTrue(approx(hook.get("C45"), 5000.0))

    def test_zero_balance_written_as_zero(self):
        """★ 有记录且余额为 0 时必须写 0（不是留空）。

        【为什么这条重要？】
            科目存在、余额为 0 是准确事实（月初账上真没钱）；
            而"查不到该科目"是取数问题。
            若两者都表现为空白，财务无法区分"确实没余额"
            与"程序没取到数"。
        """
        hook, _ = run_report(
            make_cash_rows({"01": (1.0, 0.0)}),
            balance_rows=make_balance_rows({"1001": 0.0}))
        self.assertEqual(hook.get("C45"), 0.0,
                         "有记录且余额 0 时应写 0，实际 %r" % (hook.get("C45"),))

    def test_no_data_blank(self):
        hook, _ = run_report(make_cash_rows({"01": (1.0, 0.0)}),
                             balance_rows=None)
        self.assertIsNone(hook.get("C45"))

    def test_other_codes_ignored(self):
        """非 1001/1002 的科目不得混入 C45。"""
        hook, _ = run_report(
            make_cash_rows({"01": (1.0, 0.0)}),
            balance_rows=make_balance_rows({
                "1001": 1000.0,
                "1012": 9999.0,     # 其他货币资金，不应计入
                "1122": 8888.0,     # 应收账款，不应计入
            }))
        self.assertTrue(approx(hook.get("C45"), 1000.0),
                        "C45 应只含 1001+1002，实际 %r" % (hook.get("C45"),))

    def test_balance_not_taken_from_cashtable(self):
        """★ 现金流量项目的发生额不得混入 C45。

        两者数据源不同：gl_cashtable 是"本期发生了多少"，
        GL_accsum.mb 是"期初有多少"。若混用，
        C45 会变成一个既非发生额也非余额的数。
        """
        hook, _ = run_report(
            make_cash_rows({"01": (99999.0, 0.0), "23": (88888.0, 0.0)}),
            balance_rows=make_balance_rows({"1001": 100.0}))
        self.assertTrue(approx(hook.get("C45"), 100.0),
                        "C45 应只含余额 100，实际 %r" % (hook.get("C45"),))

    def test_wrong_period_skipped(self):
        """iYPeriod 不符的余额记录必须跳过（纵深防御）。"""
        hook, _ = run_report(
            make_cash_rows({"01": (1.0, 0.0)}),
            balance_rows=[("1001", 202607, 5000.0)])   # 本应是 202605
        self.assertIsNone(hook.get("C45"),
                          "月份不符的余额应被跳过，实际 %r" % (hook.get("C45"),))


class TestReq31Prev(unittest.TestCase):
    """需求31：D45 取上一期的 mb 合计。"""

    def test_d45_prev_month(self):
        hook, _ = run_report(
            make_cash_rows({"01": (1.0, 0.0)}, month=8),
            month=8,
            prev_rows=make_cash_rows({"01": (1.0, 0.0)}, month=7),
            balance_rows=make_balance_rows({"1001": 100.0}, month=8),
            balance_prev_rows=make_balance_rows({"1001": 700.0,
                                                 "1002": 1300.0}, month=7))
        self.assertTrue(approx(hook.get("C45"), 100.0))
        self.assertTrue(approx(hook.get("D45"), 2000.0),
                        "D45 应为 700+1300=2000，实际 %r" % (hook.get("D45"),))

    def test_d45_january_cross_year(self):
        hook, ctx = run_report(
            make_cash_rows({"01": (1.0, 0.0)}, month=1),
            month=1,
            prev_rows=make_cash_rows({"01": (1.0, 0.0)},
                                     month=12, year=2025),
            balance_rows=make_balance_rows({"1001": 50.0}, month=1),
            balance_prev_rows=make_balance_rows({"1001": 400.0},
                                                month=12, year=2025),
            prev_database="UFDATA_001_2025")
        self.assertTrue(approx(hook.get("D45"), 400.0))

        pc = _balance_call(ctx, prev=True)
        self.assertEqual(pc["params"][0], 202512)
        self.assertEqual(pc["database"], "UFDATA_001_2025")

    def test_d45_no_data_blank(self):
        hook, _ = run_report(
            make_cash_rows({"01": (1.0, 0.0)}),
            month=5,
            balance_rows=make_balance_rows({"1001": 100.0}),
            balance_prev_rows=None)
        self.assertTrue(approx(hook.get("C45"), 100.0))
        self.assertIsNone(hook.get("D45"))

    def test_current_and_prev_not_crossed(self):
        """本期与上期余额绝不互相混入。"""
        hook, _ = run_report(
            make_cash_rows({"01": (1.0, 0.0)}, month=5),
            month=5,
            balance_rows=make_balance_rows({"1001": 111111.0}, month=5),
            balance_prev_rows=make_balance_rows({"1001": 7.0}, month=4))
        self.assertTrue(approx(hook.get("C45"), 111111.0))
        self.assertTrue(approx(hook.get("D45"), 7.0))


class TestReq31Query(unittest.TestCase):
    """需求31：SQL 必须是 GL_accsum + mb + 精确科目匹配。"""

    def _call(self):
        _, ctx = run_report(
            make_cash_rows({"01": (1.0, 0.0)}),
            balance_rows=make_balance_rows({"1001": 1.0}))
        return _balance_call(ctx, prev=False)

    def test_uses_accsum(self):
        self.assertIn("GL_accsum", self._call()["sql"])

    def test_selects_mb(self):
        """必须取 mb（期初余额），不能取 md/mc/me。"""
        sql = self._call()["sql"]
        self.assertIn("mb", sql)
        for bad in ("SUM(md)", "SUM(mc)", "me\n"):
            self.assertNotIn(bad, sql)

    def test_single_month(self):
        sql = self._call()["sql"]
        self.assertIn("iYPeriod = ?", sql)
        self.assertNotIn("BETWEEN", sql)

    def test_codes_exact_match(self):
        """必须是 IN 精确匹配，不能用 LIKE 前缀（会重复计入明细科目）。"""
        call = self._call()
        self.assertIn("ccode IN (", call["sql"])
        self.assertNotIn("LIKE", call["sql"])
        self.assertIn("1001", call["params"])
        self.assertIn("1002", call["params"])

    def test_month_param(self):
        self.assertEqual(self._call()["params"][0], 202605)


class TestReq30(unittest.TestCase):
    """需求30：C44 / D44 不做处理（需求25 已登记，不重复）。"""

    def test_still_registered(self):
        for cell in ("C44", "D44"):
            self.assertIn(cell, INTENTIONALLY_UNTOUCHED)

    def test_not_duplicated(self):
        """★ 不得重复登记。

        重复会让 test_lk03_req25 的"登记格数 = 公式行数 × 2"
        长期护栏失败，而失败信息会指向"是不是漏了哪一行"，
        反而误导维护者。
        """
        cells = list(INTENTIONALLY_UNTOUCHED)
        self.assertEqual(len(cells), len(set(cells)),
                         "INTENTIONALLY_UNTOUCHED 有重复项")

    def test_never_written(self):
        """即使全部明细行都有数，C44/D44 也不得写入。"""
        hook, _ = run_report(
            make_cash_rows({
                "23": (1000.0, 0.0),
                "01": (5000.0, 0.0),
            }),
            balance_rows=make_balance_rows({"1001": 800.0}))
        self.assertTrue(approx(hook.get("C43"), 1000.0))   # 前置条件
        self.assertTrue(approx(hook.get("C45"), 800.0))
        self.assertIsNone(hook.get("C44"))
        self.assertIsNone(hook.get("D44"))

    def test_not_in_cell_map(self):
        values = set(LK03Report.cell_map.values())
        self.assertNotIn("C44", values)
        self.assertNotIn("D44", values)


class TestFormulaRowsUntouched(unittest.TestCase):
    """新增两条取数线后，公式行仍不得写入。"""

    def test_formula_rows_untouched_with_all_data(self):
        hook, _ = run_report(
            make_cash_rows({"23": (1000.0, 0.0), "01": (5000.0, 0.0)}),
            balance_rows=make_balance_rows({"1001": 800.0}))
        for row_no in FORMULA_ROWS:
            excel_row = EXCEL_ROW[row_no]
            for col in ("C", "D"):
                cell = "%s%d" % (col, excel_row)
                self.assertIsNone(hook.get(cell),
                                  "%s（行次%d）是公式行，不得写入"
                                  % (cell, row_no))


class TestTemplateConsistency(unittest.TestCase):
    """读模板确认第 43、45 行的项目名与位置。"""

    def _open_lk03(self):
        try:
            import xlrd
        except ImportError:
            self.skipTest("未安装 xlrd")

        from config.settings import DEFAULT_TEMPLATE, PROJECT_ROOT

        tpl = PROJECT_ROOT / DEFAULT_TEMPLATE
        if not tpl.exists():
            self.skipTest("模板不存在：%s" % (tpl,))

        wb = xlrd.open_workbook(str(tpl))
        return wb.sheet_by_name("LK03")

    def test_row43_is_exchange_rate(self):
        sh = self._open_lk03()
        val = str(sh.cell_value(EXCEL_ROW[35] - 1, 0)).strip()
        self.assertIn("汇率变动", val,
                      "第 43 行 A 列应含「汇率变动」，实际 %r" % (val,))

    def test_row45_is_begin_balance(self):
        sh = self._open_lk03()
        val = str(sh.cell_value(EXCEL_ROW[37] - 1, 0)).strip()
        self.assertIn("期初", val,
                      "第 45 行 A 列应含「期初」，实际 %r" % (val,))

    def test_row_order(self):
        """43 < 44 < 45 < 46，且 43/45 都是明细行。"""
        rows = [EXCEL_ROW[n] for n in (35, 36, 37, 38)]
        self.assertEqual(rows, [43, 44, 45, 46])
        self.assertEqual(rows, sorted(rows))


class TestNegativeFormat(unittest.TestCase):
    """负数格的显示格式（20261110 现场反馈）。

    【背景】
        模板金额列用的是会计专用格式：
            _-* #,##0.00_-;\-* #,##0.00_-;_-* "-"_-;_-@_-
        负数段 `\-* ` 里的星号是"重复填充到列宽尽头"，
        于是负数显示成
            -            1,036,332.84
        负号贴左、数字贴右、中间全是空格，列越宽距离越远，
        财务误以为是脏数据。

    【修复】
        只对**确实为负**的格子改用
            #,##0.00;-#,##0.00;0.00
        负数段 `-#,##0.00` 没有填充指令，负号紧贴数字。
    """

    def test_negative_cell_gets_compact_format(self):
        """负数格必须被设置成"负号紧贴"的格式。

        【为什么用 1000 vs 1500？】
            行次 6 是 outflow（mc - md），要给贷方小于借方才得到负数：
            md=1500、mc=1000 → mc-md=-500
        """
        hook, _ = run_report(make_cash_rows({"04": (1500.0, 1000.0)}))
        self.assertTrue(approx(hook.get("C14"), -500.0),
                        "前置条件不成立：C14 应为 -500，实际 %r"
                        % (hook.get("C14"),))
        self.assertIn("C14", hook.number_formats,
                      "负数格 C14 未设置数字格式")
        fmt = hook.number_formats["C14"]
        self.assertIn("-#,##0.00", fmt,
                      "负数段应无填充，实际格式 %r" % (fmt,))
        self.assertNotIn("*", fmt,
                         "格式中不应再有填充指令 *，实际 %r" % (fmt,))

    def test_positive_cell_keeps_template_format(self):
        """★ 正数格不得被改格式。

        【为什么锁这条？】
            模板正数段 `_-* #,##0.00_-` 右端留了字符宽空隙，
            与相邻列的对齐关系是设计好的。整列替换会连带
            改变正数的对齐 —— 而正数本来没有任何显示问题。
        """
        hook, _ = run_report(make_cash_rows({"04": (0.0, 1000.0)}))
        self.assertTrue(approx(hook.get("C14"), 1000.0))
        self.assertNotIn("C14", hook.number_formats,
                         "正数格不应被改格式，实际 %r"
                         % (hook.number_formats.get("C14"),))

    def test_zero_not_formatted(self):
        """净额为 0 时不改格式（原格式零值段显示横杠，属模板设计）。"""
        hook, _ = run_report(make_cash_rows({"04": (500.0, 500.0)}))
        self.assertEqual(hook.get("C14"), 0.0)
        self.assertNotIn("C14", hook.number_formats)

    def test_blank_cell_not_formatted(self):
        """无数据（留空）时不设格式。"""
        hook, _ = run_report([])
        self.assertIsNone(hook.get("C14"))
        self.assertNotIn("C14", hook.number_formats)

    def test_prev_negative_also_formatted(self):
        """D 列（上期）的负数同样处理。"""
        hook, _ = run_report(
            make_cash_rows({"04": (0.0, 1.0)}, month=8),
            month=8,
            prev_rows=make_cash_rows({"04": (2000.0, 500.0)}, month=7))
        self.assertTrue(approx(hook.get("D14"), -1500.0),
                        "前置条件不成立：D14 应为 -1500，实际 %r"
                        % (hook.get("D14"),))
        self.assertIn("D14", hook.number_formats)

    def test_inflow_negative_also_formatted(self):
        """inflow 行也可能为负（退款超过收款），同样处理。"""
        hook, _ = run_report(make_cash_rows({"01": (0.0, 300.0)}))
        self.assertTrue(approx(hook.get("C10"), -300.0))
        self.assertIn("C10", hook.number_formats)

    def test_balance_negative_also_formatted(self):
        """余额类行（如银行透支）为负时同样处理。"""
        hook, _ = run_report(
            make_cash_rows({"01": (1.0, 0.0)}),
            balance_rows=make_balance_rows({"1002": -800.0}))
        self.assertTrue(approx(hook.get("C45"), -800.0),
                        "前置条件不成立：C45 应为 -800，实际 %r"
                        % (hook.get("C45"),))
        self.assertIn("C45", hook.number_formats)

    def test_format_does_not_change_value(self):
        """★ 改格式不得影响数值本身。

        【为什么要锁？】
            格式只是显示层。若误把 value 也改了（例如取绝对值），
            报表上看起来"没有负号"，实际数值已被篡改 ——
            比显示难看严重得多。
        """
        hook, _ = run_report(make_cash_rows({"04": (1500.0, 1000.0)}))
        self.assertTrue(approx(hook.get("C14"), -500.0),
                        "格式调整后数值不得变化，实际 %r"
                        % (hook.get("C14"),))

    def test_format_string_spec(self):
        """格式串结构：正数 / 负数(紧贴) / 零 三段。"""
        for part in ("#,##0.00", "-#,##0.00", "0.00"):
            self.assertIn(part, NEGATIVE_NUMBER_FORMAT,
                          "格式串缺少 %r 段：%r"
                          % (part, NEGATIVE_NUMBER_FORMAT))
        self.assertNotIn("*", NEGATIVE_NUMBER_FORMAT,
                         "格式串不应含填充指令 *")


if __name__ == "__main__":
    unittest.main(verbosity=2)
