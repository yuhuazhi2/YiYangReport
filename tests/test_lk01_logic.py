# -*- coding: utf-8 -*-
"""LK01 资产负债表的逻辑测试。

运行方式（在项目根目录执行）：
    python -m unittest discover -s tests -v

【本文件测什么？】
    1. 每一行是否按配置写对了格子与数值
       （期末取本期 me、年初取本年1月 mb）
    2. 会计期间是否写入 C4、且用的是文本写入
    3. 取不到数据时是否**不写 0**（保持空白）
    4. 取数抛异常时是否不崩
    5. 【结构校验】cell_map 与 ROW_SPECS 是否一致
    6. 【结构校验】配置里的单元格地址在模板中是否落在有项目名的行上

    第 5、6 条不是"跑一遍看看对不对",而是**长期护栏**：
    以后新增行、改地址时若写错,测试会立刻失败,不必等生成报表才发现。

【不测什么？】
    真实 Excel 的写入效果（格式、日期误判）—— 那必须用真 Excel,
    属手工验证范畴,见 docs/报表开发规范.md 的说明。
"""
from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path

# 让测试既能通过 discover 运行,也能单独 python tests/xxx.py 运行
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from tests.fakes import FakeCtx, FakeHook, approx, make_rows  # noqa: E402

from reports.lk01_sample import (                          # noqa: E402
    PENDING_ROWS,
    ROW_SPECS,
    SUM_EXCLUDED_ROWS,
    SUM_ROWS_TODO,
    LK01Report,
)


def run_report(rows=None, month=5, fail=False):
    """跑一遍 LK01 的 fetch + write,返回 (hook, ctx)。

    为什么不调 generate()？
        因为 generate() 还会调 verify(),而 verify 依赖 hook 能读回单元格。
        FakeHook 的 get() 能读,所以其实也能跑 —— 但这里只关注
        "取数与写入",保持测试聚焦。
    """
    ctx = FakeCtx(rows or [], month=month, fail=fail)
    hook = FakeHook()
    rep = LK01Report()
    rep.fetch(ctx)
    rep.write(hook, ctx)
    return hook, ctx


class TestRowValues(unittest.TestCase):
    """验证每行的取数与写入。"""

    def test_all_rows_may(self):
        """5月：每行期末取本期 me、年初取1月 mb。"""
        # 造数时 mb 与 me 刻意不同,且本期 mb=0：
        # 若程序误取本期 mb 当年初,结果会是 0 而失败。
        codes = {c: (1000.0 + i * 100, 2000.0 + i * 100)
                 for i, c in enumerate(
                     sorted({c for s in ROW_SPECS for c in s.codes}))}
        rows = make_rows(codes, month=5)
        hook, _ = run_report(rows, month=5)

        for spec in ROW_SPECS:
            # 该行所有科目的期末值之和 / 年初值之和
            exp_end = sum(codes[c][1] for c in spec.codes)
            exp_begin = sum(codes[c][0] for c in spec.codes)

            self.assertTrue(
                approx(hook.get(spec.end_cell), exp_end),
                f"{spec.key} 期末格 {spec.end_cell} 期望 {exp_end},"
                f" 实际 {hook.get(spec.end_cell)}")
            self.assertTrue(
                approx(hook.get(spec.begin_cell), exp_begin),
                f"{spec.key} 年初格 {spec.begin_cell} 期望 {exp_begin},"
                f" 实际 {hook.get(spec.begin_cell)}")

    def test_january_same_row_two_values(self):
        """1月：同一行记录同时提供 me(期末) 与 mb(年初)。"""
        codes = {c: (800.0, 900.0) for c in
                 sorted({c for s in ROW_SPECS for c in s.codes})}
        rows = make_rows(codes, month=1)
        hook, _ = run_report(rows, month=1)

        for spec in ROW_SPECS:
            exp_end = sum(codes[c][1] for c in spec.codes)
            exp_begin = sum(codes[c][0] for c in spec.codes)
            self.assertTrue(approx(hook.get(spec.end_cell), exp_end))
            self.assertTrue(approx(hook.get(spec.begin_cell), exp_begin))

    def test_intangible_assets_net(self):
        """无形资产 = 1701 - 1702（累计摊销）- 1703（减值准备）。

        【验证点】
            1. 两个减项都生效（若漏掉 1703，结果会偏大）；
            2. 与固定资产不同，本行**没有加项**，若误把 1702/1703
               当加项，结果会明显偏大；
            3. 累计摊销无记录、减值准备无记录时都应正常写入。
        """
        rows = make_rows({
            "1701": (50000.0, 60000.0),    # 无形资产原值
            "1702": (10000.0, 15000.0),    # 累计摊销
            "1703": (2000.0, 3000.0),      # 减值准备
        }, month=5)
        hook, _ = run_report(rows, month=5)

        # 期末 = 60000 - 15000 - 3000 = 42000
        self.assertTrue(approx(hook.get("C55"), 42000.0),
                        f"无形资产期末应为 42000,实际 {hook.get('C55')}")
        # 年初 = 50000 - 10000 - 2000 = 38000
        self.assertTrue(approx(hook.get("D55"), 38000.0),
                        f"无形资产年初应为 38000,实际 {hook.get('D55')}")

    def test_intangible_assets_without_deductions(self):
        """累计摊销/减值准备都无记录时,无形资产 = 1701 本身（不能空白）。

        【边界情形】新建账套或尚未开始摊销时，1702/1703 可能无记录。
        此时减项贡献 0，C55 应等于 1701，不能因此不写。
        """
        rows = make_rows({"1701": (50000.0, 60000.0)}, month=5)
        hook, _ = run_report(rows, month=5)

        self.assertTrue(approx(hook.get("C55"), 60000.0))
        self.assertTrue(approx(hook.get("D55"), 50000.0))

    def test_1606_goes_to_both_rows(self):
        """1606 同时写入 C47（参与合计）与 C49（明细，不参与合计）。

        【为什么要单独测？】
            一个科目归属两行是本表首次出现的情况，容易被误认为
            "只能归属一行"。若 fetch() 的科目映射写成单值字典，
            后写的一行会覆盖前一行，导致其中一行为空。

            本表用 "科目 → [(行, 符号)]" 的多值映射，
            本条就是验证这个设计确实生效。
        """
        rows = make_rows({
            "1601": (100000.0, 120000.0),
            "1606": (2000.0, 3000.0),
            "1602": (30000.0, 40000.0),
            "1603": (5000.0, 8000.0),
        }, month=5)
        hook, _ = run_report(rows, month=5)

        # C47 = 120000 + 3000 - 40000 - 8000 = 75000（含 1606）
        self.assertTrue(approx(hook.get("C47"), 75000.0))
        self.assertTrue(approx(hook.get("D47"), 67000.0))
        # C49 = 1606 本身 = 3000（仅明细，不重复计入合计）
        self.assertTrue(approx(hook.get("C49"), 3000.0),
                        f"C49 应等于 1606 本身 3000,实际 {hook.get('C49')}")
        self.assertTrue(approx(hook.get("D49"), 2000.0))

    def test_fixed_assets_full_formula(self):
        """固定资产 = 1601 + 1606（清理）- 1602（折旧）- 1603（减值）。

        【验证点】
            1. 加项 1606 生效（若漏掉，结果会偏小）；
            2. 两个减项都生效（若漏掉，结果会偏大）；
            3. 加减方向正确（若把 1606 当成减项、或把 1602 当成加项，
               结果会明显偏离）；
            4. 这是本表第一个"同时含多个加项与多个减项"的行。
        """
        rows = make_rows({
            "1601": (100000.0, 120000.0),   # 固定资产原值
            "1606": (2000.0, 3000.0),       # 固定资产清理（加项）
            "1602": (30000.0, 40000.0),     # 累计折旧（减项）
            "1603": (5000.0, 8000.0),       # 减值准备（减项）
        }, month=5)
        hook, _ = run_report(rows, month=5)

        # 期末 = 120000 + 3000 - 40000 - 8000 = 75000
        self.assertTrue(approx(hook.get("C47"), 75000.0),
                        f"固定资产期末应为 75000,实际 {hook.get('C47')}")
        # 年初 = 100000 + 2000 - 30000 - 5000 = 67000
        self.assertTrue(approx(hook.get("D47"), 67000.0),
                        f"固定资产年初应为 67000,实际 {hook.get('D47')}")

    def test_fixed_assets_without_cleanup(self):
        """1606 固定资产清理无记录时,固定资产 = 1601 - 1602 - 1603。

        【边界情形】固定资产清理有余额才正常，无余额很常见
        （清理完毕或从未发生）。此时加项贡献 0，仍应正常写入。
        """
        rows = make_rows({
            "1601": (100000.0, 120000.0),
            "1602": (30000.0, 40000.0),
            "1603": (5000.0, 8000.0),
        }, month=5)
        hook, _ = run_report(rows, month=5)

        self.assertTrue(approx(hook.get("C47"), 72000.0))
        self.assertTrue(approx(hook.get("D47"), 65000.0))

    def test_fixed_assets_without_impairment(self):
        """减值准备无记录时,固定资产 = 1601 - 1602（不能空白）。

        【边界情形】很多企业并未计提减值准备，1603 当期无记录是正常的。
        此时减项贡献 0，C47 应等于 1601-1602，不能因此不写。
        """
        rows = make_rows({
            "1601": (100000.0, 120000.0),
            "1602": (30000.0, 40000.0),
        }, month=5)
        hook, _ = run_report(rows, month=5)

        self.assertTrue(approx(hook.get("C47"), 80000.0))
        self.assertTrue(approx(hook.get("D47"), 70000.0))

    def test_three_deduction_codes_not_crossed(self):
        """三个减项科目各归各行,互不串用。

        当前表中有三组"主科目 - 减项"：
            应收账款   1122  - 123101
            其他应收款 1221  - 123102
            固定资产   1601  - 1602
        三者主科目与减项都不同，本条同时给全部数据，
        验证每行只扣自己的减项。
        """
        rows = make_rows({
            "1122": (10000.0, 12000.0),
            "123101": (1000.0, 2000.0),
            "1221": (5000.0, 6000.0),
            "123102": (200.0, 350.0),
            "1601": (100000.0, 120000.0),
            "1606": (2000.0, 3000.0),
            "1602": (30000.0, 40000.0),
            "1603": (5000.0, 8000.0),
        }, month=5)
        hook, _ = run_report(rows, month=5)

        self.assertTrue(approx(hook.get("C17"), 10000.0))
        self.assertTrue(approx(hook.get("D17"), 9000.0))
        self.assertTrue(approx(hook.get("C21"), 5650.0))
        self.assertTrue(approx(hook.get("D21"), 4800.0))
        # 固定资产 = 120000+3000-40000-8000 / 100000+2000-30000-5000
        self.assertTrue(approx(hook.get("C47"), 75000.0))
        self.assertTrue(approx(hook.get("D47"), 67000.0))

    def test_other_receivable_is_net_of_its_own_bad_debt(self):
        """其他应收款 = 1221 - 123102（注意不是 123101）。

        【为什么这条测试特别重要？】
            应收账款用 123101、其他应收款用 123102，两个科目编码
            只差最后一位，肉眼极易看错。若写错，报表仍会输出一个
            貌似合理的数字（因为两笔坏账准备金额通常不等，但都正常），
            只有拿计算器逐笔核对才能发现。
        """
        rows = make_rows({
            "1221": (5000.0, 6000.0),      # 其他应收款
            "123102": (200.0, 350.0),      # 其他应收款的坏账准备
            "123101": (9999.0, 9999.0),    # 应收账款的坏账准备，不应参与本行
        }, month=5)
        hook, _ = run_report(rows, month=5)

        # 期末 = 6000 - 350 = 5650
        self.assertTrue(approx(hook.get("C21"), 5650.0),
                        f"其他应收款期末应为 5650,实际 {hook.get('C21')}")
        # 年初 = 5000 - 200 = 4800
        self.assertTrue(approx(hook.get("D21"), 4800.0),
                        f"其他应收款年初应为 4800,实际 {hook.get('D21')}")

    def test_two_bad_debts_not_crossed(self):
        """两组坏账准备各归各行,互不串用。

        同时给应收账款与其他应收款数据，验证：
            应收票据/应收账款 用 123101
            其他应收款        用 123102
        若两行错用了对方的减项，本条会失败。
        """
        rows = make_rows({
            "1122": (10000.0, 12000.0),
            "123101": (1000.0, 2000.0),
            "1221": (5000.0, 6000.0),
            "123102": (200.0, 350.0),
        }, month=5)
        hook, _ = run_report(rows, month=5)

        # 应收账款 = 12000-2000 / 10000-1000
        self.assertTrue(approx(hook.get("C17"), 10000.0))
        self.assertTrue(approx(hook.get("D17"), 9000.0))
        # 其他应收款 = 6000-350 / 5000-200
        self.assertTrue(approx(hook.get("C21"), 5650.0))
        self.assertTrue(approx(hook.get("D21"), 4800.0))

    def test_receivable_is_net_of_bad_debt(self):
        """应收账款 = 1122 - 123101（坏账准备）。

        【为什么要专门测？】
            这是本表第一行"减法"口径。加法的错相对容易看出来
            （数字偏大）；减法的错更隐蔽——若把减项也当成加项，
            结果会变成"1122 + 坏账准备"，比正确值还大，但报表上
            只是个合计数，看不出方向。
        """
        rows = make_rows({
            "1122": (10000.0, 12000.0),    # 应收账款
            "123101": (500.0, 800.0),      # 坏账准备
        }, month=5)
        hook, _ = run_report(rows, month=5)

        # 期末 = 12000 - 800 = 11200
        self.assertTrue(approx(hook.get("C17"), 11200.0),
                        f"应收账款期末应为 11200,实际 {hook.get('C17')}")
        # 年初 = 10000 - 500 = 9500
        self.assertTrue(approx(hook.get("D17"), 9500.0),
                        f"应收账款年初应为 9500,实际 {hook.get('D17')}")

    def test_receivable_without_bad_debt(self):
        """坏账准备无记录时,应收账款 = 1122 本身（不因缺减项而不写）。

        【边界情形】坏账准备科目当期无余额是正常的。
        此时减项贡献 0，加项仍在，C17 应等于 1122 的金额，不能空白。
        """
        rows = make_rows({"1122": (1000.0, 1500.0)}, month=5)
        hook, _ = run_report(rows, month=5)

        self.assertTrue(approx(hook.get("C17"), 1500.0))
        self.assertTrue(approx(hook.get("D17"), 1000.0))

    def test_receivable_only_bad_debt_skipped(self):
        """只有坏账准备、没有 1122 记录时,不写负数。

        【为什么不写？】
            数据异常（或科目设错）时若照算，C17 会变成 -800 这种
            不可能出现的值。宁可留空让维护者查，也不能写一个
            看起来像"负资产"的数。
        """
        rows = make_rows({"123101": (500.0, 800.0)}, month=5)
        hook, _ = run_report(rows, month=5)

        self.assertIsNone(hook.get("C17"),
                          f"不应写入负数,实际 {hook.get('C17')}")
        self.assertIsNone(hook.get("D17"))

    def test_other_rows_unchanged_by_exclude_support(self):
        """引入减项能力后,其他行的加法口径不受影响。"""
        codes = {c: (100.0, 200.0) for c in
                 sorted({c for s in ROW_SPECS for c in (*s.codes, *s.exclude_codes)})}
        rows = make_rows(codes, month=5)
        hook, _ = run_report(rows, month=5)

        for spec in ROW_SPECS:
            exp_end = (sum(codes[c][1] for c in spec.codes)
                       - sum(codes[c][1] for c in spec.exclude_codes))
            self.assertTrue(
                approx(hook.get(spec.end_cell), exp_end),
                f"{spec.key} 期末期望 {exp_end},实际 {hook.get(spec.end_cell)}")

    def test_money_fund_excludes_1012(self):
        """货币资金必须排除 1012（1012 归入交易性金融资产）。"""
        rows = make_rows({"1001": (100.0, 200.0),
                          "1002": (300.0, 400.0),
                          "1012": (9000.0, 9000.0)}, month=5)
        hook, _ = run_report(rows, month=5)

        # 货币资金 = 1001 + 1002 = 600 / 400
        self.assertTrue(approx(hook.get("C11"), 600.0))
        self.assertTrue(approx(hook.get("D11"), 400.0))
        # 交易性金融资产 = 1012 = 9000
        self.assertTrue(approx(hook.get("C13"), 9000.0))
        self.assertTrue(approx(hook.get("D13"), 9000.0))

    def test_inventory_sums_four_codes(self):
        """存货 = 1401 + 1403 + 1405 + 1408 四科目汇总。

        【为什么要单独测这一行？】
            存货是当前唯一"由 4 个科目组成"的行（其他行最多 2 个）。
            多科目汇总最容易出的问题是"漏掉其中一个"或"多加了别的",
            而报表上只会显示一个合计数,肉眼很难发现少了哪一项。
        """
        rows = make_rows({
            "1401": (100.0, 200.0),    # 材料采购
            "1403": (300.0, 400.0),    # 原材料
            "1405": (500.0, 600.0),    # 库存商品
            "1408": (700.0, 800.0),    # 委托加工物资
            # 下面这个不在存货范围内,用于确认不会被误加进来
            "1402": (9999.0, 9999.0),
        }, month=5)
        hook, _ = run_report(rows, month=5)

        # 期末 = 200+400+600+800 = 2000
        self.assertTrue(approx(hook.get("C23"), 2000.0),
                        f"存货期末应为 2000,实际 {hook.get('C23')}")
        # 年初 = 100+300+500+700 = 1600
        self.assertTrue(approx(hook.get("D23"), 1600.0),
                        f"存货年初应为 1600,实际 {hook.get('D23')}")

    def test_no_data_writes_nothing(self):
        """无任何记录时,所有金额格保持空白（不写 0）。"""
        hook, _ = run_report([], month=5)
        for spec in ROW_SPECS:
            self.assertIsNone(
                hook.get(spec.end_cell),
                f"{spec.key} 期末格不应被写入,实际 {hook.get(spec.end_cell)}")
            self.assertIsNone(
                hook.get(spec.begin_cell),
                f"{spec.key} 年初格不应被写入,实际 {hook.get(spec.begin_cell)}")

    def test_partial_data(self):
        """只有本期记录、没有1月记录：只写期末格,年初格留空。

        这是"年中建账"可能出现的真实情形,两格必须独立判断。
        """
        # 只造本期记录（iperiod=5）,没有 iperiod=1
        rows = [(c, 5, 0.0, 500.0)
                for c in sorted({c for s in ROW_SPECS for c in s.codes})]
        hook, _ = run_report(rows, month=5)

        for spec in ROW_SPECS:
            self.assertIsNotNone(hook.get(spec.end_cell))
            self.assertIsNone(hook.get(spec.begin_cell))

    def test_fetch_failure_not_crash(self):
        """取数抛异常时不崩溃,且不写任何金额。"""
        hook, _ = run_report([], month=5, fail=True)
        for spec in ROW_SPECS:
            self.assertIsNone(hook.get(spec.end_cell))
            self.assertIsNone(hook.get(spec.begin_cell))

    def test_codes_deduplicated_in_query(self):
        """所有行的科目应合并成一次查询（去重）。"""
        _, ctx = run_report([], month=5)
        self.assertEqual(len(ctx.fetch_calls), 1, "应只发一次查询")

        params = ctx.fetch_calls[0]["params"]
        codes_in_sql = params[3:]          # 前3个是期间参数
        self.assertEqual(len(codes_in_sql), len(set(codes_in_sql)),
                         f"科目参数有重复：{codes_in_sql}")


class TestPeriodCell(unittest.TestCase):
    """验证会计期间写入 C4。"""

    def test_period_written_as_text(self):
        hook, _ = run_report([], month=5)
        self.assertEqual(hook.get("C4"), "2026年5月")

        # 必须用 set_text 写：C4 是金额格式格,
        # 普通写入会被 Excel 当成日期存成序列号（历史坑）。
        kinds = [e[0] for e in hook.events if e[1] == "C4"]
        self.assertIn("set_text", kinds,
                      f"C4 未用 set_text 写入,实际事件：{hook.events}")
        self.assertNotIn("set", kinds,
                         "C4 不应使用普通 set 写入（会被识别成日期）")


class TestConfigConsistency(unittest.TestCase):
    """结构校验：防止配置与 cell_map / 模板脱节。"""

    def test_cell_map_matches_specs(self):
        """ROW_SPECS 里的每个格子都应登记在 cell_map 中。

        【为什么重要？】
            cell_map 决定父类 verify() 读回哪些格做勾稽核对。
            新增行时若忘了登记 cell_map,该行数据不进核对清单,
            日后"勾稽通过"可能是假象（漏检）。
        """
        cm_values = set(LK01Report.cell_map.values())
        for spec in ROW_SPECS:
            self.assertIn(spec.end_cell, cm_values,
                          f"{spec.key} 的期末格 {spec.end_cell} 未登记 cell_map")
            self.assertIn(spec.begin_cell, cm_values,
                          f"{spec.key} 的年初格 {spec.begin_cell} 未登记 cell_map")

    def test_no_duplicate_cells(self):
        """不同行不得占用同一格子（否则后写的会覆盖先写的）。"""
        seen: dict[str, str] = {}
        for spec in ROW_SPECS:
            for cell in (spec.end_cell, spec.begin_cell):
                self.assertNotIn(
                    cell, seen,
                    f"格子 {cell} 被 {seen.get(cell)} 和 {spec.key} 同时使用")
                seen[cell] = spec.key

    def test_pending_rows_not_conflict(self):
        """待定项不得与已实现项占用同一格子。

        【为什么需要这条？】
            PENDING_ROWS 记录的是"口径未明确、暂不实现"的行。
            将来补实现时，若不小心与已有行撞格子，后写的会覆盖前写的，
            而报表上看不出异常（数字还在，只是属于别的项目）。
            这条在冲突发生的**当下**就报错，比事后查账便宜得多。
        """
        implemented = set()
        for spec in ROW_SPECS:
            implemented.add(spec.end_cell)
            implemented.add(spec.begin_cell)

        for name, cells, row_no in PENDING_ROWS:
            # cells 形如 "C19/D19"，拆成两个地址
            for cell in cells.split("/"):
                self.assertNotIn(
                    cell, implemented,
                    f"待定项「{name}」的 {cell} 与已实现行冲突"
                    f"（行次 {row_no}）")

    def test_pending_rows_unique(self):
        """待定项之间也不得重复（同一格被两个待定项占用）。"""
        seen: dict[str, str] = {}
        for name, cells, _ in PENDING_ROWS:
            for cell in cells.split("/"):
                self.assertNotIn(
                    cell, seen,
                    f"待定项 {cell} 被「{seen.get(cell)}」和「{name}」同时占用")
                seen[cell] = name

    def test_summary_excluded_rows_have_in_subtotal_false(self):
        """汇总时须排除的行，若已实现则必须 in_subtotal=False。

        【这条测试的口径与实现方式都变过一次，值得记录】

          第 1 版（错误）：
            用"该行不得出现在 ROW_SPECS"来防止重复计算。
            问题：等于连"填数"也禁止了，但客户其实需要看到
            C49「其中：固定资产清理」的明细金额。

          第 2 版（当前）：
            允许该行存在于 ROW_SPECS（正常填数），
            但必须 in_subtotal=False（不参与合计）。
            本测试守住这条约束。

        【为什么重复计算很难发现？】
            第 49 行的 1606 已经计入 C47。若汇总时再加一次 C49，
            非流动资产合计会偏大恰好一个 1606 的金额，
            而报表上每个单元格看起来都"很正常"。
        """
        excluded_row_nos = {row_no for row_no, _ in SUM_EXCLUDED_ROWS}

        for spec in ROW_SPECS:
            # 从格位里取行号（如 "C49" → 49）
            row_no = int("".join(ch for ch in spec.end_cell if ch.isdigit()))
            if row_no in excluded_row_nos:
                self.assertFalse(
                    spec.in_subtotal,
                    f"「{spec.label}」对应第 {row_no} 行，属汇总须排除的明细行，"
                    f"应设 in_subtotal=False，否则合计会重复计算")

    def test_all_other_rows_in_subtotal(self):
        """除排除清单外，其余行都应参与汇总（in_subtotal=True）。

        防止有人为了"先让它不出数"而随手把某行设成 False，
        导致合计静默变少。
        """
        excluded_row_nos = {row_no for row_no, _ in SUM_EXCLUDED_ROWS}
        for spec in ROW_SPECS:
            row_no = int("".join(ch for ch in spec.end_cell if ch.isdigit()))
            if row_no not in excluded_row_nos:
                self.assertTrue(
                    spec.in_subtotal,
                    f"「{spec.label}」({spec.end_cell}) 应参与汇总，"
                    f"却被设为 in_subtotal=False")

    def test_summary_todo_not_conflict(self):
        """待实现汇总行的格位,不得与现有已实现行冲突。"""
        implemented_cells = set()
        for spec in ROW_SPECS:
            implemented_cells.add(spec.end_cell)
            implemented_cells.add(spec.begin_cell)

        for row_no, name, cells in SUM_ROWS_TODO:
            for cell in cells.split("/"):
                self.assertNotIn(
                    cell, implemented_cells,
                    f"待实现汇总行「{name}」的 {cell} 与现有行冲突")

    def test_codes_are_strings(self):
        """科目编码必须是字符串 —— 若是数字,1001 会变成 '1001' 之外的形式。"""
        for spec in ROW_SPECS:
            for c in spec.codes:
                self.assertIsInstance(
                    c, str,
                    f"{spec.key} 的科目编码 {c!r} 不是字符串")

    def test_cells_in_template_have_labels(self):
        """配置里的格子所在行,模板 A 列应有项目名。

        【为什么重要？】
            这条能抓住"地址写错行"——例如本该写 C21 却写成 C12,
            而 C12 那行是英文项目名或空行,一看就露馅。

        若模板文件不存在（例如未从仓库获取 .xls）,
        本测试自动跳过,不影响其他测试。
        """
        try:
            import xlrd
        except ImportError:
            self.skipTest("未安装 xlrd")

        from config.settings import DEFAULT_TEMPLATE, PROJECT_ROOT

        tpl = PROJECT_ROOT / DEFAULT_TEMPLATE
        if not tpl.exists():
            self.skipTest(f"模板不存在：{tpl}")

        wb = xlrd.open_workbook(str(tpl))
        sh = wb.sheet_by_name("LK01")

        for spec in ROW_SPECS:
            for cell in (spec.end_cell, spec.begin_cell):
                row = int("".join(ch for ch in cell if ch.isdigit()))
                a_val = str(sh.cell_value(row - 1, 0)).strip()
                self.assertTrue(
                    a_val,
                    f"{spec.key} 的 {cell} 位于第 {row} 行,"
                    f"但模板 A 列是空的 —— 可能写错了行")


if __name__ == "__main__":
    unittest.main(verbosity=2)