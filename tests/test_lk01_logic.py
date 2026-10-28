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

from reports.lk01_sample import ROW_SPECS, LK01Report     # noqa: E402


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