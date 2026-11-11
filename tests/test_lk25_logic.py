# -*- coding: utf-8 -*-
"""LK25(B29) 制造费用明细表的逻辑测试。

运行方式（在项目根目录执行）：
    python -m unittest discover -s tests -v

【本文件测什么？】
    1. 行项目配置与模板版式是否对得上（25 行、行次 1~25、第 9~33 行）
    2. 科目未配置时的行为：**不查库、不写数**，但页眉与会计期间照常填
    3. 科目配好后的取数与写入：
       · 各月数落到正确的列
       · 取数上限是"所选月份"，6~12 月绝不写入
       · 同一行的多个科目累加
       · 减项科目（exclude_codes）生效
       · 无发生额的月份保持空白（不写 0）
    4. **O 列与第 34 行合计不得由程序写入**（模板公式，写了就覆盖公式）
    5. 【结构校验】cell_map 与 ROW_SPECS 一致、格子不重复、模板地址有效

【为什么大量用例要"临时给科目"？】
    本次开发阶段 25 行的科目清单尚未由财务提供，ROW_SPECS 里 codes 全空。
    若测试也只按"空科目"跑，那么"取数分流、月列落位、期间截断"这些
    真正容易写错的逻辑**一条都验证不到** —— 将来补上科目时才发现错，
    代价远高于现在用假科目先锁住行为。

    所以下面用 _with_codes() 临时把假科目填进配置，
    测完即还原（addCleanup 保证即使断言失败也会还原）。

【不测什么？】
    真实 Excel 的写入效果（格式、公式是否被保留）—— 那必须用真 Excel，
    属手工验证范畴，见 docs/报表开发规范.md。
"""
from __future__ import annotations

import dataclasses
import sys
import unittest
from pathlib import Path

# 让测试既能通过 discover 运行，也能单独 python tests/xxx.py 运行
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from tests.fakes import FakeCtx, FakeHook, approx  # noqa: E402

from reports import lk25_manufacturing_cost as mod  # noqa: E402
from reports.lk25_manufacturing_cost import (  # noqa: E402
    COL_OF_MONTH,
    COL_YEAR_TOTAL,
    FIRST_DATA_ROW,
    PERIOD_CELL,
    ROW_SPECS,
    TOTAL_ROW,
    WRITE_TOTAL_ROW,
    WRITE_YEAR_TOTAL,
    LK25Report,
)


def _with_codes(mapping: dict[str, tuple[str, ...]]):
    """把 {行 key: 科目元组} 临时填进 ROW_SPECS，返回原值供还原。

    【为什么改模块级常量，而不是改实例？】
        fetch() 读的是模块级 ROW_SPECS。若只改实例属性，
        取数逻辑根本看不到 —— 测试会"通过"，但什么都没验到。
    """
    patched = []
    for spec in ROW_SPECS:
        patched.append(dataclasses.replace(spec, codes=mapping.get(spec.key, ())))

    original = mod.ROW_SPECS
    mod.ROW_SPECS = tuple(patched)
    return original


class TestRowConfig(unittest.TestCase):
    """行项目配置与模板版式的对应关系。"""

    def test_row_count_matches_template(self):
        """模板数据区共 25 个费用项目（行次 1~25）。"""
        self.assertEqual(len(ROW_SPECS), 25,
                         f"应为 25 行，实际 {len(ROW_SPECS)} 行")

    def test_row_numbers_are_sequential(self):
        """行次必须连续 1~25，不能跳号或重号。

        【为什么测这个？】
            行次是模板 B 列的编号，也是本表全部格位的推算依据。
            若某行行次写错，该行所有 13 个格子都会整体错位，
            数字落到别的项目上，而报表上每格都"有数字"，极难发现。
        """
        self.assertEqual([s.row_no for s in ROW_SPECS],
                         list(range(1, 26)),
                         "行次应为 1~25 连续排列")

    def test_excel_row_mapping(self):
        """行次 → Excel 行号：行次 1 在第 9 行，行次 25 在第 33 行。"""
        self.assertEqual(FIRST_DATA_ROW, 9)
        self.assertEqual(ROW_SPECS[0].excel_row, 9)
        self.assertEqual(ROW_SPECS[-1].excel_row, 33)
        self.assertEqual(TOTAL_ROW, 34)

    def test_keys_unique(self):
        """行的 key 不得重复（否则取数会分流到同一行）。"""
        keys = [s.key for s in ROW_SPECS]
        self.assertEqual(len(keys), len(set(keys)),
                         f"存在重复的行 key：{keys}")

    def test_codes_are_strings(self):
        """科目编码必须是字符串（写成数字可能导致匹配不上）。"""
        for spec in ROW_SPECS:
            for c in (*spec.codes, *spec.exclude_codes):
                self.assertIsInstance(
                    c, str, f"{spec.key} 的科目编码 {c!r} 不是字符串")

    def test_configured_rows_match_requirements(self):
        """当前已配置科目的行应与需求逐条对应，其余待提供。

        【为什么锁这条？】
            若有人误填了"猜测的科目"，报表会静默产出
            **看起来正常但口径不对**的数字 —— 而报表上每格都很正常。

            每补一行需求，就把它加进下面的 EXPECTED 里。
        """
        # {行 key: 期望科目} —— 随需求逐条补充
        expected = {
            "工资": ("510101",),        # 需求第 2 条
            "职工福利费": ("510102",),  # 需求第 3 条
            "办公费": ("510103",),      # 需求第 4 条
            "差旅费": ("510104",),      # 需求第 5 条
            "修理费": ("510106",),      # 需求第 6 条
            "机物料消耗": ("510107",),  # 需求第 7 条
            "社保及公积金": ("510108", "510130"),  # 需求第 8 条（两科目合计）
            # 需求第 9 条：商业保险费 → 暂不处理，不进本表
            "交际应酬费": ("510109",),  # 需求第 10 条
            "汽车费用": ("510110",),    # 需求第 11 条
            "低值易耗品摊销": ("510111",),  # 需求第 12 条
            "折旧费": ("510112",),      # 需求第 13 条
            "租赁费": ("510113",),      # 需求第 14 条
            "水电费": ("510114",),      # 需求第 15 条
            "通讯费": ("510115",),      # 需求第 16 条
            "劳动保护费": ("510125",),  # 需求第 17 条
            "工会经费": ("510129",),    # 需求第 18 条
            "职工教育费": ("510119",),  # 需求第 19 条
            "检验检测费用": ("510121",),  # 需求第 20 条
            # 需求第 21、22 条：无形资产摊销、安全费用 → 暂不处理
            "运输费": ("510105",),      # 需求第 23 条
            "其他": ("510199",),        # 需求第 24 条
            # 需求第 25 条：关税及报关费用 → 暂不处理
            "加工费": ("510124", "510128"),  # 需求第 26 条（两科目合计）
        }
        actual = {s.key: s.codes for s in ROW_SPECS if s.codes}
        self.assertEqual(actual, expected,
                         f"已配置行与需求不符，实际：{actual}")

        # 【至此 25 行全部有了归属，不再有"待配置"】
        #   21 行已配置 + 4 行暂不处理 = 25。
        #   若将来某行被误删、或 status 判断出错，这里会立刻发现。
        todo = sorted(s.key for s in ROW_SPECS if s.status == "待配置")
        self.assertEqual(todo, [],
                         f"需求已全部到位，不应再有待配置行，实际：{todo}")

        # 【注意排序】sorted() 按 Unicode 码点排，"关税"在"商业"之前。
        #   这里按实际排序写死，避免每次跑都要重新推一遍顺序。
        paused = sorted(s.key for s in ROW_SPECS if s.status == "暂不处理")
        self.assertEqual(
            paused,
            ["关税及报关费用", "商业保险费", "安全费用", "无形资产摊销"],
            f"暂不处理的应有四行，实际：{paused}")

        # 21 + 4 = 25，确保没有一行既不在已配置、也不在暂不处理里
        self.assertEqual(len(actual) + len(paused), len(ROW_SPECS),
                         "已配置 + 暂不处理 应恰好覆盖全部 25 行")

    def test_salary_row_codes(self):
        """「工资」行：科目 510101，落在第 9 行（需求第 2 条）。"""
        spec = next(s for s in ROW_SPECS if s.key == "工资")
        self.assertEqual(spec.codes, ("510101",),
                         f"工资行科目应为 ('510101',)，实际 {spec.codes}")
        self.assertEqual(spec.excel_row, 9,
                         f"工资行应落在第 9 行（C9），实际 {spec.excel_row}")

    def test_welfare_row_codes(self):
        """「职工福利费」行：科目 510102，落在第 10 行（需求第 4 条）。"""
        spec = next(s for s in ROW_SPECS if s.key == "职工福利费")
        self.assertEqual(spec.codes, ("510102",),
                         f"职工福利费行科目应为 ('510102',)，实际 {spec.codes}")
        self.assertEqual(spec.excel_row, 10,
                         f"职工福利费行应落在第 10 行（C10），"
                         f"实际 {spec.excel_row}")

    def test_social_insurance_row_codes(self):
        """「社保及公积金」行：510108 + 510130 两科目合计（需求第 8 条）。

        【为什么这条要单独测？】
            本行是**第一个由两个科目构成的行**。若程序只取第一个科目，
            C15 会只显示 510108 的金额 —— 比正确值少了公积金那一块，
            而报表上只是一个偏小的数，看不出漏了什么。

            【编码不相邻，特别注意】
              社保 510108 与公积金 510130 在账套里并不相邻，
              不能按"社保后面应该跟公积金"的顺序去猜。
        """
        spec = next(s for s in ROW_SPECS if s.key == "社保及公积金")
        self.assertEqual(spec.codes, ("510108", "510130"),
                         f"社保及公积金应为两科目合计，实际 {spec.codes}")
        self.assertEqual(spec.excel_row, 15,
                         f"社保及公积金应落在第 15 行（C15），"
                         f"实际 {spec.excel_row}")

    def test_social_insurance_sums_two_codes(self):
        """C15 = 510108 + 510130（两科目的数必须相加，不能只取一个）。"""
        rows = [
            ("510108", 1, 800.0),      # 社保
            ("510130", 1, 300.0),      # 公积金
        ]
        ctx = FakeCtx(rows, year=2026, month=1)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        self.assertTrue(approx(hook.get("C15"), 1100.0),
                        f"C15 应为 800+300=1100，实际 {hook.get('C15')}")

    def test_social_insurance_without_housing_fund(self):
        """公积金无记录时，C15 = 社保本身（不能空白）。"""
        ctx = FakeCtx([("510108", 1, 800.0)], year=2026, month=1)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        self.assertTrue(approx(hook.get("C15"), 800.0),
                        f"公积金无记录时应等于社保 800，实际 {hook.get('C15')}")

    def test_social_insurance_not_crossed_with_other_rows(self):
        """社保两科目不得串入其他行（尤其 510107 之后的科目）。"""
        rows = [
            ("510108", 1, 800.0),
            ("510130", 1, 300.0),
        ]
        ctx = FakeCtx(rows, year=2026, month=1)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        # 机物料消耗在第 14 行，应为空
        self.assertIsNone(hook.get("C14"),
                          f"C14 不应被写入，实际 {hook.get('C14')}")

    def test_commercial_insurance_is_pending(self):
        """商业保险费（C16）应登记为"暂不处理"，且不配科目（需求第 9 条）。

        【为什么锁这条？】
            需求原话是"商业保险这一行不明，暂不处理"。

            "暂不处理"与"待配置"是两种不同状态：
              · 暂不处理 → 已确认本阶段不做，等口径清晰再补；
              · 待配置   → 需求还没轮到这一行。

            若不显式区分，日后维护者会：
              · 把 pending 行当成"漏做"，反复去追这条已确认不做的需求；
              · 或把待配置行当成"已确认不做"，从此不再补 —— 报表静默缺项。
        """
        spec = next(s for s in ROW_SPECS if s.key == "商业保险费")
        self.assertEqual(spec.codes, (),
                         f"暂不处理的行不应配科目，实际 {spec.codes}")
        self.assertTrue(spec.pending,
                        "商业保险费应标记 pending=True")
        self.assertEqual(spec.status, "暂不处理")

    def test_commercial_insurance_in_pending_list(self):
        """商业保险费应登记在 PENDING_ROWS 中（供人工核对）。"""
        from reports.lk25_manufacturing_cost import PENDING_ROWS

        pending_cells = {cells for _, cells, _ in PENDING_ROWS}
        self.assertIn("C16:N16", pending_cells,
                      "商业保险费应在 PENDING_ROWS 中登记为 C16:N16")

    def test_intangible_and_safety_are_pending(self):
        """无形资产摊销（C28）、安全费用（C29）应登记为"暂不处理"（需求 21、22）。

        【为什么与商业保险费同样处理？】
            需求原话都是"这一行…不明，暂不处理"，属同一类：
            已确认本阶段不做，等口径清晰再补 —— 不是"漏做"。

            两者都必须：
              · codes 为空、pending=True（不取数、不写数）；
              · 登记在 PENDING_ROWS（供人工核对，防日后被当成遗漏）。
        """
        from reports.lk25_manufacturing_cost import PENDING_ROWS

        expected = {
            "无形资产摊销": (20, "C28:N28"),
            "安全费用": (21, "C29:N29"),
        }
        pending_cells = {cells for _, cells, _ in PENDING_ROWS}

        for key, (row_no, cells) in expected.items():
            with self.subTest(key=key):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.row_no, row_no,
                                 f"{key} 行次应为 {row_no}，实际 {spec.row_no}")
                self.assertEqual(spec.codes, (),
                                 f"{key} 暂不处理，不应配科目，实际 {spec.codes}")
                self.assertTrue(spec.pending, f"{key} 应标记 pending=True")
                self.assertEqual(spec.status, "暂不处理")
                self.assertIn(cells, pending_cells,
                              f"{key} 应在 PENDING_ROWS 中登记为 {cells}")

    def test_intangible_and_safety_cells_not_written(self):
        """即使有数据，C28/C29 十三列也不得被写入。"""
        ctx = FakeCtx([("510108", 1, 100.0)], year=2026, month=1)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        for row in (28, 29):
            for col in "CDEFGHIJKLMN":
                self.assertIsNone(
                    hook.get(f"{col}{row}"),
                    f"{col}{row} 已确认暂不处理，不应被写入，"
                    f"实际 {hook.get(f'{col}{row}')}")

    def test_commercial_insurance_cell_not_written(self):
        """即使有数据，C16 也不得被写入（确认"暂不处理"真的没写数）。"""
        ctx = FakeCtx([("510108", 1, 100.0)], year=2026, month=1)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        for col in "CDEFGHIJKLMN":
            self.assertIsNone(hook.get(f"{col}16"),
                              f"{col}16 已确认暂不处理，不应被写入，"
                              f"实际 {hook.get(f'{col}16')}")

    def test_requirement_10_to_17_rows(self):
        """需求 10~17：C17~C24 八行，各配一个科目。

            C17 交际应酬费      → 510109
            C18 汽车费用        → 510110
            C19 低值易耗品摊销  → 510111
            C20 折旧费          → 510112
            C21 租赁费          → 510113
            C22 水电费          → 510114
            C23 通讯费          → 510115
            C24 劳动保护费      → 510125

        【为什么八行一起测？】
            格位连续（17~24）、口径相同，复制粘贴时最容易漏改某一行
            的科目或行号。八行一起测，任一处错位都会立刻暴露。

        【特别注意 510125】
            510115（通讯费）之后并非 510116，而是 **510125**（劳动保护费）——
            账套里另设了一段编码。若有人按"116 之后应该是 117"
            的顺序去猜，会配出一个不存在的科目，该行永远空白。
        """
        expected = {
            "交际应酬费": (("510109",), 17),
            "汽车费用": (("510110",), 18),
            "低值易耗品摊销": (("510111",), 19),
            "折旧费": (("510112",), 20),
            "租赁费": (("510113",), 21),
            "水电费": (("510114",), 22),
            "通讯费": (("510115",), 23),
            "劳动保护费": (("510125",), 24),
        }
        for key, (codes, excel_row) in expected.items():
            with self.subTest(key=key):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.codes, codes,
                                 f"{key} 科目应为 {codes}，实际 {spec.codes}")
                self.assertEqual(spec.excel_row, excel_row,
                                 f"{key} 应落在第 {excel_row} 行，"
                                 f"实际 {spec.excel_row}")

    def test_requirement_25_26_rows(self):
        """需求 25~26：关税及报关费用（C31）暂不处理、加工费（C32）双科目合计。

            C31 关税及报关费用 → 暂不处理（口径不明）
            C32 加工费         → 510124 + 510128 合计

        【为什么把 C31 与 C32 放在一起测？】
            这两条需求是同时给的，且**性质完全相反**：
              · C31 是"不做"（pending，必须空白）；
              · C32 是"要做，且要两个科目相加"。
            放在一起，能验证程序没有把两行搞混 ——
            例如误把 510124/510128 写到 C31，或让 C31 也跟着取数。
        """
        from reports.lk25_manufacturing_cost import PENDING_ROWS

        # ---- C31 关税及报关费用：暂不处理 ----
        c31 = next(s for s in ROW_SPECS if s.key == "关税及报关费用")
        self.assertEqual(c31.excel_row, 31, f"应为第 31 行，实际 {c31.excel_row}")
        self.assertEqual(c31.codes, (), f"C31 暂不处理不应配科目，实际 {c31.codes}")
        self.assertTrue(c31.pending, "C31 应标记 pending=True")
        self.assertEqual(c31.status, "暂不处理")
        self.assertIn("C31:N31", {cells for _, cells, _ in PENDING_ROWS},
                      "C31 应登记在 PENDING_ROWS 中")

        # ---- C32 加工费：两科目合计 ----
        c32 = next(s for s in ROW_SPECS if s.key == "加工费")
        self.assertEqual(c32.excel_row, 32, f"应为第 32 行，实际 {c32.excel_row}")
        self.assertEqual(c32.codes, ("510124", "510128"),
                         f"加工费应为两科目合计，实际 {c32.codes}")
        self.assertFalse(c32.pending, "C32 是已配置行，不应标 pending")

    def test_processing_fee_sums_two_codes(self):
        """C32 = 510124 + 510128（必须相加，不能只取一个）。"""
        rows = [
            ("510124", 1, 124.0),
            ("510128", 1, 128.0),
        ]
        ctx = FakeCtx(rows, year=2026, month=1)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        self.assertTrue(approx(hook.get("C32"), 252.0),
                        f"C32 应为 124+128=252，实际 {hook.get('C32')}")

    def test_processing_fee_without_second_code(self):
        """510128 无记录时，C32 = 510124 本身（不能空白）。"""
        ctx = FakeCtx([("510124", 1, 124.0)], year=2026, month=1)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        self.assertTrue(approx(hook.get("C32"), 124.0),
                        f"应为 124，实际 {hook.get('C32')}")

    def test_processing_fee_not_written_to_tariff_row(self):
        """加工费的两科目不得串入 C31（关税行已确认暂不处理）。"""
        rows = [
            ("510124", 1, 124.0),
            ("510128", 1, 128.0),
        ]
        ctx = FakeCtx(rows, year=2026, month=1)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        for col in "CDEFGHIJKLMN":
            self.assertIsNone(hook.get(f"{col}31"),
                              f"{col}31 已确认暂不处理，不应被写入，"
                              f"实际 {hook.get(f'{col}31')}")

    def test_requirement_23_24_rows(self):
        """需求 23~24：运输费（C30）、其他（C33）。

            C30 运输费  → 510105
            C33 其他    → 510199

        【为什么这两行一起测？】
            两者格位**不连续**（30 与 33，中间隔着 31 关税、32 加工费）。
            若按"填满空行"的思路去补，很容易把 510105 写到 C31 或 C32。
            金额刻意与编码同值，一旦错位，断言会指出是哪个格子。

        【注意 C33 的位置】
            本表共 25 个数据行（Excel 第 9~33 行），C33 是**最后一个**
            费用项目行，紧随其后第 34 行是模板合计（公式，程序不写）。
        """
        expected = {
            "运输费": (("510105",), 30),
            "其他": (("510199",), 33),
        }
        for key, (codes, excel_row) in expected.items():
            with self.subTest(key=key):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.codes, codes,
                                 f"{key} 科目应为 {codes}，实际 {spec.codes}")
                self.assertEqual(spec.excel_row, excel_row,
                                 f"{key} 应落在第 {excel_row} 行，"
                                 f"实际 {spec.excel_row}")

    def test_requirement_23_24_write_to_own_columns(self):
        """运输费与「其他」各写各的列，且不串到相邻空白行或合计行。"""
        rows = [
            ("510105", 1, 105.0),
            ("510199", 1, 199.0),
        ]
        ctx = FakeCtx(rows, year=2026, month=1)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        self.assertTrue(approx(hook.get("C30"), 105.0), hook.get("C30"))
        self.assertTrue(approx(hook.get("C33"), 199.0), hook.get("C33"))

        # 中间两行（C31 关税、C32 加工费）仍未配置，必须空白
        for row in (31, 32):
            self.assertIsNone(hook.get(f"C{row}"),
                              f"C{row} 仍未配置，不应被写入，"
                              f"实际 {hook.get(f'C{row}')}")

        # 第 34 行合计由模板公式负责，程序不得写入
        self.assertIsNone(hook.get("C34"),
                          f"C34 是模板公式，不应被写入，"
                          f"实际 {hook.get('C34')}")

    def test_requirement_18_to_20_rows(self):
        """需求 18~20：C25~C27 三行，科目编码不连续且不按行序递增。

            C25 工会经费      → 510129
            C26 职工教育费    → 510119   ← 比上一行小
            C27 检验检测费用  → 510121

        【为什么特别标注"不按行序递增"？】
            本表科目编码与行序毫无对应关系：
              510125(劳动保护费) → 510129(工会经费) → 510119(职工教育费)
            编码忽大忽小。若有人按照"上一行 +N"的直觉去推算，
            会配出完全不存在的科目，该行永远空白 —— 而报表上看不出异常。

            所以这三行必须**逐个按需求给的编码填**，本条测试就是守住这点。
        """
        expected = {
            "工会经费": (("510129",), 25),
            "职工教育费": (("510119",), 26),
            "检验检测费用": (("510121",), 27),
        }
        for key, (codes, excel_row) in expected.items():
            with self.subTest(key=key):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.codes, codes,
                                 f"{key} 科目应为 {codes}，实际 {spec.codes}")
                self.assertEqual(spec.excel_row, excel_row,
                                 f"{key} 应落在第 {excel_row} 行，"
                                 f"实际 {spec.excel_row}")

    def test_requirement_18_to_20_write_to_own_columns(self):
        """三行各写各的列（金额刻意与编码同值，便于肉眼核对）。"""
        rows = [
            ("510129", 1, 129.0),
            ("510119", 1, 119.0),
            ("510121", 1, 121.0),
        ]
        ctx = FakeCtx(rows, year=2026, month=1)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        self.assertTrue(approx(hook.get("C25"), 129.0), hook.get("C25"))
        self.assertTrue(approx(hook.get("C26"), 119.0), hook.get("C26"))
        self.assertTrue(approx(hook.get("C27"), 121.0), hook.get("C27"))

    def test_eight_rows_write_to_own_columns(self):
        """八行各写各的列，互不串位（给不同金额跑一遍）。"""
        values = {
            "510109": 109.0, "510110": 110.0, "510111": 111.0,
            "510112": 112.0, "510113": 113.0, "510114": 114.0,
            "510115": 115.0, "510125": 125.0,
        }
        rows = [(c, 1, v) for c, v in values.items()]
        ctx = FakeCtx(rows, year=2026, month=1)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        code_order = ["510109", "510110", "510111", "510112",
                      "510113", "510114", "510115", "510125"]
        for i, code in enumerate(code_order):
            row = 17 + i
            self.assertTrue(
                approx(hook.get(f"C{row}"), values[code]),
                f"C{row}（{code}）应为 {values[code]}，"
                f"实际 {hook.get(f'C{row}')}")

    def test_office_travel_repair_material_rows(self):
        """办公费/差旅费/修理费/机物料消耗 四行（需求 4~7）。

            C11 办公费      → 510103
            C12 差旅费      → 510104
            C13 修理费      → 510106
            C14 机物料消耗  → 510107

        【为什么四行一起测？】
            这四行口径与工资行完全相同、格位连续（11~14 行），
            复制粘贴时最容易"漏改某一行的科目"：
              · 510104 写成 510103 → 差旅费显示办公费的数
              · 510107 写成 510108 → 机物料消耗空白

            【特别注意 510105 不在本表任何一行】
              科目编码从 510103 跳到 510106，中间空了一个 510105。
              这是账套里另设的明细科目，本表不使用。
              若有人"按顺序补"把 510105 也加进来，本条会失败 ——
              这正是想要的：**科目必须来自需求，不能靠编码连续去猜**。
        """
        expected = {
            "办公费": (("510103",), 11),
            "差旅费": (("510104",), 12),
            "修理费": (("510106",), 13),
            "机物料消耗": (("510107",), 14),
        }
        for key, (codes, excel_row) in expected.items():
            with self.subTest(key=key):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.codes, codes,
                                 f"{key} 科目应为 {codes}，实际 {spec.codes}")
                self.assertEqual(spec.excel_row, excel_row,
                                 f"{key} 应落在第 {excel_row} 行，"
                                 f"实际 {spec.excel_row}")

    def test_510105_used_by_transport_row_only(self):
        """510105 应且仅应属于「运输费」行（需求第 23 条）。

        【★ 这条断言曾经是「不得使用」，现已按需求反转 ★】

            需求第 4~7 条给出 510103/510104/510106/510107，中间跳过 510105。
            当时据此断言"510105 不在需求范围内、不得出现在任何一行" ——
            目的是防止有人按编码连续去猜、把需求外的科目配进某行。

            后来需求第 23 条明确：运输费 = 510105。
            所以断言反转为"必须出现，且只能出现在运输费行"。

            【两次判断依据的其实是同一条原则】
              · 需求未给编码时：程序不得擅自启用（防配错、报表静默出错数）；
              · 需求明确给出后：按给定编码配置。
            差别只在时点，不在原则。

            【为什么仍要锁"唯一性"？】
              若 510105 同时出现在两行，同一笔金额会被算两次，
              两行各显示一部分金额，而报表上完全看不出异常。
        """
        rows_with_510105 = [
            s.key for s in ROW_SPECS
            if "510105" in (*s.codes, *s.exclude_codes)
        ]
        self.assertEqual(rows_with_510105, ["运输费"],
                         f"510105 应只属于运输费行，实际出现在："
                         f"{rows_with_510105}")


class TestUnconfiguredBehaviour(unittest.TestCase):
    """科目未配置时的行为：不查库、不写数，但页眉与期间照常。"""

    def test_unconfigured_rows_not_written(self):
        """未配置科目的行不应写入任何金额（不写 0）。

        【为什么改成"未配置行"而不是"所有行"？】
            需求第 2 条已为「工资」行配置了 510101，
            它现在是"已配置"行，不该再要求它空白。
            本条只针对仍待配置的其余 24 行。
        """
        ctx = FakeCtx([], month=5)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        for spec in ROW_SPECS:
            if spec.codes:
                continue          # 已配置的行不在本条范围内
            for col in COL_OF_MONTH.values():
                cell = f"{col}{spec.excel_row}"
                self.assertIsNone(
                    hook.get(cell),
                    f"{spec.key} 的 {cell} 不应被写入，实际 {hook.get(cell)}")

    def test_no_database_query_when_nothing_configured(self):
        """若所有行科目都为空，则不应发起任何查询。

        【测试手法】临时把 ROW_SPECS 全部清空模拟"科目尚未提供"的状态，
        验证 fetch 的短路逻辑仍然有效（这是将来新增行时可能触碰的路径）。
        """
        original = _with_codes({})       # 全部置空
        self.addCleanup(setattr, mod, "ROW_SPECS", original)

        ctx = FakeCtx([], month=5)
        LK25Report().fetch(ctx)
        self.assertEqual(ctx.fetch_calls, [],
                         f"科目全为空时不应查库，实际发起了 "
                         f"{len(ctx.fetch_calls)} 次查询")

    def test_period_still_written(self):
        """即使没有任何数据，会计期间仍应写入 C4。

        年月是报表的"身份信息"，与有没有取到数无关。
        若因为"无数据"就不写，打开文件根本分不清这是哪个月的报表。
        """
        ctx = FakeCtx([], month=5)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)
        self.assertEqual(hook.get(PERIOD_CELL), "2026年5月")


class TestPeriodCell(unittest.TestCase):
    """需求第 1 条：C4 填"界面所选的会计期间"。

    【本组用例的设计意图】
        单测一个 "2026年5月" 是**不足以**证明"年月随选择变化"的 ——
        代码里写死一个常量同样能通过那一条。

        所以这里覆盖多种情形：
          · 不同年份（2026 / 2025 / 2027）
          · 不同月份（1 月 / 5 月 / 12 月）
          · 有数据 / 无数据 / 取数失败
        只要程序在某处把年月写死，或把它与"是否有数据"绑在一起，
        这些用例必然失败。
    """

    def _write(self, year, month, rows=None, fail=False):
        ctx = FakeCtx(rows or [], year=year, month=month, fail=fail)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)
        return hook

    def test_period_follows_selected_month(self):
        """年月必须与界面所选完全一致（多组组合逐一验证）。"""
        cases = [
            (2026, 1, "2026年1月"),
            (2026, 5, "2026年5月"),
            (2026, 12, "2026年12月"),
            (2025, 11, "2025年11月"),
            (2027, 3, "2027年3月"),
        ]
        for year, month, expected in cases:
            with self.subTest(year=year, month=month):
                hook = self._write(year, month)
                self.assertEqual(
                    hook.get(PERIOD_CELL), expected,
                    f"界面选 {year}年{month}月，C4 应为 {expected!r}，"
                    f"实际 {hook.get(PERIOD_CELL)!r}")

    def test_period_not_hardcoded(self):
        """不同月份生成的 C4 必须不同 —— 防止年月被写成常量。

        【为什么单独测"不同"？】
            若代码写成 C4="2026年5月" 常量，5 月用例仍会通过。
            本条要求"选 5 月"与"选 8 月"的结果不一样，
            常量写法无法同时满足。
        """
        may = self._write(2026, 5).get(PERIOD_CELL)
        aug = self._write(2026, 8).get(PERIOD_CELL)
        self.assertNotEqual(may, aug,
                            f"不同月份生成出相同的年月 {may!r} —— "
                            f"年月疑似被写成了常量")

    def test_period_written_without_data(self):
        """科目未配置（零数据）时，C4 仍必须是所选年月。"""
        for month in (1, 6, 12):
            with self.subTest(month=month):
                hook = self._write(2026, month)
                self.assertEqual(hook.get(PERIOD_CELL), f"2026年{month}月")

    def test_period_written_when_fetch_fails(self):
        """取数抛异常时，C4 仍必须是所选年月（不受取数失败影响）。"""
        hook = self._write(2026, 9, fail=True)
        self.assertEqual(hook.get(PERIOD_CELL), "2026年9月",
                         "取数失败不应导致会计期间为空")

    def test_period_written_with_data(self):
        """有数据时，C4 也应写所选年月（不被数据流程覆盖）。"""
        original = _with_codes({"工资": ("510101",)})
        self.addCleanup(setattr, mod, "ROW_SPECS", original)

        hook = self._write(2026, 7, rows=[("510101", 7, 700.0)])
        self.assertEqual(hook.get(PERIOD_CELL), "2026年7月")
        # 顺带确认数据也进了正确月份列。
        # 【注意列位】7 月是 I 列（C=1月…H=6月，I=7月，N=12月），
        #   不要凭"第 7 列就是 J"的直觉去数 —— J 其实是 8 月。
        self.assertTrue(approx(hook.get("I9"), 700.0), hook.get("I9"))

    def test_period_written_as_text(self):
        """C4 必须用 set_text 写，不能用 set。

        【为什么？】C4 所在列是金额格式。用 set() 写 "2026年5月"，
        Excel 会把它识别成日期并存成序列号 46143，
        再显示成 "46,143.00" —— 看起来像一笔钱（LK01 踩过的坑）。
        """
        hook = self._write(2026, 5)
        kinds = [e[0] for e in hook.events if e[1] == PERIOD_CELL]
        self.assertIn("set_text", kinds,
                      f"C4 未用 set_text 写入，实际事件：{hook.events}")
        self.assertNotIn("set", kinds,
                         "C4 不应使用普通 set 写入（会被识别成日期）")

    def test_period_cell_is_c4(self):
        """需求指定的位置是 C4，不能被改成别处。"""
        self.assertEqual(PERIOD_CELL, "C4")


class TestRowValues(unittest.TestCase):
    """科目配好后的取数与写入。"""

    def setUp(self):
        original = _with_codes({
            "工资": ("510101", "510102"),   # 本行由两个科目构成
            "折旧费": ("510110",),
        })
        self.addCleanup(setattr, mod, "ROW_SPECS", original)

    def _run(self, rows, month=5):
        ctx = FakeCtx(rows, month=month)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)
        return hook, ctx, rep

    def test_month_columns(self):
        """各月发生额应落到对应月份列。"""
        rows = [
            ("510101", 1, 100.0),
            ("510101", 2, 200.0),
            ("510101", 3, 300.0),
            ("510101", 4, 400.0),
            ("510101", 5, 500.0),
        ]
        hook, _, _ = self._run(rows, month=5)

        # 工资在第 9 行；C=1月 … G=5月
        self.assertTrue(approx(hook.get("C9"), 100.0), hook.get("C9"))
        self.assertTrue(approx(hook.get("D9"), 200.0), hook.get("D9"))
        self.assertTrue(approx(hook.get("E9"), 300.0), hook.get("E9"))
        self.assertTrue(approx(hook.get("F9"), 400.0), hook.get("F9"))
        self.assertTrue(approx(hook.get("G9"), 500.0), hook.get("G9"))

    def test_multi_codes_summed(self):
        """同一行的多个科目应相加。

        【为什么重要？】"社保及公积金"这类项目在实务中一定由多个
        明细科目构成。若程序只取第一个科目，报表会显示一个偏小的数，
        而看不出少了什么。
        """
        rows = [
            ("510101", 5, 100.0),
            ("510102", 5, 25.0),
        ]
        hook, _, _ = self._run(rows, month=5)
        self.assertTrue(approx(hook.get("G9"), 125.0),
                        f"工资应为 100+25=125，实际 {hook.get('G9')}")

    def test_rows_not_crossed(self):
        """不同行各取各的科目，互不串用。"""
        rows = [
            ("510101", 5, 100.0),    # 工资 → 第 9 行
            ("510110", 5, 900.0),    # 折旧费 → 第 20 行
        ]
        hook, _, _ = self._run(rows, month=5)
        self.assertTrue(approx(hook.get("G9"), 100.0), hook.get("G9"))
        self.assertTrue(approx(hook.get("G20"), 900.0), hook.get("G20"))

    def test_future_months_never_written(self):
        """所选期间之后的月份绝不写入。

        【为什么必须守住？】
            "2026年5月"的表里若混入 6 月的数，这张表就不再是 5 月表了，
            而看表的人无从分辨哪几列属于本期。
        """
        rows = [
            ("510101", 5, 500.0),
            ("510101", 6, 600.0),    # 超出所选期间，必须被丢弃
            ("510101", 12, 1200.0),
        ]
        hook, _, _ = self._run(rows, month=5)
        self.assertTrue(approx(hook.get("G9"), 500.0))
        self.assertIsNone(hook.get("H9"), f"6月不应写入，实际 {hook.get('H9')}")
        self.assertIsNone(hook.get("N9"), f"12月不应写入，实际 {hook.get('N9')}")

    def test_missing_month_stays_blank(self):
        """某月无发生额时该月格保持空白（不写 0）。"""
        rows = [("510101", 5, 500.0)]     # 只有 5 月
        hook, _, _ = self._run(rows, month=5)
        for cell in ("C9", "D9", "E9", "F9"):
            self.assertIsNone(hook.get(cell),
                              f"{cell} 无数据应保持空白，实际 {hook.get(cell)}")

    def test_year_total_not_written(self):
        """O 列（本年累计）不得由程序写入 —— 它是模板公式 =SUM(C:N)。

        【为什么这条最重要？】
            模板 O9 = '=SUM(C9:N9)'（已用 Excel COM 核实）。
            程序若把累计值写进 O 列，会**覆盖公式**：
            本月看着对，但公式没了，下月手工改明细时累计不会跟着变，
            而且报表上数字仍然"看着对"，极难察觉。
        """
        self.assertFalse(WRITE_YEAR_TOTAL,
                         "默认不应由程序写 O 列（模板已带公式）")
        rows = [("510101", 5, 500.0)]
        hook, _, _ = self._run(rows, month=5)
        self.assertIsNone(hook.get(f"{COL_YEAR_TOTAL}9"),
                          f"O9 不应被写入，实际 {hook.get(f'{COL_YEAR_TOTAL}9')}")

    def test_total_row_not_written(self):
        """第 34 行合计不得由程序写入 —— 它是模板公式 =SUM(C9:C33)。"""
        self.assertFalse(WRITE_TOTAL_ROW,
                         "默认不应由程序写合计行（模板已带公式）")
        rows = [
            ("510101", 5, 500.0),
            ("510110", 5, 900.0),
        ]
        hook, _, _ = self._run(rows, month=5)
        for col in COL_OF_MONTH.values():
            cell = f"{col}{TOTAL_ROW}"
            self.assertIsNone(hook.get(cell),
                              f"{cell} 不应被写入，实际 {hook.get(cell)}")

    def test_fetch_failure_not_crash(self):
        """取数抛异常时不崩溃，且不写任何金额。"""
        ctx = FakeCtx([], month=5, fail=True)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)
        self.assertIsNone(hook.get("C9"))
        self.assertEqual(hook.get(PERIOD_CELL), "2026年5月")

    def test_single_query_for_all_rows(self):
        """所有行的科目应合并成一次查询（去重）。"""
        _, ctx, _ = self._run([], month=5)
        self.assertEqual(len(ctx.fetch_calls), 1,
                         f"应只发一次查询，实际 {len(ctx.fetch_calls)} 次")
        params = ctx.fetch_calls[0]["params"]
        codes = [p for p in params if isinstance(p, str)]
        self.assertEqual(len(codes), len(set(codes)),
                         f"科目参数有重复：{codes}")


class TestRequirement2(unittest.TestCase):
    """需求第 2 条：C9 一月工资的取数口径。

        ① iYPeriod = 所选年月的年月整数（2026年1月 → 202601）
        ② ccode = 510101
        ③ cdigest 必须包含「结转制造费用」
        ④ 取 mc（贷方金额）

    【为什么这组测试不能省？】
        这四条件任一写错，SQL 都仍能正常执行、返回一个数字：
          · 取 md 而非 mc  → 拿到的是费用发生额，金额偏大
          · 漏摘要筛选      → 把冲销/调账的贷方也混进来，金额偏大
          · ccode 写错      → 取到别的费用项目
        报表上只会显示一个"看起来正常"的数，没有任何异常提示。
        所以必须把口径逐条锁死在 SQL 文本上。
    """

    def _run(self, rows, year=2026, month=1):
        ctx = FakeCtx(rows, year=year, month=month)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)
        return hook, ctx

    def test_sql_uses_credit_side(self):
        """SQL 必须 SUM(mc)，不能是 SUM(md)。

        【为什么直接查 SQL 文本？】
            假数据里只有一个 amount 值，报表层看不出方向。
            只有检查 SQL 才能真正确认取的是哪一列。
        """
        _, ctx = self._run([])
        sql = ctx.fetch_calls[0]["sql"]
        flat = sql.replace(" ", "")
        self.assertIn("SUM(mc)", flat, f"应取贷方 mc，实际 SQL：{sql}")
        self.assertNotIn("SUM(md)", flat, f"不应取借方 md，实际 SQL：{sql}")

    def test_sql_filters_digest(self):
        """SQL 必须带「结转制造费用」摘要筛选。"""
        _, ctx = self._run([])
        sql = ctx.fetch_calls[0]["sql"]
        self.assertIn("cdigest LIKE", sql, f"缺少摘要筛选：{sql}")
        self.assertIn("%结转制造费用%", ctx.fetch_calls[0]["params"],
                      f"摘要参数不正确：{ctx.fetch_calls[0]['params']}")

    def test_sql_period_matches_selected_month(self):
        """iYPeriod 区间必须落在所选年月（选 2026年1月 → 202601~202601）。"""
        _, ctx = self._run([], year=2026, month=1)
        params = ctx.fetch_calls[0]["params"]
        self.assertEqual(params[0], 2026, f"年度参数应为 2026，实际 {params}")
        self.assertEqual(params[1], 202601, f"起始应为 202601，实际 {params}")
        self.assertEqual(params[2], 202601, f"截止应为 202601，实际 {params}")

    def test_sql_period_for_other_month(self):
        """换一个月，区间上限随之变化（2026年3月 → 202601~202603）。"""
        _, ctx = self._run([], year=2026, month=3)
        params = ctx.fetch_calls[0]["params"]
        self.assertEqual(params[1], 202601, f"起始应固定为 202601，实际 {params}")
        self.assertEqual(params[2], 202603, f"截止应为 202603，实际 {params}")

    def test_sql_filters_salary_code(self):
        """SQL 的科目参数必须含 510101。"""
        _, ctx = self._run([], year=2026, month=1)
        self.assertIn("510101", ctx.fetch_calls[0]["params"],
                      f"科目参数应含 510101，实际 {ctx.fetch_calls[0]['params']}")

    def test_c9_filled_from_query(self):
        """C9 取查询结果；选 2026年1月时 1000.00 → C9=1000.00。"""
        hook, _ = self._run([("510101", 1, 1000.0)], year=2026, month=1)
        self.assertTrue(approx(hook.get("C9"), 1000.0),
                        f"C9 应为 1000.00，实际 {hook.get('C9')}")
        self.assertIsNone(hook.get("D9"), "2 月不应被写入")

    def test_salary_not_written_to_other_rows(self):
        """510101 的数据只能进第 9 行，不能串到其他费用项目。"""
        hook, _ = self._run([("510101", 1, 1000.0)], year=2026, month=1)
        for row in (10, 20, 33):
            self.assertIsNone(hook.get(f"C{row}"),
                              f"C{row} 不应被写入，实际 {hook.get(f'C{row}')}")

    def test_no_data_no_write(self):
        """该科目当月无结转记录时，C9 保持空白（不写 0）。

        【为什么不写 0？】写 0 会让财务以为"本月工资结转额为 0"，
        而实际可能是"忘记做结转凭证"。留空 + 日志才能让人去查。
        """
        hook, _ = self._run([], year=2026, month=1)
        self.assertIsNone(hook.get("C9"),
                          f"C9 不应被写入，实际 {hook.get('C9')}")


class TestRequirement3(unittest.TestCase):
    """需求第 3 条：所选月份之前的各月要一并填列（累计填列）。

        选 2026年1月 → 只填 C9
        选 2026年3月 → 同时填 C9、D9、E9
        选 2026年8月 → 填 C9~J9

    【为什么这组测试必须用真实配置（510101）而不是临时假科目？】
        本条是"端到端行为"：从界面年月 → SQL 区间 → 逐月落列。
        用真实配置才能验证"工资行真的会这样填"。
        临时假科目只能验证框架逻辑，验不到实际报表的呈现。

    【本条最关键的两个断言方向相反，缺一不可】
        ① 该填的要填：选 3 月时 C9/D9/E9 三列都有数据；
        ② 不该填的绝不能填：F9~N9（4~12 月）必须全部空白。
        只测 ① 会漏掉"把未来月份也填进来"的错误；
        只测 ② 会漏掉"只填了所选那一列"的错误（正是本条要修的）。
    """

    def setUp(self):
        # 1~12 月各给一个可区分的金额，任何落列错位都能立刻看出来：
        #   1月=100、2月=200、3月=300 … 12月=1200
        self.rows = [("510101", m, m * 100.0) for m in range(1, 13)]

    def _run(self, month):
        ctx = FakeCtx(self.rows, year=2026, month=month)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)
        return hook, ctx

    def test_january_fills_only_one_column(self):
        """选 1 月：只填 C9，D9 及以后全部空白。"""
        hook, _ = self._run(1)
        self.assertTrue(approx(hook.get("C9"), 100.0),
                        f"C9 应为 100，实际 {hook.get('C9')}")
        for col in "DEFGHIJKLMN":
            self.assertIsNone(hook.get(f"{col}9"),
                              f"{col}9 在 1 月报表中不应被写入，"
                              f"实际 {hook.get(f'{col}9')}")

    def test_march_fills_three_columns(self):
        """选 3 月：C9=100、D9=200、E9=300 —— 前三个月一并填列。"""
        hook, _ = self._run(3)
        self.assertTrue(approx(hook.get("C9"), 100.0),
                        f"C9 应为 100，实际 {hook.get('C9')}")
        self.assertTrue(approx(hook.get("D9"), 200.0),
                        f"D9 应为 200，实际 {hook.get('D9')}")
        self.assertTrue(approx(hook.get("E9"), 300.0),
                        f"E9 应为 300，实际 {hook.get('E9')}")

    def test_march_does_not_fill_future_months(self):
        """选 3 月：F9~N9（4~12 月）必须全部空白。"""
        hook, _ = self._run(3)
        for col in "FGHIJKLMN":
            self.assertIsNone(hook.get(f"{col}9"),
                              f"{col}9 属 4 月及以后，3 月报表中不应写入，"
                              f"实际 {hook.get(f'{col}9')}")

    def test_august_fills_eight_columns(self):
        """选 8 月：C9~J9 全部填列（1~8 月共 8 列），K9 起空白。"""
        hook, _ = self._run(8)
        for i, col in enumerate("CDEFGHIJ", start=1):
            self.assertTrue(
                approx(hook.get(f"{col}9"), i * 100.0),
                f"选 8 月时 {col}9（第{i}月）应为 {i * 100}，"
                f"实际 {hook.get(f'{col}9')}")
        for col in "KLMN":
            self.assertIsNone(hook.get(f"{col}9"),
                              f"{col}9 属 9 月及以后，8 月报表中不应写入")

    def test_december_fills_all_twelve_columns(self):
        """选 12 月：C9~N9 十二列全填。"""
        hook, _ = self._run(12)
        for i, col in enumerate("CDEFGHIJKLMN", start=1):
            self.assertTrue(
                approx(hook.get(f"{col}9"), i * 100.0),
                f"选 12 月时 {col}9（第{i}月）应为 {i * 100}，"
                f"实际 {hook.get(f'{col}9')}")

    def test_sql_range_covers_one_to_selected(self):
        """SQL 的 iYPeriod 区间必须是 202601~所选月份。"""
        for month in (1, 3, 8, 12):
            with self.subTest(month=month):
                _, ctx = self._run(month)
                params = ctx.fetch_calls[0]["params"]
                self.assertEqual(
                    params[1], 202601,
                    f"选 {month} 月时区间起始应为 202601，实际 {params}")
                self.assertEqual(
                    params[2], int(f"2026{month:02d}"),
                    f"选 {month} 月时区间截止应为 2026{month:02d}，"
                    f"实际 {params}")

    def test_single_query_regardless_of_month(self):
        """无论选几个月，都只发一次查询（用区间一次取回，不逐月查）。

        【为什么值得测？】
            若改成"逐月查 N 次"，选 12 月就要 12 次数据库往返，
            且各月数据来自不同时刻 —— 一旦某月被补录，两次生成的结果
            可能互相矛盾。用区间一次取回可同时避免这两个问题。
        """
        for month in (1, 6, 12):
            with self.subTest(month=month):
                _, ctx = self._run(month)
                self.assertEqual(
                    len(ctx.fetch_calls), 1,
                    f"选 {month} 月时应只发 1 次查询，"
                    f"实际 {len(ctx.fetch_calls)} 次")

    def test_gap_month_stays_blank_but_others_filled(self):
        """中间某月缺数据时：该月空白，其余月份照常填。

        【场景】选 5 月，但 3 月没有结转制造费用的凭证
                （例如漏做结转、或该月确实无生产）。
        【期望】C9/D9/F9/G9 有数，E9 空白 —— 不能因为缺一个月
                就把整行都留空，也不能给缺的那月补 0。
        """
        rows = [("510101", m, m * 100.0) for m in (1, 2, 4, 5)]  # 缺 3 月
        ctx = FakeCtx(rows, year=2026, month=5)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        self.assertTrue(approx(hook.get("C9"), 100.0), hook.get("C9"))
        self.assertTrue(approx(hook.get("D9"), 200.0), hook.get("D9"))
        self.assertIsNone(hook.get("E9"),
                          f"3 月无数据应保持空白，实际 {hook.get('E9')}")
        self.assertTrue(approx(hook.get("F9"), 400.0), hook.get("F9"))
        self.assertTrue(approx(hook.get("G9"), 500.0), hook.get("G9"))

    def test_multi_rows_all_accumulate(self):
        """累计填列对每一行都生效，不只是工资行。

        【造数】临时给「折旧费」配科目，验证两行都能填 1~3 月。
        """
        original = _with_codes({
            "工资": ("510101",),
            "折旧费": ("510110",),
        })
        self.addCleanup(setattr, mod, "ROW_SPECS", original)

        rows = [
            ("510101", 1, 100.0), ("510101", 2, 200.0), ("510101", 3, 300.0),
            ("510110", 1, 10.0), ("510110", 2, 20.0), ("510110", 3, 30.0),
        ]
        ctx = FakeCtx(rows, year=2026, month=3)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        # 工资在第 9 行，折旧费在第 20 行
        for col, salary, dep in (("C", 100.0, 10.0),
                                 ("D", 200.0, 20.0),
                                 ("E", 300.0, 30.0)):
            self.assertTrue(approx(hook.get(f"{col}9"), salary),
                            f"{col}9 应为 {salary}，实际 {hook.get(f'{col}9')}")
            self.assertTrue(approx(hook.get(f"{col}20"), dep),
                            f"{col}20 应为 {dep}，实际 {hook.get(f'{col}20')}")


class TestRequirement4(unittest.TestCase):
    """需求第 4 条：C10 职工福利费（510102）口径同 C9 工资（510101）。

        各月对应 GL_accvouch 的 iperiod，摘要含「结转制造费用」，取 mc；
        **仅科目代码不同**。

    【本组测试的重点不是"再验一遍工资行的逻辑"，而是"两行并行不串"】
        两行口径完全相同，意味着它们会被 fetch() 合并进**同一次查询**
        （ccode IN ('510101','510102')），返回结果再按科目分流到两行。
        这个分流环节一旦出错，两行的数就会互换或叠加：

          · 分流写成单值字典 → 后写的行覆盖先写的，某行空白
          · 科目→行的映射写错 → 工资的数进福利费行，反之亦然

        而这类错误在报表上**每格都有数字**，看不出异常，
        只能靠"给两行不同的金额、逐格比对"才能发现。
    """

    # 两行各月金额刻意错开，任一处串位都会让断言失败
    SALARY = {1: 100.0, 2: 200.0, 3: 300.0}
    WELFARE = {1: 11.0, 2: 22.0, 3: 33.0}

    def _run(self, months=(1, 2, 3), up_to=3):
        rows = []
        for m in months:
            rows.append(("510101", m, self.SALARY[m]))
            rows.append(("510102", m, self.WELFARE[m]))

        ctx = FakeCtx(rows, year=2026, month=up_to)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)
        return hook, ctx

    def test_two_rows_accumulate_independently(self):
        """两行各月数值互不串用（选 3 月，C/D/E 三列）。"""
        hook, _ = self._run(up_to=3)
        for col, m in (("C", 1), ("D", 2), ("E", 3)):
            self.assertTrue(
                approx(hook.get(f"{col}9"), self.SALARY[m]),
                f"{col}9（工资 {m}月）应为 {self.SALARY[m]}，"
                f"实际 {hook.get(f'{col}9')}")
            self.assertTrue(
                approx(hook.get(f"{col}10"), self.WELFARE[m]),
                f"{col}10（福利费 {m}月）应为 {self.WELFARE[m]}，"
                f"实际 {hook.get(f'{col}10')}")

    def test_two_rows_not_summed_together(self):
        """两行不能被合并成一个数（防止"都累加到同一行"）。

        【典型错误】若科目→行的映射把 510101 与 510102 都指到同一行，
        该行会变成 111/222/333，而另一行空白 —— 报表上仍然"有数字"。
        """
        hook, _ = self._run(up_to=3)
        self.assertFalse(approx(hook.get("C9"), 111.0),
                         "C9 不应是两行之和 111（说明科目被合并到同一行）")
        self.assertFalse(approx(hook.get("C10"), 111.0),
                         "C10 不应是两行之和 111")

    def test_welfare_follows_selected_month_range(self):
        """累计填列规则对福利费行同样生效（选 2 月只填 C/D）。"""
        hook, _ = self._run(up_to=2)
        self.assertTrue(approx(hook.get("C10"), 11.0), hook.get("C10"))
        self.assertTrue(approx(hook.get("D10"), 22.0), hook.get("D10"))
        self.assertIsNone(hook.get("E10"),
                          f"选 2 月时 E10 不应被写入，实际 {hook.get('E10')}")

    def test_both_rows_in_single_query(self):
        """两行科目应合并进同一次查询（而非各查一次）。"""
        _, ctx = self._run()
        self.assertEqual(len(ctx.fetch_calls), 1,
                         f"两行应合并为 1 次查询，实际 {len(ctx.fetch_calls)} 次")
        params = ctx.fetch_calls[0]["params"]
        self.assertIn("510101", params)
        self.assertIn("510102", params)

    def test_welfare_only_not_written_to_salary(self):
        """只有福利费有数据时，工资行必须保持空白（不借用别的行的数）。"""
        ctx = FakeCtx([("510102", 1, 55.0)], year=2026, month=1)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        self.assertTrue(approx(hook.get("C10"), 55.0),
                        f"C10 应为 55，实际 {hook.get('C10')}")
        self.assertIsNone(hook.get("C9"),
                          f"工资行无数据应保持空白，实际 {hook.get('C9')}")

    def test_welfare_not_written_to_other_rows(self):
        """福利费的数据只能进第 10 行，不得串到其他未配置行。"""
        hook, _ = self._run(up_to=3)
        for row in (11, 20, 33):
            self.assertIsNone(hook.get(f"C{row}"),
                              f"C{row} 属未配置行，不应被写入，"
                              f"实际 {hook.get(f'C{row}')}")


class TestExcludeCodes(unittest.TestCase):
    """减项科目（exclude_codes）的行为。"""

    def setUp(self):
        patched = []
        for spec in ROW_SPECS:
            if spec.key == "其他":
                patched.append(dataclasses.replace(
                    spec, codes=("510199",), exclude_codes=("510198",)))
            else:
                patched.append(dataclasses.replace(spec, codes=()))
        original = mod.ROW_SPECS
        mod.ROW_SPECS = tuple(patched)
        self.addCleanup(setattr, mod, "ROW_SPECS", original)

    def _run(self, rows, month=5):
        ctx = FakeCtx(rows, month=month)
        hook = FakeHook()
        rep = LK25Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)
        return hook, ctx, rep

    def test_exclude_subtracted(self):
        """最终值 = sum(codes) - sum(exclude_codes)。"""
        rows = [
            ("510199", 5, 1000.0),
            ("510198", 5, 300.0),
        ]
        hook, _, _ = self._run(rows, month=5)
        # 其他在第 33 行；5 月 = G 列
        self.assertTrue(approx(hook.get("G33"), 700.0),
                        f"其他应为 1000-300=700，实际 {hook.get('G33')}")

    def test_only_exclude_not_written(self):
        """只有减项科目有记录时不写负数。

        【为什么不写？】照算会得到 -300 这种不可能的金额。
        宁可留空让人查，也不写一个方向错误的数。
        """
        hook, _, _ = self._run([("510198", 5, 300.0)], month=5)
        self.assertIsNone(hook.get("G33"),
                          f"不应写入负数，实际 {hook.get('G33')}")


class TestConfigConsistency(unittest.TestCase):
    """结构校验：防止配置与 cell_map / 模板脱节。"""

    def test_cell_map_covers_all_cells(self):
        """ROW_SPECS 的每一格（含 O 列）都应登记在 cell_map 中。

        【为什么重要？】cell_map 决定父类 verify() 读回哪些格做勾稽核对。
        漏登记 → 该行不进核对清单 → 日后"核对通过"可能是假象。
        """
        cm_values = set(LK25Report.cell_map.values())
        for spec in ROW_SPECS:
            for col in COL_OF_MONTH.values():
                cell = f"{col}{spec.excel_row}"
                self.assertIn(cell, cm_values,
                              f"{spec.key} 的 {cell} 未登记 cell_map")
            total_cell = f"{COL_YEAR_TOTAL}{spec.excel_row}"
            self.assertIn(total_cell, cm_values,
                          f"{spec.key} 的 {total_cell} 未登记 cell_map")

    def test_no_duplicate_cells(self):
        """不同行不得占用同一格子（否则后写的会覆盖先写的）。"""
        seen: dict[str, str] = {}
        for spec in ROW_SPECS:
            for col in COL_OF_MONTH.values():
                cell = f"{col}{spec.excel_row}"
                self.assertNotIn(
                    cell, seen,
                    f"格子 {cell} 被 {seen.get(cell)} 和 {spec.key} 同时使用")
                seen[cell] = spec.key

    def test_annual_total_does_not_cover_month_columns(self):
        """O 列不得与任何月列重合（否则累计与某个月抢同一格）。"""
        month_cols = set(COL_OF_MONTH.values())
        self.assertNotIn(COL_YEAR_TOTAL, month_cols,
                         f"{COL_YEAR_TOTAL} 与月列冲突：{sorted(month_cols)}")

    def test_cells_in_template_have_labels(self):
        """配置里的每个格子，其所在行在模板 A 列应有项目名。

        【为什么重要？】
            这条能抓住"地址写错行"——例如本该写 C20 却写成 C21，
            而 C21 那行是另一个费用项目，一看就露馅。
            若模板不存在则自动跳过。
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
        sh = wb.sheet_by_name("LK25(B29)")

        for spec in ROW_SPECS:
            # A 列（0-based 索引 0）是项目名
            val = str(sh.cell_value(spec.excel_row - 1, 0)).strip()
            self.assertTrue(
                val,
                f"{spec.key} 位于第 {spec.excel_row} 行，"
                f"但模板 A 列是空的 —— 可能写错了行")

    def test_template_has_formula_in_total_column(self):
        """模板 O 列与合计行应带 =SUM(...) 公式（程序不写的前提）。

        【为什么测这个？】
            本表"不写 O 列与合计行"的整个设计，前提是**模板确实有公式**。
            若某天模板被换成无公式版本，报表的累计列会一片空白，
            而程序不会报错（因为它按设计就没写）。本条会在那时报警。

        【★ 为什么整个 COM 交互都包进 try、失败就 skip？★】
            本测试依赖本机真实 Excel（win32com），而可用性取决于很多
            与报表逻辑无关的环境因素：
              · 没装 pywin32 / 没装 Excel；
              · 类型库缓存损坏；
              · 已有 Excel 进程占用 COM 单例，导致
                `app.Visible = ...` 抛 AttributeError
                （曾真实发生："Property 'Excel.Application.Visible'
                 can not be set."）。

            这些都不是"报表写错了"。若让它们判为 FAILED/ERROR，
            会造成持续的假警报 —— 维护者很快会忽略这条测试，
            届时**模板真被改坏也无人察觉**。

            所以：能连上 Excel 就严格校验公式；连不上就明确 skip。
            口径本身的校验（WRITE_YEAR_TOTAL=False 等）由其它用例保证，
            不依赖本环境。
        """
        from config.settings import DEFAULT_TEMPLATE, PROJECT_ROOT

        tpl = PROJECT_ROOT / DEFAULT_TEMPLATE
        if not tpl.exists():
            self.skipTest(f"模板不存在：{tpl}")

        app = None
        wb = None
        try:
            from core.excel_engine import _dispatch_excel
            app = _dispatch_excel()
            # 这两行在部分环境会抛 AttributeError（见上方说明），
            # 与报表逻辑无关，因此纳入同一个 try。
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(str(tpl), ReadOnly=True, UpdateLinks=0)
            ws = wb.Sheets("LK25(B29)")
            formulas = {
                "O9": ws.Cells(9, 15).Formula,
                "O33": ws.Cells(33, 15).Formula,
                "C34": ws.Cells(34, 3).Formula,
                "O34": ws.Cells(34, 15).Formula,
            }
        except Exception as exc:  # noqa: BLE001
            self.skipTest(f"无法通过 Excel 读取公式，跳过校验：{exc}")
        finally:
            if wb is not None:
                try:
                    wb.Close(False)
                except Exception:  # noqa: BLE001
                    pass
            if app is not None:
                try:
                    app.Quit()
                except Exception:  # noqa: BLE001
                    pass

        for cell, f in formulas.items():
            self.assertIsInstance(f, str, f"{cell} 的公式应为字符串，实际 {f!r}")
            self.assertTrue(str(f).lstrip().startswith("="),
                            f"{cell} 应为公式，实际 {f!r}。"
                            f"若模板确实改了，请评估是否需打开 "
                            f"WRITE_YEAR_TOTAL / WRITE_TOTAL_ROW")


if __name__ == "__main__":
    unittest.main(verbosity=2)