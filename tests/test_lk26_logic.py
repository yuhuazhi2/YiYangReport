# -*- coding: utf-8 -*-
"""LK26(B31) 管理费用明细表的逻辑测试。

运行方式（在项目根目录执行）：
    python -m unittest discover -s tests -v

【本文件测什么？】
    1. **C4 会计期间（需求第 1 条）** —— 本次开发的核心要求：
       界面选何年何月，C4 就填何年何月。
       覆盖：不同年月组合、无数据、取数失败、"是否写死常量"、
             是否用 set_text、位置是否确为 C4。
    2. 行项目配置与模板版式是否对得上（35 行、行次 1~35、第 9~43 行）
    3. 科目未配置时的行为：**不查库、不写数**，但页眉与会计期间照常填
    4. 科目配好后的取数与写入（用临时假科目验证框架逻辑）：
       · 各月数落到正确的列
       · 取数上限是"所选月份"，之后月份绝不写入
       · 同一行的多个科目累加
       · 减项科目（exclude_codes）生效
       · 无发生额的月份保持空白（不写 0）
    5. **O 列与第 44 行合计不得由程序写入**（模板公式，写了就覆盖公式）
    6. 【结构校验】cell_map 与 ROW_SPECS 一致、格子不重复、模板地址有效
    7. 【模板疑点】第 44 行列合计公式只到第 40 行 —— 显式锁住这个事实

【为什么大量用例要"临时给科目"？】
    本次开发阶段 35 行的科目清单尚未由财务提供，ROW_SPECS 里 codes 全空。
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

from reports import lk26_admin_expense as mod  # noqa: E402
from reports.lk26_admin_expense import (  # noqa: E402
    COL_OF_MONTH,
    COL_YEAR_TOTAL,
    DIGEST_KEYWORDS,
    FIRST_DATA_ROW,
    HEADER_LABEL_CELLS,
    LAST_DATA_ROW,
    PERIOD_CELL,
    PENDING_ROWS,
    ROW_SPECS,
    TOTAL_ROW,
    TOTAL_ROW_NO,
    LK26Report,
)

# 模板表页名（与模板中的页签完全一致）
SHEET_NAME = "LK26(B31)"


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


def _run(rows=None, year=2026, month=5, fail=False, hook=None):
    """跑一遍 fetch + write，返回 (hook, ctx)。"""
    ctx = FakeCtx(rows or [], year=year, month=month, fail=fail)
    hook = hook or FakeHook()
    rep = LK26Report()
    rep.fetch(ctx)
    rep.write(hook, ctx)
    return hook, ctx


# =====================================================================
# 一、行项目配置与模板版式
# =====================================================================
class TestRowConfig(unittest.TestCase):
    """行项目配置与模板版式的对应关系。"""

    def test_row_count_matches_template(self):
        """模板数据区共 35 个费用项目（行次 1~35）。"""
        self.assertEqual(len(ROW_SPECS), 35,
                         f"应为 35 行，实际 {len(ROW_SPECS)} 行")

    def test_row_numbers_are_sequential(self):
        """行次必须连续 1~35，不能跳号或重号。

        【为什么测这个？】
            行次是模板 B 列的编号，也是本表全部格位的推算依据。
            若某行行次写错，该行所有 13 个格子都会整体错位，
            数字落到别的项目上，而报表上每格都"有数字"，极难发现。
        """
        self.assertEqual([s.row_no for s in ROW_SPECS],
                         list(range(1, 36)),
                         "行次应为 1~35 连续排列")

    def test_excel_row_mapping(self):
        """行次 → Excel 行号：行次 1 在第 9 行，行次 35 在第 43 行。"""
        self.assertEqual(FIRST_DATA_ROW, 9)
        self.assertEqual(LAST_DATA_ROW, 43)
        self.assertEqual(ROW_SPECS[0].excel_row, 9)
        self.assertEqual(ROW_SPECS[-1].excel_row, 43)

    def test_total_row_position(self):
        """合计行在第 44 行、行次 36（模板已核对）。"""
        self.assertEqual(TOTAL_ROW, 44)
        self.assertEqual(TOTAL_ROW_NO, 36)

    def test_keys_unique(self):
        """行的 key 不得重复（否则取数会分流到同一行）。"""
        keys = [s.key for s in ROW_SPECS]
        self.assertEqual(len(keys), len(set(keys)),
                         f"存在重复的行 key：{keys}")

    def test_labels_unique(self):
        """行标签也不宜重复（便于看日志时区分）。"""
        labels = [s.label for s in ROW_SPECS]
        self.assertEqual(len(labels), len(set(labels)),
                         f"存在重复的行标签：{labels}")

    def test_codes_are_strings(self):
        """科目编码必须是字符串（写成数字可能导致匹配不上）。"""
        for spec in ROW_SPECS:
            for c in (*spec.codes, *spec.exclude_codes):
                self.assertIsInstance(
                    c, str, f"{spec.key} 的科目编码 {c!r} 不是字符串")

    def test_configured_rows_match_requirements(self):
        """三态分布必须与需求逐条对应：4 已配置 + 6 暂不处理 + 25 待提供。

        【为什么锁这条？】
            若有人凭猜测给某行填了科目，报表会静默产出
            **看起来正常但口径不对**的数 —— 报表上每格都有数字，
            只有逐笔核对才能发现。

            每补一行需求，就把该行加进下面的 expected 里（或移入
            expected_pending），让三态清单始终与需求逐条对应。
        """
        expected = {
            "工资": ("660201",),              # 需求第 2 条
            "职工福利费": ("660202",),        # 需求第 3 条
            "办公费": ("660203",),            # 需求第 4 条
            "差旅费": ("660204",),            # 需求第 5 条
            "修理费": ("660209",),            # 需求第 6 条
            "机物料消耗": ("660206",),        # 需求第 7 条
            "社保及公积金": ("660207", "660238"),  # 需求第 8 条（双科目）
            "交际应酬费": ("660208",),        # 需求第 9 条
            "汽车费用": ("660249",),          # 需求第 10 条
            "低值易耗品摊销": ("660210",),    # 需求第 11 条
            "折旧费": ("66021101",),          # 需求第 12 条（8 位明细）
            "租赁费": ("660212",),            # 需求第 13 条
            "水电费": ("660213",),            # 需求第 14 条
            "通讯费": ("660214",),            # 需求第 15 条
            "劳动保护费": ("660234",),        # 需求第 16 条
            "工会经费": ("660240",),          # 需求第 17 条
            "职工教育费": ("660220",),        # 需求第 18 条
            "检验检测费用": ("660215",),      # 需求第 19 条
            "无形资产摊销": ("66021102",),    # 需求第 20 条（8 位明细）
            "安全费用": ("660251",),          # 需求第 21/22 条（两条重复）
            "顾问及咨询费": ("660219",),      # 需求第 23 条
            "会务费": ("660224",),            # 需求第 24 条
            "新产品研发费用": ("660253",),    # 需求第 25 条
            "审计、鉴证费": ("660228",),      # 需求第 26 条
            "环保费用": ("660252",),          # 需求第 27 条
            "存货盘亏损失": ("660223",),      # 需求第 28 条
            "物业管理费": ("660216",),        # 需求第 29 条
            "残疾人就业保障金": ("66021704",),  # 需求第 30 条（8 位明细）
            "其他": ("660299",),              # 需求第 31 条（兜底行）
        }
        actual = {s.key: s.codes for s in ROW_SPECS if s.codes}
        self.assertEqual(actual, expected,
                         f"已配置行与需求不符，实际：{actual}")

        # 【暂不处理 —— 现场确认口径不明（20261112）】
        #   【注意排序】sorted() 按 Unicode 码点排，这里按实际排序写死，
        #   避免每次都要重新推一遍顺序，也避免"看起来顺序变了"的假警报。
        expected_pending = sorted([
            "商业保险费", "人力招聘费", "员工辞退福利",
            "律师、诉讼费", "协会会费", "用工劳务费",
        ])
        pending = sorted(s.key for s in ROW_SPECS if s.pending)
        self.assertEqual(pending, expected_pending,
                         f"暂不处理的应有六行，实际：{pending}")

        # 三态必须恰好覆盖全部 35 行，不能有第四种状态漏网
        todo = [s.key for s in ROW_SPECS if s.status == "待配置"]
        self.assertEqual(
            len(actual) + len(pending) + len(todo), len(ROW_SPECS),
            f"已配置({len(actual)}) + 暂不处理({len(pending)}) "
            f"+ 待配置({len(todo)}) 应恰好等于 {len(ROW_SPECS)} 行")
        # 【需求已全部到位：29 已配置 + 6 暂不处理 = 35，待配置为 0】
        self.assertEqual(len(todo), 0,
                         f"需求已全部到位，待配置应为 0 行，"
                         f"实际 {len(todo)} 行：{todo}")
        self.assertEqual(len(actual), 29,
                         f"已配置应为 29 行，实际 {len(actual)} 行")

    def test_first_four_rows_codes_and_positions(self):
        """需求 2~5 条：前四行的科目与格位。

            C9  工资        → 660201
            C10 职工福利费  → 660202
            C11 办公费      → 660203
            C12 差旅费      → 660204

        【为什么四行一起测？】
            四行口径完全相同、格位连续（9~12 行）、科目编码也连续
            （660201~660204）。复制粘贴时最容易出的问题是
            "漏改某一行的科目或行号"：
              · 660203 写成 660202 → 办公费显示福利费的数
              · 行次 3 写成 4     → 办公费的数落到差旅费行

            四行一起测，任一处错位都会立刻暴露。
        """
        expected = {
            "工资": (("660201",), 9),
            "职工福利费": (("660202",), 10),
            "办公费": (("660203",), 11),
            "差旅费": (("660204",), 12),
        }
        for key, (codes, excel_row) in expected.items():
            with self.subTest(key=key):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.codes, codes,
                                 f"{key} 科目应为 {codes}，实际 {spec.codes}")
                self.assertEqual(spec.excel_row, excel_row,
                                 f"{key} 应落在第 {excel_row} 行"
                                 f"（C{excel_row}），实际 {spec.excel_row}")

    def test_rows_6_to_12_codes_and_positions(self):
        """需求 6~12 条：七行的科目与格位（本组含三处编码陷阱）。

            C13 修理费         → 660209
            C14 机物料消耗     → 660206      ← 需求称"物料消耗"
            C15 社保及公积金   → 660207+660238 ← 双科目合计
            C17 交际应酬费     → 660208      ← 需求称"业务招待费"
            C18 汽车费用       → 660249      ← 编码跳号
            C19 低值易耗品摊销 → 660210
            C20 折旧费         → 66021101    ← 8 位明细

        【为什么七行必须一起测？】
            这七行的科目编码**完全不按行序递增**：
                660209 → 660206 → 660207/660238 → 660208
                → 660249 → 660210 → 66021101
            忽大忽小、中间还跳号。若有人按"上一行 +N"的直觉去推算，
            会配出完全不存在的科目，该行永远空白 —— 报表上看不出异常。

            七行一起测，任一处错位都会立刻暴露。
        """
        expected = {
            "修理费": (("660209",), 13),
            "机物料消耗": (("660206",), 14),
            "社保及公积金": (("660207", "660238"), 15),
            "交际应酬费": (("660208",), 17),
            "汽车费用": (("660249",), 18),
            "低值易耗品摊销": (("660210",), 19),
            "折旧费": (("66021101",), 20),
        }
        for key, (codes, excel_row) in expected.items():
            with self.subTest(key=key):
                spec = next(s for s in ROW_SPECS if s.key == key)
                self.assertEqual(spec.codes, codes,
                                 f"{key} 科目应为 {codes}，实际 {spec.codes}")
                self.assertEqual(spec.excel_row, excel_row,
                                 f"{key} 应落在第 {excel_row} 行"
                                 f"（C{excel_row}），实际 {spec.excel_row}")

    def test_social_insurance_has_two_codes(self):
        """社保及公积金必须是**两个**科目（660207 + 660238）。

        【为什么单独锁这条？】
            本行是当前唯一的多科目合计行。若只配了一个科目，
            报表会显示一个偏小的数 —— 而看不出少了公积金那一块。
        """
        spec = next(s for s in ROW_SPECS if s.key == "社保及公积金")
        self.assertEqual(spec.codes, ("660207", "660238"),
                         f"社保及公积金应为两科目合计，实际 {spec.codes}")

    def test_depreciation_uses_eight_digit_detail_code(self):
        """折旧费必须是 8 位明细 66021101，不能是 6 位的 660211。

        【为什么这条值得单测？】
            本表用 ccode IN (...) **精确匹配**。若写成 660211（6 位），
            只能取到一级科目自身那条记录 —— 通常为空，
            结果是折旧费整行空白，而报表上看不出"少了明细"。

            代码本身不会报错，SQL 也正常执行，只是返回空结果。
        """
        spec = next(s for s in ROW_SPECS if s.key == "折旧费")
        self.assertEqual(spec.codes, ("66021101",),
                         f"折旧费应为 8 位明细 66021101，实际 {spec.codes}")
        self.assertEqual(len(spec.codes[0]), 8,
                         f"折旧费科目应为 8 位，实际 {len(spec.codes[0])} 位")

    def test_car_expense_code_does_not_follow_sequence(self):
        """汽车费用是 660249，不是按顺序推算的 660209/660210。

        【为什么锁这条？】
            上一行交际应酬费是 660208，若按"下一行应该是 209"去猜，
            会把修理费的科目（660209）配到汽车费用上 ——
            两行都会显示"正常的数字"，只有逐笔核对才能发现。

            而 660209 已归属修理费（行次 5），本条同时验证它没被挪用。
        """
        car = next(s for s in ROW_SPECS if s.key == "汽车费用")
        repair = next(s for s in ROW_SPECS if s.key == "修理费")
        self.assertEqual(car.codes, ("660249",),
                         f"汽车费用应为 660249，实际 {car.codes}")
        self.assertEqual(repair.codes, ("660209",),
                         f"修理费应为 660209，实际 {repair.codes}")
        self.assertNotEqual(car.codes, repair.codes,
                            "汽车费用与修理费不应共用同一科目")

    def test_each_code_belongs_to_exactly_one_row(self):
        """每个科目编码只能归属一行（防止同一笔金额被算两次）。

        【为什么重要？】
            若 660210 同时出现在两行，同一笔低值易耗品摊销会被
            算两次，两行各显示一部分金额 —— 而每格看起来都正常。
            本条遍历全部已配置行，统计每个编码出现次数，必须全为 1。
        """
        seen: dict[str, str] = {}
        for spec in ROW_SPECS:
            for c in spec.codes:
                self.assertNotIn(
                    c, seen,
                    f"科目 {c} 同时被「{seen.get(c)}」和「{spec.key}」使用，"
                    f"同一笔金额会被重复计算")
                seen[c] = spec.key

        # 【数量校验】11 行已配置，其中社保及公积金占 2 个科目，
        #   且没有任何重复，故应有 12 个不同科目。
        #
        #   【为什么用"行数 + 多科目行多出的个数"表达，而不写死 12？】
        #     写死数字在补行时总要跟着改。用这个表达式，
        #     增删行时本条自动成立，只在**真出现重复**时失败 ——
        #     这正是本用例想守的东西。
        configured = [s for s in ROW_SPECS if s.codes]
        extra = sum(len(s.codes) - 1 for s in configured)
        self.assertEqual(len(seen), len(configured) + extra,
                         f"应有 {len(configured) + extra} 个不同科目，"
                         f"实际 {len(seen)} 个：{sorted(seen)}")

    def test_pending_rows_registered(self):
        """PENDING_ROWS 应登记六行，且与 pending=True 完全一致。

        【这条是"配置 ↔ 人工清单"的交叉校验】
            pending=True 让程序不写这六行；PENDING_ROWS 让人一眼看全。
            两者若不一致（例如补了 pending 却忘了登记，或口径明确后
            只删了一边），本条会立刻失败 ——
            而这类不一致在报表上完全看不出来（都只是空白）。
        """
        self.assertEqual(len(PENDING_ROWS), 6,
                         f"应登记六行，实际 {len(PENDING_ROWS)} 行")

        # 由 pending=True 的各行反算出应有的 (项目名, 行次)
        expected = sorted(
            (spec.label, spec.row_no)
            for spec in ROW_SPECS if spec.pending
        )
        actual = sorted((name, row_no)
                        for name, _cells, row_no in PENDING_ROWS)
        self.assertEqual(
            actual, expected,
            f"PENDING_ROWS 与 pending=True 不一致：\n"
            f"  清单：{actual}\n"
            f"  specs：{expected}")

    def test_pending_rows_cells_match_excel_row(self):
        """PENDING_ROWS 里的格位必须与行次自洽（起始行 = 行次 + 8，列 C~N）。

        【为什么值得校验？】
            格位是给**人工核对**用的。若行次填对但格位写成别的行，
            维护者按清单去 Excel 里找会找错行 —— 而这种错
            在程序运行上完全没有症状（程序不看这个清单取数）。
        """
        for name, cells, row_no in PENDING_ROWS:
            with self.subTest(name=name):
                start, end = cells.split(":")
                self.assertEqual(start[0], "C",
                                 f"{name} 起始列应为 C，实际 {start}")
                self.assertEqual(end[0], "N",
                                 f"{name} 结束列应为 N，实际 {end}")

                start_row = int("".join(ch for ch in start if ch.isdigit()))
                end_row = int("".join(ch for ch in end if ch.isdigit()))
                self.assertEqual(start_row, row_no + 8,
                                 f"{name} 行次 {row_no} 应对应第 "
                                 f"{row_no + 8} 行，格位却写成 {cells}")
                self.assertEqual(end_row, start_row,
                                 f"{name} 的起止行号应一致，实际 {cells}")


# =====================================================================
# 二、需求第 1 条：C4 会计期间  ★ 本次开发的核心
# =====================================================================
class TestPeriodCell(unittest.TestCase):
    """需求第 1 条：界面选何年何月，C4 就填何年何月。

    【本组用例的设计意图】
        单测一个 "2026年5月" 是**不足以**证明"年月随选择变化"的 ——
        代码里写死一个常量同样能通过那一条。

        所以这里覆盖多种情形：
          · 不同年份（2025 / 2026 / 2027）
          · 不同月份（1 月 / 5 月 / 12 月）
          · 有数据 / 无数据 / 取数失败
        只要程序在某处把年月写死，或把它与"是否有数据"绑在一起，
        这些用例必然失败。
    """

    def test_period_cell_is_c4(self):
        """需求指定的位置是 C4，不能被改成别处。"""
        self.assertEqual(PERIOD_CELL, "C4")

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
                hook, _ = _run(year=year, month=month)
                self.assertEqual(
                    hook.get(PERIOD_CELL), expected,
                    f"界面选 {year}年{month}月，C4 应为 {expected!r}，"
                    f"实际 {hook.get(PERIOD_CELL)!r}")

    def test_period_not_hardcoded(self):
        """不同月份生成的 C4 必须不同 —— 防止年月被写成常量。

        【为什么单独测"不同"？】
            若代码写成 C4="2026年5月" 常量，上面的 5 月用例仍会通过。
            本条要求"选 5 月"与"选 8 月"的结果不一样，
            常量写法无法同时满足。
        """
        may, _ = _run(year=2026, month=5)
        aug, _ = _run(year=2026, month=8)
        self.assertNotEqual(may.get(PERIOD_CELL), aug.get(PERIOD_CELL),
                            f"不同月份生成出相同的年月 "
                            f"{may.get(PERIOD_CELL)!r} —— 年月疑似被写成了常量")

    def test_period_year_changes_too(self):
        """仅换年份（月相同）也必须体现 —— 防止只拼了月份。"""
        a, _ = _run(year=2025, month=6)
        b, _ = _run(year=2026, month=6)
        self.assertEqual(a.get(PERIOD_CELL), "2025年6月")
        self.assertEqual(b.get(PERIOD_CELL), "2026年6月")

    def test_period_written_without_data(self):
        """科目未配置（零数据）时，C4 仍必须是所选年月。"""
        for month in (1, 6, 12):
            with self.subTest(month=month):
                hook, _ = _run(year=2026, month=month)
                self.assertEqual(hook.get(PERIOD_CELL), f"2026年{month}月")

    def test_period_written_when_fetch_fails(self):
        """取数抛异常时，C4 仍必须是所选年月（不受取数失败影响）。"""
        original = _with_codes({"工资": ("660201",)})
        self.addCleanup(setattr, mod, "ROW_SPECS", original)

        hook, _ = _run(year=2026, month=9, fail=True)
        self.assertEqual(hook.get(PERIOD_CELL), "2026年9月",
                         "取数失败不应导致会计期间为空")

    def test_period_written_with_data(self):
        """有数据时，C4 也应写所选年月（不被数据流程覆盖）。"""
        original = _with_codes({"工资": ("660201",)})
        self.addCleanup(setattr, mod, "ROW_SPECS", original)

        hook, _ = _run(rows=[("660201", 7, 700.0)], year=2026, month=7)
        self.assertEqual(hook.get(PERIOD_CELL), "2026年7月")
        # 顺带确认数据也进了正确月份列。
        # 【注意列位】7 月是 I 列（C=1月…H=6月，I=7月，N=12月），
        #   不要凭"第 7 列就是 J"的直觉去数 —— J 其实是 8 月。
        self.assertTrue(approx(hook.get("I9"), 700.0), hook.get("I9"))

    def test_period_written_as_text(self):
        """C4 必须用 set_text 写，不能用 set。

        【为什么？】C4 所在列是金额格式（实测
        '_ * #,##0.00_ ;...'）。用 set() 写 "2026年5月"，
        Excel 会把它识别成日期并存成序列号 46143，
        再显示成 "46,143.00" —— 看起来像一笔钱（LK01 踩过的坑）。
        """
        hook, _ = _run(year=2026, month=5)
        kinds = [e[0] for e in hook.events if e[1] == PERIOD_CELL]
        self.assertIn("set_text", kinds,
                      f"C4 未用 set_text 写入，实际事件：{hook.events}")
        self.assertNotIn("set", kinds,
                         "C4 不应使用普通 set 写入（会被识别成日期）")

    def test_period_written_only_to_c4_among_value_cells(self):
        """年月本身只写 C4；A4 只承载"会计期间："标签，不含年月。

        【与上一版的差异（需求第 2 条）】
            上一版断言"A4 完全不被写入"，理由是 A4 是 =LK01!A4 引用公式。
            现在需求明确要求**清除该公式、改为纯文本标签**，所以 A4
            会被写入 —— 但写的是标签本身，不是年月。

            因此本条的断言方向改为：
              · A4 的文本里**不得出现年月**（否则同一页两个年月）；
              · B4 / D4 等真正的空白格仍不得被写入。
        """
        hook, _ = _run(year=2026, month=5)

        a4 = hook.get("A4")
        self.assertEqual(a4, "会计期间：",
                         f"A4 应为纯文本标签'会计期间：'，实际 {a4!r}")
        self.assertNotIn("2026", str(a4),
                         f"A4 不应含年月，否则与 C4 重复，实际 {a4!r}")
        self.assertNotIn("5月", str(a4),
                         f"A4 不应含年月，否则与 C4 重复，实际 {a4!r}")

        for cell in ("B4", "D4"):
            self.assertIsNone(hook.get(cell),
                              f"{cell} 不应被写入，实际 {hook.get(cell)}")


class TestHeaderLabels(unittest.TestCase):
    """需求第 2 条：A2 / A4 清除跨页引用公式，改为纯文本标签。

    【为什么要专门测这一条？】
        模板里 A2='=LK01!A2'、A4='=LK01!A4'，画面上的"公司名称:""
        会计期间："只是公式的缓存值。这带来两个只有真跑才会暴露的问题：

          ① 只生成 LK26 时，LK01 页被删除 → 公式断链 → #REF!；
             fill_common_header 按内容搜索读不到关键字，公司名
             **静默填不进去**，且不报任何错。
          ② 同时生成 LK01 与 LK26 时，两页页眉被隐式绑定 ——
             改 LK01 的页眉，LK26 跟着变。43 张表并行开发时，
             这类耦合会让"改一张表弄坏另一张"变成常态。

        所以本表必须在 write() 里主动把公式换成文本。
    """

    def test_header_cells_defined(self):
        """A2 / A4 两格都应登记在 HEADER_LABEL_CELLS 中。"""
        cells = dict(HEADER_LABEL_CELLS)
        self.assertIn("A2", cells, "A2 应登记为页眉标签格")
        self.assertIn("A4", cells, "A4 应登记为页眉标签格")

    def test_header_labels_keep_keywords(self):
        """标签文本必须含"公司名称""会计期间"关键字。

        【为什么锁这条？】
            fill_common_header 靠"搜索含『公司名称』的格子"决定写哪。
            若把措辞改成"单位名称"，公司名再也填不进去，
            而且**不会报任何错** —— 只能靠人打开文件肉眼发现。
        """
        cells = dict(HEADER_LABEL_CELLS)
        self.assertIn("公司名称", cells["A2"],
                      f"A2 文本须含'公司名称'，实际 {cells['A2']!r}")
        self.assertIn("会计期间", cells["A4"],
                      f"A4 文本须含'会计期间'，实际 {cells['A4']!r}")

    def test_header_labels_written_as_text(self):
        """A2 / A4 必须用 set_text 写（覆盖公式 + 防日期误判）。"""
        hook, _ = _run(year=2026, month=5)

        for cell, text in HEADER_LABEL_CELLS:
            with self.subTest(cell=cell):
                self.assertEqual(hook.get(cell), text,
                                 f"{cell} 应为 {text!r}，"
                                 f"实际 {hook.get(cell)!r}")
                kinds = [e[0] for e in hook.events if e[1] == cell]
                self.assertIn("set_text", kinds,
                              f"{cell} 未用 set_text 写入，"
                              f"实际事件：{hook.events}")

    def test_header_labels_written_before_period(self):
        """A2 / A4 的写入必须发生在写 C4 之前（顺序要求）。

        【为什么顺序重要？】
            fill_common_header 按内容搜索"公司名称"定位该写哪格。
            若标签还没写、格子仍是公式（或 #REF!），它就搜不到，
            公司名会被静默跳过 —— 报表上看不出异常，只是没有公司名。
            同理，C4 也应排在标签之后，保持"先版式、后数据"的顺序。

        【为什么不直接断言"早于 fill_common_header 的追加写入"？】
            FakeHook.find_cell() 恒返回 (0,0)，fill_common_header 在
            假对象下**不会发起任何写入** —— 若按那个事件断言，
            这条用例在测试环境里永远无法通过，只能靠真 Excel 验证。
            所以改为断言它早于 C4 的写入，同样能守住顺序。
        """
        hook, _ = _run(year=2026, month=5)

        label_idx = {
            cell: next((i for i, e in enumerate(hook.events)
                        if e[0] == "set_text" and e[1] == cell), None)
            for cell, _ in HEADER_LABEL_CELLS
        }
        for cell, idx in label_idx.items():
            self.assertIsNotNone(
                idx, f"未找到 {cell} 标签写入事件：{hook.events}")

        period_idx = next(
            (i for i, e in enumerate(hook.events)
             if e[0] == "set_text" and e[1] == PERIOD_CELL), None)
        self.assertIsNotNone(period_idx,
                             f"未找到 C4 写入事件：{hook.events}")

        for cell, idx in label_idx.items():
            self.assertLess(
                idx, period_idx,
                f"{cell} 的标签写入应早于 C4，"
                f"实际下标 {idx} vs {period_idx}：{hook.events}")

    def test_header_labels_written_even_without_data(self):
        """无任何数据时，A2 / A4 标签也必须写入。

        页眉标签是报表的"身份信息"，与有没有取到数无关。
        """
        hook, _ = _run(year=2026, month=12)
        for cell, text in HEADER_LABEL_CELLS:
            self.assertEqual(hook.get(cell), text,
                             f"{cell} 即使无数据也应为 {text!r}，"
                             f"实际 {hook.get(cell)!r}")

    def test_header_labels_written_when_fetch_fails(self):
        """取数抛异常时，页眉标签仍须写入。"""
        hook, _ = _run(year=2026, month=9, fail=True)
        for cell, text in HEADER_LABEL_CELLS:
            self.assertEqual(hook.get(cell), text,
                             f"取数失败不应导致 {cell} 为空，"
                             f"实际 {hook.get(cell)!r}")

    def test_template_header_cells_are_formulas(self):
        """模板 A2 / A4 确实是公式（"需要清除"这一判断的前提）。

        【为什么测"模板里是公式"？】
            本表之所以要主动写入 A2/A4，前提是模板里这两格为公式。
            若模板将来改成纯文本，本条会失败并提醒：
            "这条需求已由模板本身满足，程序里的写入可以评估是否保留"。
            模板不存在或 xlrd 不可用时自动跳过。
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
        sh = wb.sheet_by_name(SHEET_NAME)

        # A2 是第 2 行第 1 列（0-based: 1, 0）；A4 是 (3, 0)
        for coord, cell in (((1, 0), "A2"), ((3, 0), "A4")):
            self.assertEqual(
                sh.cell_type(*coord), xlrd.XL_CELL_TEXT,
                f"模板 {cell} 应读到文本（公式的缓存值），"
                f"实际类型 {sh.cell_type(*coord)}")
            self.assertTrue(
                str(sh.cell_value(*coord)).strip(),
                f"模板 {cell} 的缓存值不应为空")


# =====================================================================
# 三、科目未配置时的行为
# =====================================================================
class TestUnconfiguredBehaviour(unittest.TestCase):
    """科目未配置时的行为：不查库、不写数，但页眉与期间照常。"""

    def test_no_database_query_when_nothing_configured(self):
        """若所有行科目都为空，则不应发起任何查询。

        【测试手法】临时把 ROW_SPECS 全部清空，模拟"科目尚未提供"的状态，
        验证 fetch 的短路逻辑仍然有效（这是新增行时可能触碰的路径）。

        【★ 为什么必须清空，不能直接用真实配置？★】
            需求第 2 条已为「工资」行配了 660201，它是"已配置"行，
            此时**发一次查询才是正确行为**。若本条仍按真实配置断言
            "零查询"，就会把正确行为判成失败 —— 那是测试写错了。
        """
        original = _with_codes({})       # 全部置空
        self.addCleanup(setattr, mod, "ROW_SPECS", original)

        ctx = FakeCtx([], month=5)
        LK26Report().fetch(ctx)
        self.assertEqual(ctx.fetch_calls, [],
                         f"科目全为空时不应查库，实际发起了 "
                         f"{len(ctx.fetch_calls)} 次查询")

    def test_unconfigured_rows_not_written(self):
        """未配置科目的行不应写入任何金额（不写 0）。

        【为什么跳过已配置行？】
            需求第 2 条已为「工资」行配了 660201，它现在属"已配置"，
            有数据时就该写。本条只针对仍待配置的其余 34 行。
        """
        hook, _ = _run()
        for spec in ROW_SPECS:
            if spec.codes:
                continue          # 已配置行不在本条范围内
            for col in COL_OF_MONTH.values():
                cell = f"{col}{spec.excel_row}"
                self.assertIsNone(hook.get(cell),
                                  f"{spec.key} 的 {cell} 不应被写入，"
                                  f"实际 {hook.get(cell)}")

    def test_year_total_and_total_row_not_written(self):
        """O 列与第 44 行在无数据时也不得写入。"""
        hook, _ = _run()
        for spec in ROW_SPECS:
            cell = f"{COL_YEAR_TOTAL}{spec.excel_row}"
            self.assertIsNone(hook.get(cell),
                              f"{cell} 不应被写入，实际 {hook.get(cell)}")
        # 【第 44 行：写公式，不写值】需求第 32 条
        for col in COL_OF_MONTH.values():
            cell = f"{col}{TOTAL_ROW}"
            formula = hook.formulas.get(cell)
            self.assertIsNotNone(
                formula, f"{cell} 应写入修正后的合计公式")
            self.assertTrue(str(formula).startswith("="),
                            f"{cell} 应是公式而非数值，实际 {formula!r}")

    def test_period_still_written(self):
        """即使没有任何数据，会计期间仍应写入 C4。

        年月是报表的"身份信息"，与有没有取到数无关。
        若因为"无数据"就不写，打开文件根本分不清这是哪个月的报表。
        """
        hook, _ = _run(year=2026, month=5)
        self.assertEqual(hook.get(PERIOD_CELL), "2026年5月")


# =====================================================================
# 四、科目配好后的取数与写入（临时假科目）
# =====================================================================
class TestRowValues(unittest.TestCase):
    """科目配好后的取数与写入。"""

    def setUp(self):
        # 用 66 开头的假科目（管理费用真实编码段），避免与 LK25 的
        # 51 段混淆；这些编码仅用于测试，不写入 ROW_SPECS 的正式配置。
        original = _with_codes({
            "工资": ("660201", "660202"),   # 本行由两个科目构成
            "折旧费": ("660210",),
            "其他": ("660299",),
        })
        self.addCleanup(setattr, mod, "ROW_SPECS", original)

    def test_month_columns(self):
        """各月发生额应落到对应月份列。"""
        rows = [("660201", m, m * 100.0) for m in range(1, 6)]
        hook, _ = _run(rows=rows, year=2026, month=5)

        # 工资在第 9 行；C=1月 … G=5月
        for col, m in zip("CDEFG", range(1, 6)):
            self.assertTrue(
                approx(hook.get(f"{col}9"), m * 100.0),
                f"{col}9（{m}月）应为 {m * 100.0}，"
                f"实际 {hook.get(f'{col}9')}")

    def test_multi_codes_summed(self):
        """同一行的多个科目应相加。

        【为什么重要？】"社保及公积金"这类项目在实务中一定由多个
        明细科目构成。若程序只取第一个科目，报表会显示一个偏小的数，
        而看不出少了什么。
        """
        rows = [("660201", 5, 100.0), ("660202", 5, 25.0)]
        hook, _ = _run(rows=rows, year=2026, month=5)
        self.assertTrue(approx(hook.get("G9"), 125.0),
                        f"工资应为 100+25=125，实际 {hook.get('G9')}")

    def test_rows_not_crossed(self):
        """不同行各取各的科目，互不串用。"""
        rows = [
            ("660201", 5, 100.0),   # 工资 → 第 9 行
            ("660210", 5, 900.0),   # 折旧费 → 第 20 行
        ]
        hook, _ = _run(rows=rows, year=2026, month=5)
        self.assertTrue(approx(hook.get("G9"), 100.0), hook.get("G9"))
        self.assertTrue(approx(hook.get("G20"), 900.0), hook.get("G20"))

    def test_future_months_never_written(self):
        """所选期间之后的月份绝不写入。

        【为什么必须守住？】
            "2026年5月"的表里若混入 6 月的数，这张表就不再是 5 月表了，
            而看表的人无从分辨哪几列属于本期。
        """
        rows = [
            ("660201", 5, 500.0),
            ("660201", 6, 600.0),    # 超出所选期间，必须被丢弃
            ("660201", 12, 1200.0),
        ]
        hook, _ = _run(rows=rows, year=2026, month=5)
        self.assertTrue(approx(hook.get("G9"), 500.0))
        self.assertIsNone(hook.get("H9"), f"6月不应写入，实际 {hook.get('H9')}")
        self.assertIsNone(hook.get("N9"), f"12月不应写入，实际 {hook.get('N9')}")

    def test_accumulates_one_to_selected_month(self):
        """累计填列：选 3 月时 C/D/E 三列都要有数（与 LK25 同口径）。"""
        rows = [("660201", m, m * 100.0) for m in range(1, 13)]
        hook, _ = _run(rows=rows, year=2026, month=3)

        for col, m in (("C", 1), ("D", 2), ("E", 3)):
            self.assertTrue(approx(hook.get(f"{col}9"), m * 100.0),
                            f"{col}9 应为 {m * 100}，"
                            f"实际 {hook.get(f'{col}9')}")
        for col in "FGHIJKLMN":
            self.assertIsNone(hook.get(f"{col}9"),
                              f"{col}9 属 4 月及以后，3 月报表中不应写入")

    def test_missing_month_stays_blank(self):
        """某月无发生额时该月格保持空白（不写 0），其余月份照常填。"""
        rows = [("660201", m, m * 100.0) for m in (1, 2, 4, 5)]  # 缺 3 月
        hook, _ = _run(rows=rows, year=2026, month=5)

        self.assertTrue(approx(hook.get("C9"), 100.0), hook.get("C9"))
        self.assertTrue(approx(hook.get("D9"), 200.0), hook.get("D9"))
        self.assertIsNone(hook.get("E9"),
                          f"3 月无数据应保持空白，实际 {hook.get('E9')}")
        self.assertTrue(approx(hook.get("F9"), 400.0), hook.get("F9"))
        self.assertTrue(approx(hook.get("G9"), 500.0), hook.get("G9"))

    def test_year_total_not_written(self):
        """O 列（本年累计）不得由程序写入 —— 它是模板公式 =SUM(C:N)。

        【需求 20261112 现场确认】O 列各行保持当前现状，不做处理。

        【为什么这里不再断言某个开关为 False？】
            原先有一个 `WRITE_YEAR_TOTAL = False` 的模块常量。
            但"不做处理"是明确结论，不存在"需要打开"的场景 ——
            留一个恒为 False 的开关，会给后来者"也许该打开它"的错觉。
            所以该开关已被**彻底移除**：不是关掉，而是不存在。
        """
        hook, _ = _run(rows=[("660201", 5, 500.0)], year=2026, month=5)
        self.assertIsNone(hook.get(f"{COL_YEAR_TOTAL}9"),
                          f"O9 不应被写入，"
                          f"实际 {hook.get(f'{COL_YEAR_TOTAL}9')}")

    def test_year_total_switch_removed(self):
        """模块中不应再存在 WRITE_YEAR_TOTAL 常量。

        【为什么用"断言属性不存在"来锁住一条约束？】
            若日后有人从 LK25 复制代码回来、顺手把开关加回去，
            本条会立刻失败并提示：O 列是硬约束，不要再引入开关。
            这比在文档里写一句"不要加"有效得多 ——
            文档会被跳过，测试不会。
        """
        self.assertFalse(
            hasattr(mod, "WRITE_YEAR_TOTAL"),
            "WRITE_YEAR_TOTAL 已被移除（O 列不做处理是硬约束），"
            "若确需恢复，请先回到需求层面确认")

    def test_year_total_column_never_written_all_months(self):
        """选 12 月、全部月份都有数据时，O 列仍一格都不得写入。"""
        rows = [("660201", m, m * 100.0) for m in range(1, 13)]
        hook, _ = self._run_12(rows)

        for spec in ROW_SPECS:
            cell = f"{COL_YEAR_TOTAL}{spec.excel_row}"
            self.assertIsNone(hook.get(cell),
                              f"{cell} 不应被写入，实际 {hook.get(cell)}")
        self.assertIsNone(hook.get(f"{COL_YEAR_TOTAL}{TOTAL_ROW}"),
                          f"{COL_YEAR_TOTAL}{TOTAL_ROW} 不应被写入")

    def _run_12(self, rows):
        """辅助：选 12 月跑一遍（12 月才能覆盖全部月列）。"""
        ctx = FakeCtx(rows, year=2026, month=12)
        hook = FakeHook()
        rep = LK26Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)
        return hook, ctx

    def test_total_row_formulas_fixed(self):
        """第 44 行合计：程序把 C44~N44 的区间修正为 9:43（需求第 32 条）。

        【与旧版断言的本质区别】
            旧版断言"C44~N44 完全不被写入" —— 那时保留模板的
            '=SUM(C9:C40)'（漏了第 41~43 行）。

            现在需求要求覆盖全区间，程序**改写公式**（不是写值）：
                C44 ~ N44  →  '=SUM(X9:X43)'

        【为什么必须写公式而不是写值？】
            写值会让 Excel 失去自动汇总：下月改明细、合计不跟着变。
            写公式则每次重算都按当时的明细求和 —— 详见模块注释。
        """
        hook, _ = _run(rows=[("660201", 5, 500.0)], year=2026, month=5)

        for col in COL_OF_MONTH.values():
            cell = f"{col}{TOTAL_ROW}"
            formula = hook.formulas.get(cell)
            self.assertIsNotNone(
                formula,
                f"{cell} 应写入修正后的合计公式，实际未写入")
            self.assertEqual(
                formula, f"=SUM({col}9:{col}43)",
                f"{cell} 的公式应为 =SUM({col}9:{col}43)，实际 {formula!r}")

    def test_total_row_formula_covers_all_data_rows(self):
        """合计公式的区间必须覆盖第 9~43 行（与 O44 口径一致）。

        【为什么单独锁"区间端点"？】
            模板原公式是 '=SUM(C9:C40)' —— 差 3 行。
            只看"有没有公式"发现不了（原公式也是公式），
            必须逐字比对区间端点，才能抓出"漏了三行"这个缺陷。
        """
        hook, _ = _run(year=2026, month=5)
        for col in COL_OF_MONTH.values():
            formula = hook.formulas.get(f"{col}{TOTAL_ROW}", "")
            self.assertIn(f"{col}9:", formula,
                          f"{col}{TOTAL_ROW} 的起点应为 {col}9，"
                          f"实际 {formula!r}")
            self.assertIn(f":{col}43", formula,
                          f"{col}{TOTAL_ROW} 的终点应为 {col}43"
                          f"（不是 40），实际 {formula!r}")
            self.assertNotIn(f":{col}40", formula,
                             f"{col}{TOTAL_ROW} 不应仍是模板原区间 9:40，"
                             f"实际 {formula!r}")

    def test_total_row_written_as_formula_not_value(self):
        """合计行必须用 set_formula 写，不能用 set 写值。

        【为什么要区分？】
            两者在报表上都让 C44 "有数"，但后果完全不同：
              · set 写值 → 公式被覆盖，自动汇总能力**永久丢失**
              · set_formula → 公式仍在，Excel 每次重算都会更新
            只断言"格子有内容"发现不了这个区别。
        """
        hook, _ = _run(year=2026, month=5)
        for col in COL_OF_MONTH.values():
            cell = f"{col}{TOTAL_ROW}"
            kinds = [e[0] for e in hook.events if e[1] == cell]
            self.assertIn("set_formula", kinds,
                          f"{cell} 未用 set_formula 写入，实际事件：{kinds}")
            self.assertNotIn("set", kinds,
                             f"{cell} 不应使用普通 set 写值（会覆盖公式）")

    def test_total_row_switch_removed(self):
        """模块中不应存在 WRITE_TOTAL_ROW 常量（已改为硬约束）。"""
        self.assertFalse(
            hasattr(mod, "WRITE_TOTAL_ROW"),
            "WRITE_TOTAL_ROW 已被移除（合计行修正是硬约束），"
            "若确需恢复，请先回到需求层面确认")

    def test_total_row_formula_template_constant(self):
        """TOTAL_FORMULA_TEMPLATE 常量应存在且格式正确。"""
        self.assertTrue(hasattr(mod, "TOTAL_FORMULA_TEMPLATE"))
        tpl = mod.TOTAL_FORMULA_TEMPLATE
        for ph in ("{col}", "{first}", "{last}"):
            self.assertIn(ph, tpl)
        self.assertTrue(tpl.startswith("="), f"公式模板应以 = 开头：{tpl!r}")

    def test_fetch_failure_not_crash(self):
        """取数抛异常时不崩溃，且不写任何金额，但 C4 照写。"""
        hook, _ = _run(year=2026, month=5, fail=True)
        self.assertIsNone(hook.get("C9"))
        self.assertEqual(hook.get(PERIOD_CELL), "2026年5月")

    def test_single_query_for_all_rows(self):
        """所有行的科目应合并成一次查询（去重）。"""
        _, ctx = _run(year=2026, month=5)
        self.assertEqual(len(ctx.fetch_calls), 1,
                         f"应只发一次查询，实际 {len(ctx.fetch_calls)} 次")
        params = ctx.fetch_calls[0]["params"]
        codes = [p for p in params if isinstance(p, str)]
        self.assertEqual(len(codes), len(set(codes)),
                         f"科目参数有重复：{codes}")

    def test_sql_period_range_is_one_to_selected(self):
        """SQL 的 iYPeriod 区间必须是 202601~所选月份（不取未来月份）。"""
        for month in (1, 3, 8, 12):
            with self.subTest(month=month):
                _, ctx = _run(year=2026, month=month)
                params = ctx.fetch_calls[0]["params"]
                self.assertEqual(params[0], 2026,
                                 f"年度参数应为 2026，实际 {params}")
                self.assertEqual(params[1], 202601,
                                 f"区间起始应为 202601，实际 {params}")
                self.assertEqual(params[2], int(f"2026{month:02d}"),
                                 f"区间截止应为 2026{month:02d}，实际 {params}")

    def test_sql_uses_credit_side_and_digest(self):
        """SQL 必须取 mc（贷方）并带「期间损益结转」摘要筛选。"""
        _, ctx = _run(year=2026, month=5)
        sql = ctx.fetch_calls[0]["sql"]
        flat = sql.replace(" ", "")
        self.assertIn("SUM(mc)", flat, f"应取贷方 mc，实际 SQL：{sql}")
        self.assertNotIn("SUM(md)", flat, f"不应取借方 md，实际 SQL：{sql}")
        self.assertIn("cdigest LIKE", sql, f"缺少摘要筛选：{sql}")
        for kw in DIGEST_KEYWORDS:
            self.assertIn(f"%{kw}%", ctx.fetch_calls[0]["params"],
                          f"摘要参数应含 %{kw}%，"
                          f"实际 {ctx.fetch_calls[0]['params']}")


class TestExcludeCodes(unittest.TestCase):
    """减项科目（exclude_codes）的行为。"""

    def setUp(self):
        patched = []
        for spec in ROW_SPECS:
            if spec.key == "其他":
                patched.append(dataclasses.replace(
                    spec, codes=("660299",), exclude_codes=("660298",)))
            else:
                patched.append(dataclasses.replace(spec, codes=()))
        original = mod.ROW_SPECS
        mod.ROW_SPECS = tuple(patched)
        self.addCleanup(setattr, mod, "ROW_SPECS", original)

    def test_exclude_subtracted(self):
        """最终值 = sum(codes) - sum(exclude_codes)。"""
        rows = [
            ("660299", 5, 1000.0),
            ("660298", 5, 300.0),
        ]
        hook, _ = _run(rows=rows, year=2026, month=5)
        # 其他在第 43 行；5 月 = G 列
        self.assertTrue(approx(hook.get("G43"), 700.0),
                        f"其他应为 1000-300=700，实际 {hook.get('G43')}")

    def test_only_exclude_not_written(self):
        """只有减项科目有记录时不写负数。

        【为什么不写？】照算会得到 -300 这种不可能的金额。
        宁可留空让人查，也不写一个方向错误的数。
        """
        hook, _ = _run(rows=[("660298", 5, 300.0)], year=2026, month=5)
        self.assertIsNone(hook.get("G43"),
                          f"不应写入负数，实际 {hook.get('G43')}")


# =====================================================================
# 五、结构校验
# =====================================================================
class TestConfigConsistency(unittest.TestCase):
    """结构校验：防止配置与 cell_map / 模板脱节。"""

    def test_cell_map_covers_all_cells(self):
        """ROW_SPECS 的每一格（含 O 列）都应登记在 cell_map 中。

        【为什么重要？】cell_map 决定父类 verify() 读回哪些格做勾稽核对。
        漏登记 → 该行不进核对清单 → 日后"核对通过"可能是假象。
        """
        cm_values = set(LK26Report.cell_map.values())
        for spec in ROW_SPECS:
            for col in COL_OF_MONTH.values():
                cell = f"{col}{spec.excel_row}"
                self.assertIn(cell, cm_values,
                              f"{spec.key} 的 {cell} 未登记 cell_map")
            total_cell = f"{COL_YEAR_TOTAL}{spec.excel_row}"
            self.assertIn(total_cell, cm_values,
                          f"{spec.key} 的 {total_cell} 未登记 cell_map")

    def test_cell_map_includes_period(self):
        """C4 会计期间应登记进 cell_map（勾稽核对要读回年月）。"""
        self.assertEqual(LK26Report.cell_map.get("会计期间"), PERIOD_CELL)

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

    def test_month_columns_are_c_to_n(self):
        """12 个月的列必须正好是 C~N，且与模板表头一致。"""
        self.assertEqual(COL_OF_MONTH[1], "C")
        self.assertEqual(COL_OF_MONTH[12], "N")
        self.assertEqual(len(COL_OF_MONTH), 12)
        self.assertEqual("".join(COL_OF_MONTH[m] for m in range(1, 13)),
                         "CDEFGHIJKLMN")

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
        sh = wb.sheet_by_name(SHEET_NAME)

        for spec in ROW_SPECS:
            # A 列（0-based 索引 0）是项目名
            val = str(sh.cell_value(spec.excel_row - 1, 0)).strip()
            self.assertTrue(
                val,
                f"{spec.key} 位于第 {spec.excel_row} 行，"
                f"但模板 A 列是空的 —— 可能写错了行")

        # 合计行也要有项目名
        total_label = str(sh.cell_value(TOTAL_ROW - 1, 0)).strip()
        self.assertTrue(total_label,
                        f"第 {TOTAL_ROW} 行应为合计行，但模板 A 列是空的")

    def test_template_row_count_matches(self):
        """模板 LK26(B31) 的数据区行数与配置一致（防止模板改版）。"""
        try:
            import xlrd
        except ImportError:
            self.skipTest("未安装 xlrd")

        from config.settings import DEFAULT_TEMPLATE, PROJECT_ROOT

        tpl = PROJECT_ROOT / DEFAULT_TEMPLATE
        if not tpl.exists():
            self.skipTest(f"模板不存在：{tpl}")

        wb = xlrd.open_workbook(str(tpl))
        sh = wb.sheet_by_name(SHEET_NAME)

        # 行次 1 → 第 9 行；行次 35 → 第 43 行
        first_row_no = sh.cell_value(FIRST_DATA_ROW - 1, 1)
        last_row_no = sh.cell_value(LAST_DATA_ROW - 1, 1)
        self.assertEqual(int(first_row_no), 1,
                         f"模板第 {FIRST_DATA_ROW} 行的行次应为 1，"
                         f"实际 {first_row_no!r}")
        self.assertEqual(int(last_row_no), 35,
                         f"模板第 {LAST_DATA_ROW} 行的行次应为 35，"
                         f"实际 {last_row_no!r}")


class TestTemplateFormulas(unittest.TestCase):
    """模板 O 列与合计行确实带公式（"程序不写"的前提）。

    【★ 为什么整段都包进 try、失败就 skip？★】
        本组依赖本机真实 Excel（win32com），而可用性取决于很多
        与报表逻辑无关的环境因素：没装 pywin32、没装 Excel、
        类型库缓存损坏、已有 Excel 进程占用 COM 单例等。

        这些都不是"报表写错了"。若让它们判为 FAILED/ERROR，
        会造成持续的假警报 —— 维护者很快会忽略这条测试，
        届时**模板真被改坏也无人察觉**。

        所以：能连上 Excel 就严格校验；连不上就明确 skip。
        口径本身的校验（O 列不写、合计行不写等）由其它用例保证，
        不依赖本环境。
    """

    @classmethod
    def setUpClass(cls):
        from config.settings import DEFAULT_TEMPLATE, PROJECT_ROOT

        cls._app = None
        cls._wb = None
        cls._ws = None

        tpl = PROJECT_ROOT / DEFAULT_TEMPLATE
        if not tpl.exists():
            raise unittest.SkipTest(f"模板不存在：{tpl}")

        try:
            from core.excel_engine import _dispatch_excel
            app = _dispatch_excel()
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(str(tpl), ReadOnly=True, UpdateLinks=0)
            ws = wb.Sheets(SHEET_NAME)
        except Exception as exc:  # noqa: BLE001
            cls._cleanup(app if "app" in locals() else None, None)
            raise unittest.SkipTest(
                f"无法通过 Excel 读取公式，跳过校验：{exc}")

        cls._app = app
        cls._wb = wb
        cls._ws = ws

    @classmethod
    def _cleanup(cls, app, wb):
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

    @classmethod
    def tearDownClass(cls):
        cls._cleanup(cls._app, cls._wb)

    def _formula(self, row: int, col: int) -> str:
        return self._ws.Cells(row, col).Formula

    def test_year_total_columns_have_sum_formula(self):
        """O 列每行都应是 =SUM(C:N) 形式的公式。"""
        for row in (ROW_SPECS[0].excel_row, ROW_SPECS[-1].excel_row):
            f = self._formula(row, 15)      # 15 = O 列
            self.assertIsInstance(f, str, f"O{row} 应为公式，实际 {f!r}")
            self.assertTrue(str(f).lstrip().startswith("="),
                            f"O{row} 应为公式，实际 {f!r}")

    def test_total_row_has_sum_formula(self):
        """第 44 行（合计）应有公式。"""
        for col in (3, 15):                 # 3=C 列，15=O 列
            f = self._formula(TOTAL_ROW, col)
            self.assertIsInstance(f, str,
                                  f"第 {TOTAL_ROW} 行第 {col} 列应为公式，"
                                  f"实际 {f!r}")
            self.assertTrue(str(f).lstrip().startswith("="),
                            f"第 {TOTAL_ROW} 行第 {col} 列应为公式，"
                            f"实际 {f!r}")

    def test_total_row_column_formula_only_covers_to_row_40(self):
        """【模板疑点】第 44 行列合计公式只到第 40 行，漏了第 41~43 行。

        【为什么要"锁住"一个缺陷？】
            模板实测 C44 = '=SUM(C9:C40)'，而数据区到第 43 行。
            也就是说：物业管理费（第41行）、残疾人就业保障金（第42行）、
            其他（第43行）填了数，**列合计不会包含它们**。

            本程序的处理是"不改模板、不补写合计行"，并在交付说明中
            提请财务确认（详见 reports/lk26_admin_expense.py 的
            模块头部口径②）。

            本条测试把该事实固定下来：
              · 若模板将来修好（区间改成 C9:C43），本条会失败并提醒
                更新代码里的说明与交付结论 —— 这是"好事"的失败；
              · 若有人顺手把程序改成补写合计行，WRITE_TOTAL_ROW 的
                断言会先一步失败。
            两条一起，保证"模板与程序口径不一致"这件事不会被悄悄遗忘。
        """
        f = str(self._formula(TOTAL_ROW, 3)).replace(" ", "").upper()
        self.assertEqual(
            f, "=SUM(C9:C40)",
            f"第 {TOTAL_ROW} 行列合计公式已变化（实际 {f!r}）。"
            f"若模板已把区间修正为覆盖第 43 行，请同步更新 "
            f"lk26_admin_expense.py 模块头部口径②的说明与交付结论。")

        # 交叉验证：O 列合计则是全区间（两者口径不一致，正是疑点所在）
        f_o = str(self._formula(TOTAL_ROW, 15)).replace(" ", "").upper()
        self.assertEqual(f_o, "=SUM(O9:O43)",
                         f"第 {TOTAL_ROW} 行 O 列公式已变化（实际 {f_o!r}）")


if __name__ == "__main__":
    unittest.main(verbosity=2)

# =====================================================================
# 六、需求第 2 条：C9 一月工资的取数与累计填列
# =====================================================================
class TestRequirement2(unittest.TestCase):
    """需求第 2 条：C9 一月工资的取数口径 + 累计填列。

        ① iYPeriod = 界面所选年月的年月整数（选 2026年1月 → 202601）
        ② ccode = 660201
        ③ cdigest 必须包含「期间损益结转」
        ④ 取 mc（贷方金额）
        ⑤ 选 3 月时须同时填 C9、D9、E9（前面月份一并填列）

    【★ 为什么这组用例必须用真实配置（不打 _with_codes 补丁）★】
        本条是"端到端行为"：界面年月 → SQL 区间 → 逐月落列。
        用真实配置才能验证"工资行真的会这样填"。
        临时假科目只能验证框架逻辑，验不到实际报表的呈现。

    【★ 四个条件任一写错，SQL 都能正常执行并返回结果 ★】
          · 取 md 而非 mc   → 拿到的是费用发生额，金额偏大
          · 摘要写错措辞     → 返回空结果，报表全行空白但不报错
          · ccode 写错       → 取到别的费用项目
          · 区间上限写死 12  → 把未来月份的数也填进来
        报表上只会显示一个"看起来正常"的数，没有任何异常提示，
        所以必须把口径逐条锁死在 SQL 文本与落列行为上。
    """

    # 12 个月各给一个可区分的金额，任何落列错位都能立刻看出来
    ROWS_12 = [("660201", m, m * 100.0) for m in range(1, 13)]

    def _run(self, rows, month, year=2026):
        ctx = FakeCtx(rows, year=year, month=month)
        hook = FakeHook()
        rep = LK26Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)
        return hook, ctx

    # ---- 口径校验：直接查 SQL 文本 ----
    def test_sql_uses_credit_side(self):
        """SQL 必须 SUM(mc)，不能是 SUM(md)。

        【为什么直接查 SQL 文本？】
            假数据里只有一个 amount 值，报表层看不出方向。
            只有检查 SQL 才能真正确认取的是哪一列。
        """
        _, ctx = self._run([], month=1)
        sql = ctx.fetch_calls[0]["sql"]
        flat = sql.replace(" ", "")
        self.assertIn("SUM(mc)", flat, f"应取贷方 mc，实际 SQL：{sql}")
        self.assertNotIn("SUM(md)", flat, f"不应取借方 md，实际 SQL：{sql}")

    def test_sql_filters_digest_period_profit_settle(self):
        """SQL 必须带「期间损益结转」摘要筛选（不是「结转管理费用」）。"""
        _, ctx = self._run([], month=1)
        sql = ctx.fetch_calls[0]["sql"]
        params = ctx.fetch_calls[0]["params"]

        self.assertIn("cdigest LIKE", sql, f"缺少摘要筛选：{sql}")
        self.assertIn("%期间损益结转%", params,
                      f"摘要参数应含 %期间损益结转%，实际 {params}")

        # 【反向断言】防止有人照抄 LK25 的措辞
        self.assertNotIn("%结转管理费用%", params,
                         "不应使用 LK25 的『结转管理费用』措辞 —— "
                         "本表口径是『期间损益结转』")
        self.assertNotIn("%结转制造费用%", params,
                         "不应使用 LK25 制造费用的摘要措辞")

    def test_sql_filters_exactly_configured_codes(self):
        """SQL 的科目参数必须正好等于"已配置行涉及的科目集合"。

        【为什么不写死一份科目清单？】
            写死清单在补行时必然要同步改两处（配置 + 测试）——
            漏改就会误报。这里改为从 ROW_SPECS 反算期望值：
            "配置里有什么，SQL 里就该正好有什么"，永不失配。

        【为什么断言"正好"而不是"包含"？】
            若有人把某个科目配到了别的行（或多配了一个需求外的科目），
            "包含"式断言仍然通过 —— 报表会多取一笔钱，
            而每格看起来都正常。断言"正好"才能守住配置与需求一致。
        """
        expected = sorted({
            c
            for spec in ROW_SPECS
            if spec.codes
            for c in spec.codes
        })
        _, ctx = self._run([], month=1)
        params = ctx.fetch_calls[0]["params"]
        codes = sorted(p for p in params if isinstance(p, str) and p.isdigit())
        self.assertEqual(codes, expected,
                         f"科目参数应正好是已配置行的科目集合\n"
                         f"  期望：{expected}\n"
                         f"  实际：{codes}")

    def test_all_rows_data_lands_in_right_rows(self):
        """端到端：11 行各月数据落到正确的格位，互不串位。

        【为什么这条最关键？】
            11 行口径完全相同、科目编码不连续、还有一行是双科目合计。
            配置全对但 fetch 的分流写错（例如"科目→行"映射写成单值），
            就会让某些行的数互相覆盖或叠加 ——
            而报表上每格都有数字，看不出异常。

            本条给每行一个**可区分的基数**，任何串位都会立刻暴露。
        """
        # 【为什么用 enumerate 自动生成基数，而不是手写 30 个数字？】
        #   手写时每补一行就要加一个数字，漏加会 KeyError、重复会让
        #   两行数字相同而掩盖串位 —— 两种都很烦。
        #   这里按科目排序后依次分配 11.0, 22.0, 33.0… 保证：
        #     · 每个科目基数唯一（串位必然暴露）；
        #     · 补行时无需改动本处。
        all_codes = sorted({
            c for spec in ROW_SPECS if spec.codes for c in spec.codes
        })
        bases = {c: (i + 1) * 11.0 for i, c in enumerate(all_codes)}
        rows = []
        for m in (1, 2, 3):
            for code, base in bases.items():
                rows.append((code, m, base * m))

        hook, _ = self._run(rows, month=3)

        for spec in ROW_SPECS:
            if not spec.codes:
                continue
            expected = sum(bases[c] for c in spec.codes)   # 双科目行会相加
            for col, m in (("C", 1), ("D", 2), ("E", 3)):
                cell = f"{col}{spec.excel_row}"
                self.assertTrue(
                    approx(hook.get(cell), expected * m),
                    f"{spec.key} 的 {cell}（第{m}月）应为 {expected * m}，"
                    f"实际 {hook.get(cell)}")
            for col in "FGHIJKLMN":
                cell = f"{col}{spec.excel_row}"
                self.assertIsNone(hook.get(cell),
                                  f"{spec.key} 的 {cell} 属 4 月及以后，"
                                  f"3 月报表中不应写入")

    def test_first_four_rows_in_single_query(self):
        """四行科目应合并进**同一次**查询，而不是各查一次。

        【为什么重要？】
            四行口径完全相同，合并查询后：
              · 只需 1 次数据库往返（而不是 4 次）；
              · 四行数据来自同一快照，不会出现"两行取自不同时刻"。
            若有人图省事改成"每行各查一次"，本条会失败。
        """
        _, ctx = self._run([], month=1)
        self.assertEqual(len(ctx.fetch_calls), 1,
                         f"四行应合并为 1 次查询，实际 "
                         f"{len(ctx.fetch_calls)} 次")

    def test_sql_period_is_one_to_selected_month(self):
        """iYPeriod 区间必须是 202601~所选月份。"""
        for month in (1, 2, 3, 8, 12):
            with self.subTest(month=month):
                _, ctx = self._run([], month=month)
                params = ctx.fetch_calls[0]["params"]
                self.assertEqual(params[0], 2026,
                                 f"年度参数应为 2026，实际 {params}")
                self.assertEqual(params[1], 202601,
                                 f"区间起始应固定为 202601，实际 {params}")
                self.assertEqual(params[2], int(f"2026{month:02d}"),
                                 f"区间截止应为 2026{month:02d}，"
                                 f"实际 {params}")

    # ---- 落列行为 ----
    def test_january_fills_only_c9(self):
        """选 2026年1月：只填 C9，D9 及以后全部空白。"""
        hook, _ = self._run(self.ROWS_12, month=1)

        self.assertTrue(approx(hook.get("C9"), 100.0),
                        f"C9 应为 100，实际 {hook.get('C9')}")
        for col in "DEFGHIJKLMN":
            self.assertIsNone(hook.get(f"{col}9"),
                              f"{col}9 在 1 月报表中不应被写入，"
                              f"实际 {hook.get(f'{col}9')}")

    def test_march_fills_three_columns(self):
        """选 2026年3月：C9=100、D9=200、E9=300 —— 前三个月一并填列。

        【这是需求第 2 条的核心】财务拿到"3月表"时，要看到的是
        1~3 月的完整情况，而不是"3月单月"——单月数据无法判断趋势。
        """
        hook, _ = self._run(self.ROWS_12, month=3)

        self.assertTrue(approx(hook.get("C9"), 100.0),
                        f"C9（1月）应为 100，实际 {hook.get('C9')}")
        self.assertTrue(approx(hook.get("D9"), 200.0),
                        f"D9（2月）应为 200，实际 {hook.get('D9')}")
        self.assertTrue(approx(hook.get("E9"), 300.0),
                        f"E9（3月）应为 300，实际 {hook.get('E9')}")

    def test_march_does_not_fill_future_months(self):
        """选 2026年3月：F9~N9（4~12 月）必须全部空白。

        【为什么这条与上一条同样重要？】
            "3月表"里若混入 4 月的数，这张表就不再是 3 月表了，
            而看表的人无从分辨哪几列属于本期。
            只测"该填的填了"会漏掉这个方向相反的错。
        """
        hook, _ = self._run(self.ROWS_12, month=3)
        for col in "FGHIJKLMN":
            self.assertIsNone(hook.get(f"{col}9"),
                              f"{col}9 属 4 月及以后，3 月报表中不应写入，"
                              f"实际 {hook.get(f'{col}9')}")

    def test_august_fills_eight_columns(self):
        """选 2026年8月：C9~J9 全部填列（1~8 月共 8 列），K9 起空白。"""
        hook, _ = self._run(self.ROWS_12, month=8)

        for i, col in enumerate("CDEFGHIJ", start=1):
            self.assertTrue(
                approx(hook.get(f"{col}9"), i * 100.0),
                f"选 8 月时 {col}9（第{i}月）应为 {i * 100}，"
                f"实际 {hook.get(f'{col}9')}")
        for col in "KLMN":
            self.assertIsNone(hook.get(f"{col}9"),
                              f"{col}9 属 9 月及以后，8 月报表中不应写入")

    def test_december_fills_all_twelve_columns(self):
        """选 2026年12月：C9~N9 十二列全填。"""
        hook, _ = self._run(self.ROWS_12, month=12)
        for i, col in enumerate("CDEFGHIJKLMN", start=1):
            self.assertTrue(
                approx(hook.get(f"{col}9"), i * 100.0),
                f"选 12 月时 {col}9（第{i}月）应为 {i * 100}，"
                f"实际 {hook.get(f'{col}9')}")

    def test_gap_month_stays_blank_but_others_filled(self):
        """中间某月缺数据时：该月空白，其余月份照常填。

        【场景】选 5 月，但 3 月没有「期间损益结转」凭证
                （例如漏做结转）。
        【期望】C9/D9/F9/G9 有数，E9 空白 —— 不能因为缺一个月
                就把整行留空，也不能给缺的那月补 0。
        """
        rows = [("660201", m, m * 100.0) for m in (1, 2, 4, 5)]
        hook, _ = self._run(rows, month=5)

        self.assertTrue(approx(hook.get("C9"), 100.0), hook.get("C9"))
        self.assertTrue(approx(hook.get("D9"), 200.0), hook.get("D9"))
        self.assertIsNone(hook.get("E9"),
                          f"3 月无数据应保持空白，实际 {hook.get('E9')}")
        self.assertTrue(approx(hook.get("F9"), 400.0), hook.get("F9"))
        self.assertTrue(approx(hook.get("G9"), 500.0), hook.get("G9"))

    def test_no_data_no_write(self):
        """该科目当期无结转记录时，C9 保持空白（不写 0）。

        【为什么不写 0？】写 0 会让财务以为"本月工资结转额为 0"，
        而实际可能是"忘记做结转凭证"。留空 + 日志才能让人去查。
        """
        hook, _ = self._run([], month=1)
        self.assertIsNone(hook.get("C9"),
                          f"C9 不应被写入，实际 {hook.get('C9')}")

    def test_single_query_regardless_of_month(self):
        """无论选几个月，都只发一次查询（用区间一次取回，不逐月查）。"""
        for month in (1, 6, 12):
            with self.subTest(month=month):
                _, ctx = self._run(self.ROWS_12, month=month)
                self.assertEqual(
                    len(ctx.fetch_calls), 1,
                    f"选 {month} 月时应只发 1 次查询，"
                    f"实际 {len(ctx.fetch_calls)} 次")

    def test_salary_not_written_to_other_rows(self):
        """660201 的数据只能进第 9 行，不能串到其他费用项目。

        【注意：只检查**仍未配置**的行】
            第 10/11/12 行（职工福利费、办公费、差旅费）已在需求
            3~5 条中配置，第 13~20 行已在需求 6~12 条中配置 ——
            它们有数据时**就该写**，不能当作"不该有数"的行来断言。

            所以这里改为从 ROW_SPECS 动态取出"确实没配置的行"，
            补一行需求就自动少检查一行，无需改本用例。
        """
        hook, _ = self._run(self.ROWS_12, month=3)
        unconfigured_rows = [
            spec.excel_row for spec in ROW_SPECS if not spec.codes
        ]
        self.assertTrue(unconfigured_rows, "当前应有未配置的行")
        for row in unconfigured_rows:
            self.assertIsNone(
                hook.get(f"C{row}"),
                f"C{row} 不应被写入，实际 {hook.get(f'C{row}')}")
        # 【第 44 行不在此列】它现在是"修正后的公式"（需求第 32 条），
        #   不是数值，也不是空白。

    def test_first_four_rows_not_crossed(self):
        """需求 2~5 条：前四行各取各的科目，互不串用、也不合并。

        【为什么这条必须用"金额与科目末位错开"的造数？】
            四行口径完全相同、科目编码连续、格位连续（C9~C12）。
            若某行的科目写错成邻行的，两行都会显示"一个正常的数" ——
            只有给每行一个可区分的金额，错位才会立刻暴露。

            这里用：660201→101、660202→202、660203→303、660204→404。
            任何一处科目或格位串位，断言会直接指出是哪个格子。

            【月份说明】数据造在 1 月、报表也选 1 月，
            这样"该行该月的数"就落在 C 列（1月），断言最直观。
        """
        rows = [
            ("660201", 1, 101.0),
            ("660202", 1, 202.0),
            ("660203", 1, 303.0),
            ("660204", 1, 404.0),
        ]
        hook, _ = self._run(rows, month=1)

        expected = {"C9": 101.0, "C10": 202.0, "C11": 303.0, "C12": 404.0}
        for cell, value in expected.items():
            self.assertTrue(
                approx(hook.get(cell), value),
                f"{cell} 应为 {value}，实际 {hook.get(cell)}")

        # 反向确认没有被合并：任何一格都不应等于四者之和
        total = sum(expected.values())
        for cell in expected:
            self.assertFalse(
                approx(hook.get(cell), total),
                f"{cell} 不应是四行合计 {total}（说明科目被合并了）")

    def test_first_four_rows_accumulate_independently(self):
        """前四行各自累计 1~3 月（选 3 月时每行都填三列）。"""
        rows = []
        for m in (1, 2, 3):
            rows += [
                ("660201", m, m * 100.0),
                ("660202", m, m * 200.0),
                ("660203", m, m * 300.0),
                ("660204", m, m * 400.0),
            ]
        hook, _ = self._run(rows, month=3)

        for i, (code_mult, row) in enumerate(
                ((100.0, 9), (200.0, 10), (300.0, 11), (400.0, 12))):
            for col, m in (("C", 1), ("D", 2), ("E", 3)):
                expect = m * code_mult
                self.assertTrue(
                    approx(hook.get(f"{col}{row}"), expect),
                    f"{col}{row}（第{m}月）应为 {expect}，"
                    f"实际 {hook.get(f'{col}{row}')}")
            # 4 月及以后不得写入
            for col in "FGHIJKLMN":
                self.assertIsNone(
                    hook.get(f"{col}{row}"),
                    f"{col}{row} 属 4 月及以后，3 月报表中不应写入")

    def test_period_and_header_still_written(self):
        """有数据时，C4 年月与 A2/A4 页眉标签同样照常写入。"""
        hook, _ = self._run(self.ROWS_12, month=3, year=2026)

        self.assertEqual(hook.get(PERIOD_CELL), "2026年3月")
        for cell, text in HEADER_LABEL_CELLS:
            self.assertEqual(hook.get(cell), text,
                             f"{cell} 应为 {text!r}，"
                             f"实际 {hook.get(cell)!r}")

    def test_year_total_and_total_row_not_written_with_data(self):
        """有数据时，O 列与第 44 行合计仍不得写入（保留模板公式）。"""
        hook, _ = self._run(self.ROWS_12, month=12)

        self.assertIsNone(hook.get(f"{COL_YEAR_TOTAL}9"),
                          f"O9 不应被写入，"
                          f"实际 {hook.get(f'{COL_YEAR_TOTAL}9')}")

        # 【第 44 行：写公式，不写值】需求第 32 条。
        # 有数据时公式会算出非零值 —— 若误用 set 写值，
        # 报表上的数字一模一样，只有检查"公式是否还在"才能发现。
        for col in COL_OF_MONTH.values():
            cell = f"{col}{TOTAL_ROW}"
            formula = hook.formulas.get(cell)
            self.assertIsNotNone(
                formula, f"{cell} 应写入修正后的合计公式")
            self.assertEqual(
                formula, f"=SUM({col}9:{col}43)",
                f"{cell} 的公式应为 =SUM({col}9:{col}43)，实际 {formula!r}")
            kinds = [e[0] for e in hook.events if e[1] == cell]
            self.assertNotIn("set", kinds,
                             f"{cell} 不应使用普通 set 写值（会覆盖公式）")


# =====================================================================
# 七、暂不处理行（现场确认"口径不明"，20261112）
# =====================================================================
class TestPendingRows(unittest.TestCase):
    """六行暂不处理：即使给了数据，也不得写入任何金额。

        C16 商业保险费    C32 人力招聘费    C33 员工辞退福利
        C36 律师、诉讼费  C37 协会会费      C39 用工劳务费

    【为什么要专门测"即使有数据也不写"？】
        配置正确（pending=True、codes 为空）不代表行为正确 ——
        若某处逻辑误写了这几格（例如复制粘贴某行配置时漏改格位），
        报表就会多出数，而公式、合计看起来都正常。
        本条给足数据跑一遍，这六行的月列必须全为 None。
    """

    # 六行的 (项目名, Excel 行号)
    PENDING_CELLS = (
        ("商业保险费", 16),
        ("人力招聘费", 32),
        ("员工辞退福利", 33),
        ("律师、诉讼费", 36),
        ("协会会费", 37),
        ("用工劳务费", 39),
    )

    def test_pending_rows_not_written_even_with_data(self):
        """给前四行充足数据时，六行的月列仍必须全部空白。"""
        rows = [("660201", m, m * 100.0) for m in range(1, 13)]
        ctx = FakeCtx(rows, year=2026, month=12)
        hook = FakeHook()
        rep = LK26Report()
        rep.fetch(ctx)
        rep.write(hook, ctx)

        for name, excel_row in self.PENDING_CELLS:
            for col in COL_OF_MONTH.values():
                cell = f"{col}{excel_row}"
                self.assertIsNone(
                    hook.get(cell),
                    f"{name}（第{excel_row}行）的 {cell} 已确认暂不处理，"
                    f"不应被写入，实际 {hook.get(cell)}")

    def test_pending_rows_status_is_pending_not_todo(self):
        """六行的 status 必须是"暂不处理"，不能退化成"待配置"。

        【为什么这条重要？】
            两者在报表上都表现为空白，但含义完全相反：
              · 暂不处理 → 已与财务确认，不必再追；
              · 待配置   → 需求还没给到，属待办事项。
            若 pending 标志被误删，维护者会一直去追这六条
            早已确认不做的需求 —— 而代码上没有任何报错。
        """
        for name, _excel_row in self.PENDING_CELLS:
            with self.subTest(name=name):
                spec = next(s for s in ROW_SPECS if s.key == name)
                self.assertTrue(spec.pending,
                                f"{name} 应标记 pending=True")
                self.assertEqual(spec.codes, (),
                                 f"{name} 暂不处理，不应配科目，"
                                 f"实际 {spec.codes}")
                self.assertEqual(spec.status, "暂不处理",
                                 f"{name} 的 status 应为'暂不处理'，"
                                 f"实际 {spec.status!r}")

    def test_pending_rows_not_in_query_codes(self):
        """暂不处理的六行不参与取数 —— 它们不应贡献任何科目到 SQL。

        【为什么不写死一份期望清单？】
            写死清单在补行时必然要同步改两处（配置 + 测试），
            漏改就会误报。这里改为：
              ① 期望值从"已配置行"反算；
              ② 断言 SQL 科目正好等于该期望值。
            这样"暂不处理的行没参与取数"就自动被覆盖 ——
            若某天有人给 pending 行误填了 codes，期望值会变，
            而 SQL 也会跟着变，这两者仍相等 ——
            所以另有一条 test_configured_rows_match_requirements
            专门守住"pending 行不得有 codes"。
        """
        rows = [("660201", m, m * 100.0) for m in range(1, 13)]
        ctx = FakeCtx(rows, year=2026, month=12)
        LK26Report().fetch(ctx)

        expected = sorted({
            c for spec in ROW_SPECS if spec.codes for c in spec.codes
        })
        codes = sorted(p for p in ctx.fetch_calls[0]["params"]
                       if isinstance(p, str) and p.isdigit())
        self.assertEqual(codes, expected,
                         f"SQL 科目应等于已配置行的科目集合\n"
                         f"  期望：{expected}\n"
                         f"  实际：{codes}")

        # 六行暂不处理的行都不应贡献科目（它们 codes 为空）
        for spec in ROW_SPECS:
            if not spec.pending:
                continue
            for c in spec.codes:
                self.assertNotIn(
                    c, codes,
                    f"暂不处理行「{spec.key}」不应贡献科目 {c} 到 SQL")
