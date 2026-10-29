# -*- coding: utf-8 -*-
"""
LK01 资产负债表
============================================================

【模块职责】
    按模板 LK01 的版式，从 U8 年度库取数并填列。
    同时作为 LK02~LK33 的编写范例。

【当前已实现的行项目】
    ┌──────┬────────────────────┬────────────────┬─────────┬─────────┐
    │ 行次 │ 项目               │ 汇总科目       │ 期末格  │ 年初格  │
    ├──────┼────────────────────┼────────────────┼─────────┼─────────┤
    │  1   │ 货币资金           │ 1001 + 1002    │  C11    │  D11    │
    │  2   │ 交易性金融资产     │ 1012           │  C13    │  D13    │
    │  3   │ 应收票据           │ 1121           │  C15    │  D15    │
    │  4   │ 应收账款(净额)     │ 1122-123101    │  C17    │  D17    │
    │  6   │ 其他应收款(净额)   │ 1221-123102    │  C21    │  D21    │
    │  7   │ 存货               │ 1401+1403      │  C23    │  D23    │
    │      │                    │ +1405+1408     │         │         │
    │  9   │ 其他流动资产       │ 1131（应收股利）│  C31    │  D31    │
    │ 16   │ 长期应收款         │ 1531           │  C41    │  D41    │
    │ 17   │ 长期股权投资       │ 1511           │  C43    │  D43    │
    │ 19   │ 固定资产(净额)     │ 1601-1602-1603 │  C47    │  D47    │
    │ 22   │ 使用权资产         │ 1462           │  C53    │  D53    │
    └──────┴────────────────────┴────────────────┴─────────┴─────────┘

    另：C4 写会计期间（需求 20261024 第 1 条）。

【待定的行项目（口径未明确，暂不实现）】
    以下 6 项已与财务确认"暂不处理、留待后期"，不是遗漏：
        预付款项（C19/D19，行次5）、合同资产（C25/D25，行次8）、
        持有待售资产（C27/D27，行次9）、
        一年内到期的非流动资产（C29/D29，行次10）、
        债权投资（C37/D37，行次14）、其他债权投资（C39/D39，行次15）

    详见下方 PENDING_ROWS 常量。补实现时把它移入 ROW_SPECS 即可。

【★ 新增一行项目怎么做？（照抄即可）★】
    在下面 ROW_SPECS 里加一条 BalanceRow 即可，**不需要改 fetch/write**。
    需要提供：行标识、显示名、科目集合、期末格、年初格。
    随后把这两个格子登记进 cell_map（供勾稽核对读回）。

【取数口径的关键约定（现场已确认，改动前务必确认）】
    1. GL_accsum 是"每月一行"的科目汇总表，带 iperiod（月份）。
    2. mb（期初）字段语义**随月份变化**：
         · iperiod = 1 时，mb 是"年初余额"
         · iperiod ≠ 1 时，mb 是"本月期初"（即上月末余额）
       所以"上年年末余额"只能取**本年 1 月记录的 mb**。
    3. 期末余额统一取**本期记录的 me**。
"""
from __future__ import annotations

from dataclasses import dataclass, field

from core.context import ReportContext
from core.excel_engine import _SheetHook
from core.logging_setup import get_logger
from queries import u8
from reports.base import BaseReport
from reports.header import fill_common_header
from reports.registry import register

logger = get_logger(__name__)


@dataclass
class BalanceRow:
    """资产负债表里"一行指标"的配置。

    【为什么用配置而不是逐个写 if？】
        本表 30 多行，取数逻辑完全一致（某几个科目的 me/mb 汇总），
        差异只在"科目集合"和"写哪两个格子"。
        抽成配置后：
          · 新增一行 = 加一条配置，不动任何逻辑代码；
          · 口径集中在一处，便于与财务逐行核对；
          · 不会因为复制粘贴而漏改某个格子地址。

    字段：
        key          内部标识（用于日志与调试）
        label        财务名称（写日志用，建议与模板行项目名一致）
        codes        参与汇总的科目编码（**加项**，精确匹配）
        end_cell     期末余额单元格（如 "C11"）
        begin_cell   上年年末余额单元格（如 "D11"）
        exclude_codes **减项**科目编码（默认空）

        【注意】codes 是"精确匹配"。若某行需要包含明细科目
        （如应收账款 1122 及其下级 112201…），应改用前缀匹配的
        取数函数，不能在这里直接写 "1122"。

        【exclude_codes 是为什么加的？】
            资产负债表里有"净额列示"的行，最典型的是应收账款：
                应收账款(净) = 应收账款余额 - 坏账准备
            即需求第 6 条调整后的口径：1122 - 123101。
            若只提供加法，这类行就没法表达，只能退回到"给这张表
            写专用逻辑"——那会破坏"新增一行只加配置"的原则。

            所以这里加一个减项集合：
                最终值 = sum(codes) - sum(exclude_codes)
            默认空元组，对前面 7 行没有任何影响。
    """

    key: str
    label: str
    codes: tuple[str, ...]
    end_cell: str
    begin_cell: str
    exclude_codes: tuple[str, ...] = ()


# =====================================================================
# 行项目配置表 —— 新增行项目只改这里
# =====================================================================
ROW_SPECS: tuple[BalanceRow, ...] = (
    BalanceRow(
        key="货币资金",
        label="货币资金",
        # 【口径变更 20261028】1012 已移出本行，只算 1001+1002
        codes=("1001", "1002"),
        end_cell="C11",
        begin_cell="D11",
    ),
    BalanceRow(
        key="交易性金融资产",
        label="交易性金融资产（其他货币资金 1012）",
        # 【需求 20261028 第 4 条】1012「其他货币资金」归入本行核算
        codes=("1012",),
        end_cell="C13",
        begin_cell="D13",
    ),
    BalanceRow(
        key="应收票据",
        label="应收票据",
        # 【需求 20261028 第 5 条】科目 1121
        # 模板第 15 行 A15="  应收票据"，行次 3
        codes=("1121",),
        end_cell="C15",
        begin_cell="D15",
    ),
    BalanceRow(
        key="应收账款",
        label="应收账款（净额，已扣坏账准备）",
        # 模板第 17 行 A17="  应收账款"，行次 4
        #
        # 【需求 20261028 第 6 条 → 20261029 调整】
        #   原口径：只取 1122 本身。
        #   调整后：应收账款按**净额**列示，需扣除坏账准备：
        #       C17 = 1122 月末数 - 123101 月末数
        #       D17 = 1122 年初数 - 123101 年初数
        #
        # 【为什么 label 要写"净额，已扣坏账准备"？】
        #   报表上只显示一个合计数。日后若核对发现 C17 与用友前台
        #   "应收账款"余额不一致，看到这个 label 就能立刻想到
        #   "差的那部分是不是坏账准备"，不必翻代码。
        #
        # 【口径】仍为精确匹配（不含 112201 等下级）。
        codes=("1122",),
        exclude_codes=("123101",),
        end_cell="C17",
        begin_cell="D17",
    ),
    BalanceRow(
        key="其他应收款",
        label="其他应收款（净额，已扣坏账准备）",
        # 模板第 21 行 A21="  其他应收款"，行次 6
        #
        # 【需求 20261028 第 8 条 → 20261029 调整】
        #   原口径：只取 1221 本身。
        #   调整后：按**净额**列示，扣除其他应收款的坏账准备：
        #       C21 = 1221 月末数 - 123102 月末数
        #       D21 = 1221 年初数 - 123102 年初数
        #   与应收账款(1122-123101)完全相同的计算方式。
        #
        # 【注意减项科目不同】应收账款的坏账准备是 123101，
        #   其他应收款的是 123102，切勿混用。
        #
        # 【注意】格位出现跳格：应收账款在 C17，本行在 C21，
        #   中间第 18~20 行是预付款项(及其英文行)。
        #   依旧按需求指定地址写入，不推算。
        codes=("1221",),
        exclude_codes=("123102",),
        end_cell="C21",
        begin_cell="D21",
    ),
    BalanceRow(
        key="存货",
        label="存货",
        # 【需求 20261029 第 9 条】由四个科目组成：
        #   1401 材料采购、1403 原材料、1405 库存商品、1408 委托加工物资
        # 模板第 23 行 A23="  存货"，行次 7
        # 【口径】与货币资金(C11/D11)一致，即"多科目汇总"。
        #   这里再确认一次四者的对应关系，避免科目编码与名称错位：
        #     1401 → 材料采购
        #     1403 → 原材料
        #     1405 → 库存商品
        #     1408 → 委托加工物资
        codes=("1401", "1403", "1405", "1408"),
        end_cell="C23",
        begin_cell="D23",
    ),
    BalanceRow(
        key="应收股利",
        # 【标签为何这样写？】
        #   模板第 31 行的项目名是"其他流动资产"，全表并无"应收股利"行。
        #   现场确认：应收股利(1131) 并入该行列示。
        #   所以日志/核对指标里两者都标出来 —— 日后若要追查
        #   "C31 这个数由什么构成"，一眼能看出它只含 1131，
        #   不会被误认为"其他流动资产的全部内容"。
        label="其他流动资产（含应收股利 1131）",
        # 【需求 20261028 第 7 条】科目 1131，C31=月末数、D31=年初数
        # 【注意】本行格位与前面几行不连续（C17 之后跳到 C31），
        #   按需求明确指定的地址写入，不要"顺着上一行推算"。
        codes=("1131",),
        end_cell="C31",
        begin_cell="D31",
    ),
    BalanceRow(
        key="长期应收款",
        label="长期应收款",
        # 【需求 20261029 第 12 条】科目 1531
        # 模板第 41 行 A41="  长期应收款"，行次 16
        #
        # 【注意格位顺序】本行 C41 在长期股权投资 C43 之前，
        #   符合模板中"长期应收款(第41行) → 长期股权投资(第43行)"的顺序。
        codes=("1531",),
        end_cell="C41",
        begin_cell="D41",
    ),
    BalanceRow(
        key="长期股权投资",
        label="长期股权投资",
        # 【需求 20261029 第 11 条】科目 1511
        # 模板第 43 行 A43="  长期股权投资"，行次 17
        #
        # 【注意：本行后续可能需要改为净额】
        #   长期股权投资通常要扣减值准备。本次需求只要求取 1511，
        #   故暂不设 exclude_codes。若日后确认需扣减值准备，
        #   加一个 exclude_codes=("1512",) 即可，无需改逻辑。
        codes=("1511",),
        end_cell="C43",
        begin_cell="D43",
    ),
    BalanceRow(
        key="固定资产",
        label="固定资产（净额，已扣累计折旧及减值准备）",
        # 【需求 20261029 第 13 条 + 补充确认】科目 1601
        # 模板第 47 行 A47="  固定资产"，行次 19
        #
        # 【口径】净额列示，需扣除两项：
        #       C47 = 1601 月末数 - 1602 累计折旧 - 1603 减值准备
        #       D47 = 1601 年初数 - 1602 累计折旧 - 1603 减值准备
        #
        # 【重要：这是本表第一个"多项减项"的行】
        #   exclude_codes 支持多个科目，求和后一并扣除。
        #   报告上只显示一个净额，若日后核对发现与用友前台"固定资产"
        #   原值不一致，需意识到这里已经扣过两项。
        codes=("1601",),
        exclude_codes=("1602", "1603"),
        end_cell="C47",
        begin_cell="D47",
    ),
    BalanceRow(
        key="使用权资产",
        label="使用权资产",
        # 【需求 20261029 第 10 条】科目 1462
        # 模板第 53 行 A53="  使用权资产"，行次 22
        #
        # 【重要：曾差点写错行】
        #   需求最初写的是"C25 合同资产填列使用权资产 1462"，但模板里
        #   "合同资产"在第 25 行、"使用权资产"在第 53 行，是两个独立项目。
        #   经核对模板后已更正为 C53/D53。此处留记录，提醒后续：
        #   **需求里的项目名与单元格地址若不自洽，必须先查模板再动手。**
        codes=("1462",),
        end_cell="C53",
        begin_cell="D53",
    ),
)


# =====================================================================
# 待定行项目清单 —— 口径未明确，暂不实现
# =====================================================================
# 【为什么要把"不做的行"也记下来？】
#   有两种"没做"看起来一样、后果却不同：
#     ① 忘了做 → 日后没人知道漏了，报表静默缺项；
#     ② 有意不做 → 已与财务确认口径待定，等条件明确再补。
#   把 ② 显式列在这里，读代码的人一眼能分清，不必靠猜。
#
#   测试里有一条会检查"待定项不得与已实现项占用同一格子"，
#   防止将来补实现时把数值写进冲突位置。
#
# 【恢复条件】财务明确了科目口径后，把该项从本表移入 ROW_SPECS 即可。
PENDING_ROWS: tuple[tuple[str, str, int], ...] = (
    # (模板项目名, 格位, 行次)
    ("预付款项", "C19/D19", 5),
    ("合同资产", "C25/D25", 8),
    ("持有待售资产", "C27/D27", 9),
    ("一年内到期的非流动资产", "C29/D29", 10),
    ("债权投资", "C37/D37", 14),
    ("其他债权投资", "C39/D39", 15),
)


@register
class LK01Report(BaseReport):
    """资产负债表"""

    code = "LK01"
    name = "资产负债表"

    # 关键指标 → 单元格，供父类 verify() 读回做勾稽核对。
    # 【原则】只登记"确实写入了数值"的格子，不要登记表头文字格。
    cell_map = {
        "会计期间": "C4",
        "货币资金_期末": "C11",
        "货币资金_上年年末": "D11",
        "交易性金融资产_期末": "C13",
        "交易性金融资产_上年年末": "D13",
        "应收票据_期末": "C15",
        "应收票据_上年年末": "D15",
        "应收账款_期末": "C17",
        "应收账款_上年年末": "D17",
        "其他应收款_期末": "C21",
        "其他应收款_上年年末": "D21",
        "存货_期末": "C23",
        "存货_上年年末": "D23",
        "应收股利_期末": "C31",
        "应收股利_上年年末": "D31",
        "长期应收款_期末": "C41",
        "长期应收款_上年年末": "D41",
        "长期股权投资_期末": "C43",
        "长期股权投资_上年年末": "D43",
        "固定资产_期末": "C47",
        "固定资产_上年年末": "D47",
        "使用权资产_期末": "C53",
        "使用权资产_上年年末": "D53",
    }

    def __init__(self):
        super().__init__()
        # 每行的取数结果：{row_key: (期末值, 年初值)}
        self.row_values: dict[str, tuple[float, float]] = {}
        # 每行的取数状态：{row_key: (本期是否取到, 本年1月是否取到)}
        # 【为什么分两个标记？】两格来自不同期间，可能只有一格有数据
        # （如年中建账、缺 1 月记录），必须各自判断是否写入。
        self.row_has: dict[str, tuple[bool, bool]] = {}
        # 取数明细，供日志核对"这个数由哪几个科目构成"
        self.row_detail: dict[str, tuple[list[str], list[str]]] = {}

    # ------------------------------------------------------------------
    # 第1步：取数
    # ------------------------------------------------------------------
    def fetch(self, ctx: ReportContext) -> None:
        """一次取回所有行涉及的科目，再按行配置分流汇总。

        【为什么把所有科目合并成一次查询？】
            本表各行涉及的科目是分散的（货币资金 2 个、存货 4 个…），
            若每行各查一次，30 多行就是 30 多次数据库往返。
            合并成一次（对 ROW_SPECS 里所有科目去重），只需 1 次，
            且所有行看到的是同一份快照 —— 不会出现
            "两行数据来自不同时刻"的一致性问题。

            全部行做完后科目约几十个，一次 IN 查询即可，代价可忽略。
        """
        year, month = ctx.period.year, ctx.period.month

        # 汇总所有行用到的科目（**含减项**，去重 + 排序，保证缓存键稳定）
        # 【为什么要带上 exclude_codes？】
        #   减项科目（如坏账准备 123101）也必须查回来，否则减不掉。
        all_codes = tuple(sorted({
            c
            for spec in ROW_SPECS
            for c in (*spec.codes, *spec.exclude_codes)
        }))

        # 【缓存键】必须带上科目集合，否则将来加行项目时会撞键
        key = ctx.make_key("GL_accsum", "LK01资产行", *all_codes)

        sql, params = u8.gl_balance_by_codes(year, month, all_codes)

        try:
            rows = ctx.fetch(key, sql, params, report_code=self.code)
        except Exception as exc:
            # 单表取数失败不中断整体流程（43 张表里坏 1 张不牵连其他）
            logger.warning("[%s] 取数失败：%s", self.code, exc)
            rows = []

        # ---- 初始化每行的累加器 ----
        acc_end = {spec.key: 0.0 for spec in ROW_SPECS}
        acc_begin = {spec.key: 0.0 for spec in ROW_SPECS}
        has_end = {spec.key: False for spec in ROW_SPECS}
        has_begin = {spec.key: False for spec in ROW_SPECS}
        det_end: dict[str, list[str]] = {spec.key: [] for spec in ROW_SPECS}
        det_begin: dict[str, list[str]] = {spec.key: [] for spec in ROW_SPECS}

        # 科目 → [(所属行, 符号)]
        #   符号 +1 = 加项（codes），-1 = 减项（exclude_codes）
        #   一个科目可能同时被多行使用，故用 list。
        code_to_rows: dict[str, list[tuple[BalanceRow, int]]] = {}
        for spec in ROW_SPECS:
            for c in spec.codes:
                code_to_rows.setdefault(c, []).append((spec, +1))
            for c in spec.exclude_codes:
                code_to_rows.setdefault(c, []).append((spec, -1))

        # ---- 逐条分流累加 ----
        for row in rows:
            try:
                # 列顺序由 u8.gl_balance_by_codes 的 SELECT 决定：
                #   第0列 ccode、第1列 iperiod、第2列 mb、第3列 me
                ccode = str(row[0]).strip()
                iperiod = int(row[1])
                mb = float(row[2]) if row[2] is not None else 0.0
                me = float(row[3]) if row[3] is not None else 0.0
            except (IndexError, TypeError, ValueError) as exc:
                logger.warning("[%s] 跳过异常记录 %r：%s", self.code, row, exc)
                continue

            specs = code_to_rows.get(ccode, [])
            if not specs:
                continue

            for spec, sign in specs:
                # 【关键：has_xxx 只由"加项"决定】
                #   若某行只有减项科目有记录、加项科目没有（数据异常），
                #   累加结果会是负数。此时 has_end 仍为 False，
                #   该格不会被写入 —— 宁可不写，也不写一个负数误导人。
                if sign > 0:
                    mark = "+"
                else:
                    mark = "-"

                if iperiod == month:
                    # 本期记录 → 期末余额
                    acc_end[spec.key] += sign * me
                    if sign > 0:
                        has_end[spec.key] = True
                    det_end[spec.key].append(f"{mark}{ccode}={me:.2f}")

                if iperiod == 1:
                    # 本年1月记录 → 年初余额（上年年末）
                    acc_begin[spec.key] += sign * mb
                    if sign > 0:
                        has_begin[spec.key] = True
                    det_begin[spec.key].append(f"{mark}{ccode}={mb:.2f}")

                # 【本期恰为1月时，同一行会同时命中上面两个分支：
                #   me 作期末、mb 作年初，互不干扰，也不会重复计算。】

        # ---- 收口保存 ----
        for spec in ROW_SPECS:
            self.row_values[spec.key] = (acc_end[spec.key], acc_begin[spec.key])
            self.row_has[spec.key] = (has_end[spec.key], has_begin[spec.key])
            self.row_detail[spec.key] = (det_end[spec.key], det_begin[spec.key])

            end_v, begin_v = self.row_values[spec.key]
            e_ok, b_ok = self.row_has[spec.key]
            logger.info(
                "[%s] %s：期末(%s)=%.2f ← [%s] | 年初(%s)=%.2f ← [%s]",
                self.code, spec.label, spec.end_cell, end_v,
                "、".join(det_end[spec.key]) if e_ok else "未取到本期记录",
                spec.begin_cell, begin_v,
                "、".join(det_begin[spec.key]) if b_ok else "未取到本年1月记录",
            )

    # ------------------------------------------------------------------
    # 第2步：填表
    # ------------------------------------------------------------------
    def write(self, hook: _SheetHook, ctx: ReportContext) -> None:
        """按配置逐行写入，并填页眉与会计期间。"""
        # 1) 通用页眉（公司名称等）
        # 【为什么 fill_period=False？】
        #   通用页眉会把年月追加到"会计期间："标签格（C1）。
        #   但本表按需求要把年月写到 C4，若两处都写，
        #   同一页会出现两个年月，客户会以为是数据重复。
        fill_common_header(hook, ctx, fill_period=False)

        # 2) 会计期间 → C4（需求 20261024 第 1 条）
        # 【必须用 set_text，不能用 set】
        #   C4 是"金额格式"的格子。若用 set() 写 "2026年5月"，
        #   Excel 会把它识别成日期 2026-05-01 并存成序列号 46143，
        #   再按金额格式显示成 "46,143.00" —— 看起来像一笔莫名其妙的钱。
        #   set_text() 先设文本格式再写值，才能原样显示年月。
        hook.set_text(self.cell_map["会计期间"],
                      f"{ctx.period.year}年{ctx.period.month}月")

        # 3) 逐行写入金额
        # 【为什么"没取到数据就不写"？】
        #   数据库连不上时累加值仍是 0.0。若照写不误，
        #   C11 会显示 0.00，客户看到"货币资金为零"，
        #   会误认为公司账上真的没钱 —— 这个误读代价太大。
        #   所以约定：**没取到数据就保持空白**，让空白本身成为
        #   "未取数"的信号；原因已在 fetch() 写入日志。
        for spec in ROW_SPECS:
            end_v, begin_v = self.row_values.get(spec.key, (0.0, 0.0))
            e_ok, b_ok = self.row_has.get(spec.key, (False, False))

            if e_ok:
                hook.set(spec.end_cell, end_v)
                logger.info("[%s] 已写入 %s=%.2f（%s期末余额）",
                            self.code, spec.end_cell, end_v, spec.label)
            else:
                logger.warning("[%s] 未取得本期记录，%s 保持空白（不写 0）",
                               self.code, spec.end_cell)

            if b_ok:
                hook.set(spec.begin_cell, begin_v)
                logger.info("[%s] 已写入 %s=%.2f（%s上年年末余额）",
                            self.code, spec.begin_cell, begin_v, spec.label)
            else:
                logger.warning("[%s] 未取得本年1月记录，%s 保持空白（不写 0）",
                               self.code, spec.begin_cell)