# -*- coding: utf-8 -*-
"""
LK02 利润表
============================================================

【模块职责】
    按模板 LK02 的版式，从 U8 年度库取数并填列。
    与 LK01（资产负债表，取"月末余额"）不同，本表取的是**发生额**，
    且数据源是凭证分录表 GL_accvouch（不是科目汇总表 GL_accsum）——
    因为本表多行必须按"摘要 + 对方科目"筛选，只有分录级数据能表达。

【模板版式（已逐格核对 2026版新修订企业财务报模板.xls 的 LK02 页）】

    行 1  A1='力劲集团'
    行 2  A2='公司名称:'            J2='制表：'
    行 3  A3='币种：人民币'          J3='复核：'
    行 4  A4='会计期间：'            **C4 = 会计期间填写位（需求第 1 条）**
    行 6  A6='利       润       表'
    行 8  表头：A='项  目' B='行' C~N='1月份'~'12月份' O='本年累计数' P='上年同期累计数'
    行 9 起为各项目行（详见下方 EXCEL_ROW）

    【列位（A1 记法，1-based）】
        A=项目名  B=行次
        C=1月  D=2月  E=3月  F=4月  G=5月  H=6月
        I=7月  J=8月  K=9月  L=10月 M=11月 N=12月
        O=本年累计数   P=上年同期累计数

【模板行次 ↔ Excel 行号 对照（务必以此为准，不要按"序号顺序"推算）】

    行次   Excel行   项目
      1      9      一、营业收入
      2     10       其中：出口主营业务收入
      3     11              内销主营业务收入
      4     12              其他业务收入
      5     13       减：营业成本
      6     14         其中：出口主营业务成本
      7     15               内销主营业务成本
      8     16               其他业务成本
      9     17               税金及附加
     10     18               销售费用
     11     19               管理费用
     12     20               研发费用
     13     21               财务费用
     14     22            其中：利息费用
     15     23                  利息收入
     16     24       加：其他收益（损失以“-”号填列）
     17     25           投资收益（损失以“-”号填列）
     18     26           公允价值变动收益（损失以“-”号填列）
     19     27           信用减值损失（损失以“-”号填列）
     20     28           资产减值损失（损失以“-”号填列）
     21     29           资产处置收益（损失以“-”号填列）
     22     30       二、营业利润（亏损以“-”号填列）
     23     31       加：营业外收入
     24     32       减：营业外支出
     25     33       三、利润总额（亏损总额以“-”号填列）
     26     34       减：所得税费用
     27     35       四、净利润（净亏损以“-”号填列）

【★ 已确认的开发口径（勿擅自改动）★】

    ① **保留模板公式不动**（需求现场确认）
       模板里行次 22（营业利润）、25（利润总额）、27（净利润）在 C~O 列
       本来就带公式。这些格子**绝不能由程序写数**，否则会覆盖公式、
       且下个月再生成时公式就没了。程序只填"明细行"。

       同理，模板中任何"整行都有 0.0/公式"的格子，若后续确认属公式行，
       一律登记进 FORMULA_ROWS，不写入数值。

    ② 明细行填数范围：**1~12 月各月列（C~N）+ 本年累计（O）**
       · 某月无数据时该月格保持空白（不写 0）——与 LK01 的约定一致；
       · O 列 = 各"有数据的月份"之和（不写死 1~12 一定都有），
         因此 5 月生成时 O 自然等于 1~5 月合计，无需额外判断期间；
       · P 列（上年同期累计数）暂不填，待需求明确后处理。

    ③ 取数类型为**本期发生额**（不是余额），且方向按行声明：
       · `direction="debit"`  → 取 md（借方金额）
       · `direction="credit"` → 取 mc（贷方金额）

       【重要：收入行为什么取借方 md？】
           收入类科目平时记贷方；月末结转到本年利润时分录为
               「借：主营业务收入 xxx / 贷：本年利润 xxx」
           即结转金额出现在**借方**。所以"该月已结转的收入"
           = 结转凭证里该科目行的 md 合计。取 mc 会得到 0。

    ④ 数据源为 GL_accvouch（凭证分录），不是 GL_accsum（科目汇总）。
       原因见下方「取数路线」说明与 docs/GL_accvouch字段补充说明.md。

【取数路线：为什么选 GL_accvouch】

    | 路线 | 表            | 能否按摘要/对方科目筛选 | 速度 |
    |------|---------------|--------------------------|------|
    | A    | GL_accsum     | ❌ 不能                   | 快   |
    | B    | GL_accvouch   | ✅ 能                     | 较慢 |

    本表已确认的口径普遍需要 A 做不到的筛选：
        · 只认摘要含"结转损益"类的分录
        · 只认对方科目 = 4103（本年利润）
    因此统一走 B。取数函数见
    queries/u8.py 的 gl_profit_amount_by_period()。

【★ 开发进度（逐条追加）★】

    [x] 需求第 1 条：C4 写会计期间
        取值来源 = ctx.period（程序界面上录入的年月），**不是写死的常量**。
        选 2026年5月 → C4="2026年5月"；选 2026年8月 → C4="2026年8月"；
        选 2025年11月 → C4="2025年11月"。年、月两段都随选择变化。

    [x] 需求第 2 条：行次 2「其中：出口主营业务收入」→ C10 等 12 个月格 + O10
        口径（现场 20261102 确认）：
            ① GL_accvouch 中 ccode = '600102'
            ② iyear = 目标年度；iperiod = 该月
            ③ cdigest 含「损益结转」或「结转损益」或「自动结转损益」
            ④ ccode_equal（对方科目）= '4103'（本年利润）
            ⑤ 取 **md（借方金额）** 合计
        结果：1 月合计 → C10，2 月 → D10，……，12 月 → N10；
              本年累计（各有效月份之和） → O10。

    [ ] 需求后续条目：其余行项目待财务逐条给出

    每追加一行，只需在 ROW_SPECS 里加一条 ProfitRow，
    **不需要改 fetch/write**（详见 ProfitRow 的字段说明）。

【新增一行项目怎么做？（照抄即可）】
    在 ROW_SPECS 里加一条 ProfitRow，**不需要改 fetch/write**：

        ProfitRow(
            key="内销主营业务收入",
            label="其中：内销主营业务收入",
            codes=("600101",),
            direction="debit",
            excel_row=EXCEL_ROW[3],          # = 11
            digest_keywords=DIGEST_SETTLE,
            equal_codes=(EQUAL_PROFIT,),
        )

    cell_map 会自动生成（见 _make_cell_map），无需手工登记 13 个格子。
"""
from __future__ import annotations

from dataclasses import dataclass

from core.context import ReportContext
from core.excel_engine import _SheetHook
from core.logging_setup import get_logger
from queries import u8
from reports.base import BaseReport
from reports.header import fill_common_header
from reports.registry import register

logger = get_logger(__name__)


# =====================================================================
# 列位与行位常量 —— 所有坐标集中在此，避免散落在逻辑代码里
# =====================================================================

# 会计期间单元格（需求第 1 条明确指定）
#
# 【写入的值是动态的，不是常量】
#   本常量只固定"写在哪一格"，具体年月来自 ctx.period：
#       hook.set_text(PERIOD_CELL, f"{ctx.period.year}年{ctx.period.month}月")
#   界面上选 2026年8月 → C4="2026年8月"
#   界面上选 2025年11月 → C4="2025年11月"
#   切勿把年月当成代码里的固定值 —— 那会导致每月生成都是同一个月。
PERIOD_CELL = "C4"

# 月份 → 列字母（1-based 的列序：C=3 … N=14）
COL_OF_MONTH: dict[int, str] = {
    1: "C", 2: "D", 3: "E", 4: "F", 5: "G", 6: "H",
    7: "I", 8: "J", 9: "K", 10: "L", 11: "M", 12: "N",
}

# 本年累计数 / 上年同期累计数
COL_YEAR_TOTAL = "O"
COL_LAST_YEAR_TOTAL = "P"

# 模板行次 → Excel 行号（**已逐格核对模板**，不要推算）
EXCEL_ROW: dict[int, int] = {
    1: 9, 2: 10, 3: 11, 4: 12,
    5: 13, 6: 14, 7: 15, 8: 16, 9: 17, 10: 18,
    11: 19, 12: 20, 13: 21, 14: 22, 15: 23,
    16: 24, 17: 25, 18: 26, 19: 27, 20: 28, 21: 29,
    22: 30, 23: 31, 24: 32,
    25: 33, 26: 34, 27: 35,
}

# 模板自带公式的汇总行 —— 程序绝不写入这些行（只保留模板公式）
#   【为什么用行次而不是 Excel 行号？】
#     行次是模板的业务编号（“行次 22 = 营业利润”），不随插删行变化，
#     日后若模板调整行位置，这里仍然可读、可核对。
FORMULA_ROWS: tuple[int, ...] = (22, 25, 27)


# =====================================================================
# 取数口径常量（现场确认的语义，集中定义避免各处拼写不一致）
# =====================================================================

# 「结转损益」类摘要关键字。
#
# 【为什么只列两种写法，却能覆盖三种？】
#   需求给的三种写法是「损益结转」「结转损益」「自动结转损益」。
#   后两者都以「结转损益」结尾 —— 用 LIKE '%结转损益%' 就能同时命中，
#   所以不必再单列「自动结转损益」。
#   「损益结转」是倒装写法，是独立的子串，必须单独列。
#
# 【为什么用"包含"（LIKE %…%）而不是"等于"？】
#   实务摘要常常是「月末自动结转损益」「12月结转损益-出口」这类带前缀/后缀的
#   完整句子，要求完全相等会大量漏取。项目里其他关键字定位也一律用"包含"。
DIGEST_SETTLE: tuple[str, ...] = ("损益结转", "结转损益")

# 本年利润科目（结转损益的对方科目）
#   【为什么必须限定它？】
#       同一科目也可能因冲销、调账、红字冲回产生分录，
#       只看摘要会被误纳。加上对方科目=4103 后，
#       只认"确实结转到本年利润"的那一条，口径才干净。
EQUAL_PROFIT = "4103"


# =====================================================================
# 行项目配置
# =====================================================================
@dataclass
class ProfitRow:
    """利润表里"一行指标"的配置。

    【字段说明】
        key         内部标识（日志与调试用）
        label       财务名称（写日志用，建议与模板 A 列一致）
        codes       参与汇总的科目编码（**加项**，精确匹配）
        direction   取数方向：
                      "credit" → 取 mc（贷方发生额），用于收入类
                      "debit"  → 取 md（借方发生额），用于成本费用类
        excel_row   模板中的 Excel 行号（1-based，见 EXCEL_ROW）
        exclude_codes   减项科目编码（默认空）
        digest_keywords 摘要关键字（任一命中即可）；空元组 = 不限摘要
        equal_codes     对方科目编码（精确匹配）；空元组 = 不限对方科目
        fill_total      是否填 O 列（本年累计数），默认 True
        note            口径来源与备注（写日志，便于与财务对账）

    【为什么 direction 要做成字段？】
        利润表里"收入"和"费用"在账上方向相反。
        更关键的是：**结转损益时收入记在借方**（见模块头部说明③），
        若想当然地"收入取贷方"，整张表会全是 0 且不报错。
        方向必须显式声明、逐行可查。

    【为什么摘要/对方科目也做成字段，而不是写死在 fetch 里？】
        每一行的筛选条件不同（有的要筛摘要、有的要筛对方科目、
        有的两者都要、有的都不要）。做成字段后：
            · 新增一行 = 加一条配置，不动逻辑代码；
            · 口径集中在一处，便于与财务逐行核对；
            · 测试可直接遍历配置检查"哪一行的口径没写"。

    【为什么用 excel_row 而不是 A 列文字定位？】
        LK02 的 A 列项目名带大量前导空格（如 '  其中： 出口主营业务收入'），
        且同一段文字可能出现在多行（'其他业务收入'与'其他业务成本'）。
        按行号定位更精确。行号已逐格核对模板并集中写在 EXCEL_ROW 里，
        模板若挪行，只需改 EXCEL_ROW 一处。
    """

    key: str
    label: str
    codes: tuple[str, ...]
    direction: str                          # "debit" 或 "credit"
    excel_row: int                          # Excel 行号（1-based）
    exclude_codes: tuple[str, ...] = ()
    digest_keywords: tuple[str, ...] = ()
    equal_codes: tuple[str, ...] = ()
    fill_total: bool = True
    note: str = ""


# =====================================================================
# 行项目配置表 —— 每确认一条需求，就在这里加一行
# =====================================================================
ROW_SPECS: tuple[ProfitRow, ...] = (
    ProfitRow(
        key="出口主营业务收入",
        label="其中：出口主营业务收入",
        # 【需求第 2 条（20261102 现场确认）】
        #   模板第 10 行 A 列 = '         其中： 出口主营业务收入'，行次 2
        #   金额列：C10（1月）… N10（12月）、O10（本年累计）
        codes=("600102",),
        direction="debit",
        excel_row=EXCEL_ROW[2],
        digest_keywords=DIGEST_SETTLE,
        equal_codes=(EQUAL_PROFIT,),
        note="ccode=600102 且摘要含结转损益类 且对方科目=4103，取 md 合计",
    ),
)


def _make_cell_map() -> dict[str, str]:
    """按 ROW_SPECS 自动生成 cell_map（指标名 → 单元格）。

    【为什么自动生成，而不是手写 13 个格子？】
        每行要登记 12 个月 + 1 个累计 = 13 个格子。
        手写时漏一个，该格就不进勾稽核对清单，
        日后"核对通过"可能是假象（漏检）—— 而这种漏检**没有任何症状**。
        自动生成后，"配置里有这一行"就必然意味着"这 13 格都登记了"。

    【为什么仍保留 "会计期间"？】
        C4 是需求明确要求写入的格子，不属于任何 ROW_SPECS 行，
        必须单独登记。
    """
    mapping: dict[str, str] = {"会计期间": PERIOD_CELL}
    for spec in ROW_SPECS:
        for month, col in COL_OF_MONTH.items():
            mapping[f"{spec.key}_{month}月"] = f"{col}{spec.excel_row}"
        if spec.fill_total:
            mapping[f"{spec.key}_本年累计"] = f"{COL_YEAR_TOTAL}{spec.excel_row}"
    return mapping


@register
class LK02Report(BaseReport):
    """利润表"""

    code = "LK02"
    name = "利润表"

    # 关键指标 → 单元格，供父类 verify() 读回做勾稽核对。
    # 【原则】只登记"确实写入"的格子，不要登记表头文字格。
    #   注意：C4 写的是年月文字，勾稽核对模块遇到非数值会判"无法核对"，
    #   这是已知行为（与 LK01 的 C4 完全一致），不影响其他指标。
    cell_map = _make_cell_map()

    def __init__(self):
        super().__init__()
        # {row_key: {month: amount}} —— 各月取数结果
        self.row_values: dict[str, dict[int, float]] = {}
        # {row_key: 有数据的月份集合} —— 用于决定"写哪些格、哪些留空"
        #
        # 【为什么必须单独记，不能靠 amount != 0 判断？】
        #   某月真实结转金额可能恰好为 0（例如该月无出口收入），
        #   此时"有数据且为 0"与"没取到数据"是两种不同事实：
        #     · 有数据为 0 → 应写 0（说明该月确实没有出口收入）
        #     · 没取到     → 应留空（说明未取数/未结转，需人查）
        #   若用金额是否为 0 来区分，这两种情况会被混为一谈。
        self.row_has: dict[str, set[int]] = {}

    # ------------------------------------------------------------------
    # 第1步：取数
    # ------------------------------------------------------------------
    def fetch(self, ctx: ReportContext) -> None:
        """按"筛选条件分组"批量取数，再按行分流汇总。

        【为什么按筛选条件分组，而不是每行一次查询？】
            本表多数行共用同一套口径（同样的摘要关键字、同样的对方科目），
            只是科目编码不同。把口径相同的行合并成一次查询，
            既能减少数据库往返，也保证同组数据来自同一快照。

            【分组键必须是什么？】
            direction + digest_keywords + equal_codes，三者缺一不可。
            少一个都会把口径不同的行混进同一批数据，
            结果错但不报错 —— 这类错误在报表上看不出任何异常。
        """
        year = ctx.period.year

        if not ROW_SPECS:
            logger.info("[%s] 暂无明细行项目配置，跳过取数", self.code)
            return

        # ---- 按筛选条件分组 ----
        groups: dict[tuple, list[ProfitRow]] = {}
        for spec in ROW_SPECS:
            sig = (spec.direction, spec.digest_keywords, spec.equal_codes)
            groups.setdefault(sig, []).append(spec)

        # ---- 各行的累加器 ----
        values: dict[str, dict[int, float]] = {s.key: {} for s in ROW_SPECS}
        has: dict[str, set[int]] = {s.key: set() for s in ROW_SPECS}

        for (direction, digests, equals), specs in groups.items():
            # 组内科目的并集（含减项）
            all_codes = tuple(sorted({
                c for s in specs for c in (*s.codes, *s.exclude_codes)
            }))
            if not all_codes:
                continue

            # 【缓存键】必须把口径也编进键里。
            #   只带科目会撞键：同一科目在不同口径（如"全部发生额"
            #   与"仅结转部分"）下结果完全不同，撞键会让后取的覆盖先取的，
            #   而且两次结果都不会报错。
            key = ctx.make_key(
                "GL_accvouch", "LK02利润行",
                direction,
                "|".join(digests),
                "|".join(equals),
                *all_codes,
            )
            sql, params = u8.gl_profit_amount_by_period(
                year, all_codes,
                amount_side=direction,
                digest_keywords=digests,
                equal_codes=equals,
            )

            try:
                rows = ctx.fetch(key, sql, params, report_code=self.code)
            except Exception as exc:  # noqa: BLE001
                # 单表取数失败不中断整体流程（43 张表里坏 1 张不牵连其他）
                logger.warning("[%s] 取数失败：%s", self.code, exc)
                rows = []

            # 科目 → [(所属行, 符号)]；符号 +1=加项，-1=减项
            #   与 LK01 完全相同的多值映射：一个科目允许同时归属多行。
            code_to_rows: dict[str, list[tuple[ProfitRow, int]]] = {}
            for spec in specs:
                for c in spec.codes:
                    code_to_rows.setdefault(c, []).append((spec, +1))
                for c in spec.exclude_codes:
                    code_to_rows.setdefault(c, []).append((spec, -1))

            for row in rows:
                try:
                    # 列顺序由 gl_profit_amount_by_period 的 SELECT 决定：
                    #   第0列 ccode、第1列 iperiod、第2列 amount
                    ccode = str(row[0]).strip()
                    iperiod = int(row[1])
                    amount = float(row[2]) if row[2] is not None else 0.0
                except (IndexError, TypeError, ValueError) as exc:
                    logger.warning("[%s] 跳过异常记录 %r：%s", self.code, row, exc)
                    continue

                if not 1 <= iperiod <= 12:
                    # 出现 13 期（U8 支持调整期）等情形时明确跳过，
                    # 不静默丢弃 —— 否则"某月没数"永远查不出原因。
                    logger.warning("[%s] 跳过非法会计期间 iperiod=%s（科目 %s）",
                                   self.code, iperiod, ccode)
                    continue

                for spec, sign in code_to_rows.get(ccode, []):
                    bucket = values[spec.key]
                    bucket[iperiod] = bucket.get(iperiod, 0.0) + sign * amount
                    # 【has 只由加项设置】与 LK01 同一原则：
                    #   若某行只有减项科目有记录、加项没有（数据异常），
                    #   累加结果会是负数。此时不标记"有数据"，
                    #   该格不写 —— 宁可不写，也不写一个负数误导人。
                    if sign > 0:
                        has[spec.key].add(iperiod)

        # ---- 收口保存 + 日志 ----
        self.row_values = values
        self.row_has = has

        for spec in ROW_SPECS:
            vals = values.get(spec.key, {})
            ok = has.get(spec.key, set())
            if not ok:
                logger.warning(
                    "[%s] %s（第%d行）：整年未取到符合条件的记录，全部保持空白。"
                    "请检查：①科目 %s 是否使用；②摘要是否含 %s；"
                    "③对方科目是否为 %s",
                    self.code, spec.label, spec.excel_row,
                    "、".join(spec.codes),
                    "、".join(spec.digest_keywords) or "（不限）",
                    "、".join(spec.equal_codes) or "（不限）",
                )
                continue

            detail = "、".join(f"{m}月={vals[m]:.2f}" for m in sorted(ok))
            logger.info("[%s] %s（第%d行）：%s",
                        self.code, spec.label, spec.excel_row, detail)

    # ------------------------------------------------------------------
    # 第2步：填表
    # ------------------------------------------------------------------
    def write(self, hook: _SheetHook, ctx: ReportContext) -> None:
        """填页眉 + 会计期间 + 各明细行。"""
        # 1) 通用页眉（公司名称等）
        # 【为什么 fill_period=False？】
        #   通用页眉会把年月写回"会计期间："标签所在格（本表是 A4）。
        #   而需求要求写到 C4。若两处都写，同一页会出现两个年月，
        #   财务会以为是数据重复。
        fill_common_header(hook, ctx, fill_period=False)

        # 2) 会计期间 → C4（需求第 1 条）
        # 【必须用 set_text，不能用 set】
        #   本页金额列普遍是数值/千分位格式。若用 set() 写 "2026年5月"，
        #   Excel 会把它识别成日期 2026-05-01 并存成序列号 46143，
        #   再按金额格式显示成 "46,143.00" —— 看起来像一笔钱。
        #   （LK01 的 C4 就踩过这个坑，见 docs/报表开发规范.md 第三节）
        hook.set_text(PERIOD_CELL, f"{ctx.period.year}年{ctx.period.month}月")
        logger.info("[%s] 已写入 %s=%s年%s月",
                    self.code, PERIOD_CELL, ctx.period.year, ctx.period.month)

        # 3) 各明细行：逐月写 + 本年累计
        for spec in ROW_SPECS:
            vals = self.row_values.get(spec.key, {})
            ok = self.row_has.get(spec.key, set())

            if not ok:
                # 【为什么不写 0？】
                #   数据库连不上、或该行口径未命中时，累加值是 0。
                #   若照写不误，报表会显示"出口主营业务收入为 0"，
                #   客户会误认为公司整年没有出口 —— 这个误读代价太大。
                #   约定：**没取到数据就保持空白**，让空白本身成为
                #   "未取数"的信号；原因已在上方 fetch() 写入日志。
                logger.warning("[%s] %s 无数据，整行保持空白（不写 0）",
                               self.code, spec.label)
                continue

            for month, col in COL_OF_MONTH.items():
                if month in ok:
                    hook.set(f"{col}{spec.excel_row}", vals[month])
                # 无数据的月份不写，保持模板原样（空白）

            if spec.fill_total:
                # 本年累计 = 各"有数据的月份"之和。
                # 【为什么不写死 SUM(C:N) 公式？】
                #   用值累加可避免"模板公式被覆盖"与"插删行导致范围失效"；
                #   也与 LK01 的"写数值不写公式"约定一致。
                total = sum(vals[m] for m in ok)
                hook.set(f"{COL_YEAR_TOTAL}{spec.excel_row}", total)
                logger.info("[%s] 已写入 %s 行：%s，本年累计 %s=%.2f",
                            self.code, spec.label,
                            "、".join(f"{COL_OF_MONTH[m]}{spec.excel_row}={vals[m]:.2f}"
                                      for m in sorted(ok)),
                            f"{COL_YEAR_TOTAL}{spec.excel_row}", total)
            else:
                logger.info("[%s] 已写入 %s 行：%s",
                            self.code, spec.label,
                            "、".join(f"{COL_OF_MONTH[m]}{spec.excel_row}={vals[m]:.2f}"
                                      for m in sorted(ok)))