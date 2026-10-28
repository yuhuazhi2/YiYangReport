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
    └──────┴────────────────────┴────────────────┴─────────┴─────────┘

    另：C4 写会计期间（需求 20261024 第 1 条）。

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
        key        内部标识（用于日志与调试）
        label      财务名称（写日志用，建议与模板行项目名一致）
        codes      参与汇总的科目编码（精确匹配）
        end_cell   期末余额单元格（如 "C11"）
        begin_cell 上年年末余额单元格（如 "D11"）

        【注意】codes 是"精确匹配"。若某行需要包含明细科目
        （如应收账款 1122 及其下级 112201…），应改用前缀匹配的
        取数函数，不能在这里直接写 "1122"。
    """

    key: str
    label: str
    codes: tuple[str, ...]
    end_cell: str
    begin_cell: str


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
            本表当前涉及 1001/1002/1012 三个科目，若每行各查一次，
            就是 2 次数据库往返。合并成一次（去重后的科目集合），
            只需 1 次，且所有行看到的是同一份快照 —— 不会出现
            "两行数据来自不同时刻"的一致性问题。

            33 行都做完后科目会多一些，但仍是几十个，一次查询即可。
        """
        year, month = ctx.period.year, ctx.period.month

        # 汇总所有行用到的科目（去重 + 排序，保证缓存键稳定）
        all_codes = tuple(sorted({c for spec in ROW_SPECS for c in spec.codes}))

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

        # 科目 → 所属行（一个科目可能属于多行，故用 list）
        code_to_rows: dict[str, list[BalanceRow]] = {}
        for spec in ROW_SPECS:
            for c in spec.codes:
                code_to_rows.setdefault(c, []).append(spec)

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

            for spec in specs:
                if iperiod == month:
                    # 本期记录 → 期末余额
                    acc_end[spec.key] += me
                    has_end[spec.key] = True
                    det_end[spec.key].append(f"{ccode}={me:.2f}")

                if iperiod == 1:
                    # 本年1月记录 → 年初余额（上年年末）
                    acc_begin[spec.key] += mb
                    has_begin[spec.key] = True
                    det_begin[spec.key].append(f"{ccode}={mb:.2f}")

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