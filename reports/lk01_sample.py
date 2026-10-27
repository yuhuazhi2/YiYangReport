# -*- coding: utf-8 -*-
"""
LK01 资产负债表 —— 示例实现（框架演示）
============================================================

【模块职责】
    本文件演示"报表类"的完整写法，作为 LK02~LK33 的编写范例。
    想新增一张表？照这个文件的结构抄一份改改即可。

【重要说明】
    由于 U8 后台的真实字段与"资产负债表"各指标行的对应关系需现场确认，
    本实现仅演示：
    1. 如何从上下文取数（带缓存）。
    2. 如何定位并写入单元格。
    3. 如何按模板行项目填写。
    实际生产使用时，需将 cell_map 与取数逻辑替换为真实映射。

【★ 新增一张报表的标准步骤（新手必读）★】
    第1步：复制本文件为 lk02_xxx.py（比如 lk02_profit.py）。
    第2步：改类名（如 LK02Report）、code（"LK02"）、name（"利润表"）。
    第3步：在 fetch() 里写"怎么从数据库取数"。
    第4步：在 write() 里写"把数据写到哪些单元格"。
    第5步：在 cell_map 里声明"关键指标对应的单元格"，供勾稽核对用。
    第6步：到 reports/registry.py 的 load_builtin_reports() 里
           追加一行 import（触发注册），完成！
    就这样，不用改任何框架代码。
"""
from __future__ import annotations

from core.context import ReportContext
from core.excel_engine import _SheetHook
from core.logging_setup import get_logger
from queries import u8                       # U8 取数 SQL 集合
from reports.base import BaseReport
from reports.header import fill_common_header
from reports.registry import register        # 用它在导入时自动登记本类

logger = get_logger(__name__)


@register   # ← 【关键】这一行让本类被自动登记到注册表，别漏！
class LK01Report(BaseReport):
    """资产负债表"""

    # ---- 类属性：这张表的"身份证" ----
    code = "LK01"              # 表页代码（必须与模板中的代码一致）
    name = "资产负债表"         # 报表名称

    # 资产/负债指标 -> 模板中的"期末余额"单元格（示例，需按实际模板微调）
    # 模板中资产负债表金额一般位于 C 列（期末）与 D 列（期初）附近，此处为演示。
    #
    # 【新手必学：cell_map 的两大用途】
    #   1. 填表时知道往哪写（hook.set(cell_map["货币资金_期末"], 值)）。
    #   2. 核对时知道从哪读（父类 verify() 会按它读回单元格的值）。
    #   键命名建议用"指标名_期末/期初"，一目了然。
    # 【重要：为什么这里只有"会计期间"一项？】
    #   原先这里还有 "货币资金_期末": "C8"、"应收账款_期末": "C11"，
    #   那是搭框架时随手写的占位地址。实测发现：
    #     · C8 其实是表头文字"期末余额"所在的格；
    #     · C11 是空格。
    #   父类 verify() 会按本表把值读回来做勾稽核对，于是读到了
    #   字符串"期末余额"，换算数字时抛错，导致整份核对报告丢失。
    #
    #   结论：cell_map 里只放"已确认写入了数值/文字"的格子。
    #   等 LK01 的金额行项目真正实现时，再把每个指标对应的
    #   期末余额格逐条登记进来（届时地址要对着模板逐格核实）。
    cell_map = {
        # 【需求对接】《开发需求》20261024 第1条：
        #   "如果选择的会计期间是何年何月，那么 LK01 表页中的 C4 位置就填写何年何月"
        #   注意：模板里"会计期间："这个标签文字在第4行第1列(C1)，
        #   而需求指定的书写位置是同一行的 C4 —— 二者不是同一格。
        "会计期间": "C4",
        # 【需求对接】《开发需求》20261027 第2、3条：
        #   模板实测确认：R8 表头为 C8="期末余额"、D8="上年年末余额"；
        #   R11 的 A11="  货币资金"、B11=1（行次），C11/D11 为空待填。
        #   即 C11=本期期末余额，D11=上年年末余额，二者同一行。
        "货币资金_期末": "C11",
        "货币资金_上年年末": "D11",
    }

    def __init__(self):
        """
        构造函数。

        【新手必学：先 super() 再扩展】
            父类 __init__ 会建好 self.data、self._hook。
            我们必须先调用它，再添加本表专属的属性。
        """
        super().__init__()
        # 【本表取数结果】
        #   货币资金涉及 1001/1002/1012 三个科目，这里保存"汇总后的两个数"，
        #   而不是逐科目保存 —— 因为模板这一行只呈现合计数。
        #   若将来需要区分"现金/银行存款/其他货币资金"（如 LK04 表），
        #   再另外明细保存。
        self.monetary_end = 0.0      # 期末余额（本期 me 汇总）→ C11
        self.monetary_begin = 0.0    # 上年年末余额（本年1月 mb 汇总）→ D11
        # 【关键标记：本次是否真的取到了数据】
        #   数据库连不上时，若把 0 写进 C11/D11，客户会看到"货币资金 = 0.00"，
        #   误以为公司账上真没钱 —— 这是极危险的信息误读。
        #   所以约定：**没取到数据就不写格子，保持空白**，
        #   并在日志里说明原因。空白 + 日志，比错误的 0 诚实得多。
        #
        # 【为什么分成两个标记？】
        #   C11 来自本期记录、D11 来自本年1月记录，是两个独立的数据源。
        #   可能出现"本期有数据、1月记录缺失"的情况（例如中途建账）。
        #   若共用一个标记，就会把"只缺 D11"误判成"两格都不写"，
        #   白白丢掉已经取到的 C11。
        self.has_end_data = False
        self.has_begin_data = False
        self.has_data = False        # 兼容旧字段：任一有数据即为 True

    # ------------------------------------------------------------------
    # 第1步：取数
    # ------------------------------------------------------------------
    def fetch(self, ctx: ReportContext) -> None:
        """
        取科目汇总数据。

        演示要点：
        - 优先从缓存取（ctx.fetch 自带缓存）。
        - 用 ctx.make_key 生成规范缓存键。
        - 用 queries.u8 里的函数生成 SQL（把 SQL 与逻辑分离）。

        【新手必学：为什么 SQL 要放在 queries/u8.py？】
            如果 SQL 散落在各报表里，一旦 U8 表结构变了，
            要找遍所有文件才能改。集中管理 + 注释说明，
            将来维护（尤其是换版本/换账套）会轻松很多。
        """
        # 【新手必学：多重赋值】
        #   year, month = ctx.period.year, ctx.period.month
        #   一行同时给两个变量赋值，简洁清晰。
        year, month = ctx.period.year, ctx.period.month

        # 【缓存键】第二个参数是"区分参数"，标明这次取的是哪一类数据。
        #   用 "货币资金" 而不是 "all"，是因为将来本表还要取存货、应收等
        #   其他科目集合的数据；若都用 "all" 会撞键，导致拿到上次的结果。
        key = ctx.make_key("GL_accsum", "货币资金")

        # 【取数】一次查询同时拿到 mb(期初) 与 me(期末)，
        #   保证 C11 与 D11 来自同一批记录，不会错位。
        sql, params = u8.gl_monetary_funds(year, month)

        try:
            # 【关键】用 ctx.fetch 而不是 db.query：
            #   ctx.fetch 会自动查缓存 + 记录审计，是"正确姿势"。
            rows = ctx.fetch(key, sql, params, report_code=self.code)
        except Exception as exc:
            # 【为什么不抛异常？】
            #   单表取数失败不应让整个生成流程中断（43 张表里坏 1 张，
            #   不该让另外 42 张也白干）。这里记日志、标记无数据，
            #   由 write() 决定"不写格子"。
            logger.warning("[%s] 货币资金取数失败：%s", self.code, exc)
            rows = []

        # 【汇总：必须按 iperiod 分流，不能混在一起加】
        #
        # 【这是本项目最容易出错的一处，务必看懂】
        #   SQL 一次返回两个期间的数据：
        #     · 本期（iperiod = 当前月）→ 用 me 填 C11
        #     · 本年1月（iperiod = 1）  → 用 mb 填 D11
        #
        #   为什么不能简单地把所有行的 me 相加？
        #     因为 1 月那条记录也被查回来了，它的 me 是"1月末余额"，
        #     若混进 C11 就会重复计算。
        #
        #   为什么 D11 只取 1 月的 mb？
        #     现场确认：mb 的语义随月份变化 —— 只有 1 月记录的 mb 是
        #     "年初余额"；其他月份记录的 mb 是"本月期初"。
        #     所以 D11 必须专取 1 月那条的 mb，取 5 月的 mb 会得到 4 月末余额。
        self.monetary_end = 0.0
        self.monetary_begin = 0.0

        end_detail: list[str] = []    # 本期明细（供核对 C11 构成）
        begin_detail: list[str] = []  # 1月明细（供核对 D11 构成）
        has_end = False               # 是否取到"本期"记录
        has_begin = False             # 是否取到"本年1月"记录

        for row in rows:
            try:
                # 列顺序由 u8.gl_monetary_funds 的 SELECT 决定：
                #   第0列 ccode、第1列 iperiod、第2列 mb、第3列 me
                ccode = str(row[0]).strip()
                iperiod = int(row[1])
                mb = float(row[2]) if row[2] is not None else 0.0
                me = float(row[3]) if row[3] is not None else 0.0
            except (IndexError, TypeError, ValueError) as exc:
                logger.warning("[%s] 跳过异常记录 %r：%s", self.code, row, exc)
                continue

            if iperiod == month:
                # 本期记录 → 期末余额
                self.monetary_end += me
                end_detail.append(f"{ccode}={me:.2f}")
                has_end = True

            if iperiod == 1:
                # 本年1月记录 → 年初余额（即上年年末）
                self.monetary_begin += mb
                begin_detail.append(f"{ccode}={mb:.2f}")
                has_begin = True

            # 【注意：当本期就是 1 月时，上面两个 if 会同时命中同一行。
            #   这正是我们想要的结果：1 月的 me 是期末、mb 是年初，
            #   同一行同时提供两个口径，不会重复计算，也不会遗漏。】

        # 【C11 与 D11 分别独立判断是否写】
        #   两者来自不同期间，可能一个取到一个没取到（例如 1 月记录被删）。
        #   若共用一个 has_data，会把"只缺 D11"误判成"两个都没有"。
        self.has_end_data = has_end
        self.has_begin_data = has_begin
        self.has_data = has_end or has_begin   # 兼容旧字段

        logger.info(
            "[%s] 货币资金取数：共 %d 行 | 期末(C11)=%.2f ← [%s] | 年初(D11)=%.2f ← [%s]",
            self.code, len(rows), self.monetary_end,
            "、".join(end_detail) if end_detail else "未取到本期记录",
            self.monetary_begin,
            "、".join(begin_detail) if begin_detail else "未取到本年1月记录",
        )

    # ------------------------------------------------------------------
    # 第2步：填表
    # ------------------------------------------------------------------
    def write(self, hook: _SheetHook, ctx: ReportContext) -> None:
        """
        把数据写入表页。

        真实实现应形如：
            hook.set(self.cell_map["货币资金_期末"], money)
        """
        # 1) 先填通用页眉（公司名称等）
        #
        # 【关键：为什么传 fill_period=False？】
        #   通用页眉的默认行为是"找到写着'会计期间'的格子(C1)，在后面追加年月"。
        #   但本表按需求要把年月写到 C4。如果两处都写，同一页会出现两个年月，
        #   财务人员会以为是数据重复或程序出错。
        #   所以这里关掉通用页眉的期间填写，把"写哪里"完全交给本表决定。
        fill_common_header(hook, ctx, fill_period=False)

        # 2) 【需求实现】把会计期间写入 C4（《开发需求》20261024 第1条）
        #
        # 【新手必学：为什么用 cell_map 而不是直接写 "C4"？】
        #   与"货币资金_期末"一样，单元格地址统一登记在 cell_map 里：
        #     ① 万一模板调整、格位变了，只改 cell_map 一处；
        #     ② 父类 verify() 会按 cell_map 把值读回来做核对，
        #        这样"有没有真的写进去"能被自动验证，不必人工开 Excel 看。
        # 【必须用 set_text，不能用 set】
        #   C4 在模板里是"金额格式"的格子。若用 set() 写 "2026年5月"，
        #   Excel 会把它识别成日期 2026-05-01 并存成序列号 46143，
        #   再按金额格式显示成 "46,143.00" —— 看起来像一笔莫名其妙的钱。
        #   set_text() 会先设为文本格式再写值，从而原样显示年月。
        hook.set_text(self.cell_map["会计期间"],
                      f"{ctx.period.year}年{ctx.period.month}月")

        # 3) 【需求实现】货币资金：C11=期末余额，D11=上年年末余额
        #
        # 【为什么先判断 has_data？】
        #   若数据库连不上，rows 为空，此时 monetary_end 仍是 0.0。
        #   如果照写不误，C11 会显示 0.00 —— 客户看到"货币资金为零"，
        #   会误认为公司账上真的没钱。这个误读的代价太大。
        #
        #   所以约定：**没取到数据就保持空白**，让空白本身成为"未取数"
        #   的信号；同时上面 fetch() 已把原因写进日志，维护者可查。
        # 【C11 与 D11 分别独立判断】
        #   两者数据源不同（本期 me / 本年1月 mb），必须各自判断。
        #   若共用 has_data，会在"只缺其中一项"时把另一项也丢掉。
        if self.has_end_data:
            hook.set(self.cell_map["货币资金_期末"], self.monetary_end)
            logger.info("[%s] 已写入 C11=%.2f（期末余额）",
                        self.code, self.monetary_end)
        else:
            logger.warning("[%s] 未取得本期记录，C11 保持空白（不写 0）", self.code)

        if self.has_begin_data:
            hook.set(self.cell_map["货币资金_上年年末"], self.monetary_begin)
            logger.info("[%s] 已写入 D11=%.2f（上年年末余额，取自本年1月记录）",
                        self.code, self.monetary_begin)
        else:
            logger.warning("[%s] 未取得本年1月记录，D11 保持空白（不写 0）",
                           self.code)