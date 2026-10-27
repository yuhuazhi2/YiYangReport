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
    }

    def __init__(self):
        """
        构造函数。

        【新手必学：先 super() 再扩展】
            父类 __init__ 会建好 self.data、self._hook。
            我们必须先调用它，再添加本表专属的属性。
        """
        super().__init__()
        # self.balance：本表专属数据，存放"科目编码 → 期末余额"。
        self.balance: dict[str, float] = {}

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

        # 生成缓存键。第二个参数 "all" 是"区分参数"，
        # 表示这次要的是"全部科目"，避免将来加"单科目"查询时撞键。
        key = ctx.make_key("GL_accsum", "all")

        # 让 u8 模块返回 (SQL, 参数) 二元组
        sql, params = u8.gl_account_summary(year, month)

        try:
            # 【关键】用 ctx.fetch 而不是 db.query：
            #   ctx.fetch 会自动查缓存 + 记录审计，是"正确姿势"。
            rows = ctx.fetch(key, sql, params, report_code=self.code)
        except Exception as exc:
            # 演示环境无数据库时，使用空数据，保证框架仍可运行。
            # 真实生产环境应改为抛出异常或明确报错。
            logger.warning("[%s] 取数失败，使用空数据演示：%s", self.code, exc)
            rows = []

        # 按科目编码聚合：把查询结果整理成 {科目编码: 期末余额}
        self.balance = {}
        for row in rows:
            try:
                # 【新手须知：为什么用 try 包住？】
                #   数据库返回的字段类型可能不是预期的（比如 None、字符串），
                #   float() 转换时可能报错。用 try 保证"一条坏数据不会毁掉整张表"。
                ccode = str(row[0]).strip()              # 第0列：科目编码
                me = float(row[5]) if row[5] is not None else 0.0  # 第5列：期末余额
                self.balance[ccode] = me
            except (IndexError, TypeError, ValueError):
                # 【新手必学：一次捕获多种异常】
                #   括号里列出的异常类型都会被捕获：
                #     IndexError  → 列数不够
                #     TypeError   → 类型不对（如 None 参与运算）
                #     ValueError  → 值不对（如 "abc" 转 float 失败）
                continue    # 跳过这条，处理下一条

        logger.info("[%s] 取得科目汇总 %d 条", self.code, len(self.balance))

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

        # 3) 示例：计算货币资金
        #    货币资金 = 现金(1001) + 银行存款(1002) + 其他货币资金(1012)
        #
        # 【新手必学：sum + 生成器 + startswith】
        #   startswith(("1001", "1002", "1012"))
        #     参数是"元组"，表示"以其中任意一个开头都算"。
        #   整体含义：把所有科目编码以这三个开头的余额加起来。
        money = sum(
            v for k, v in self.balance.items()
            if k.startswith(("1001", "1002", "1012"))
        )

        # 4) 演示阶段先只记日志，不实际写单元格，避免破坏模板结构。
        #    等 cell_map 与真实取数核对无误后，再打开下面这行。
        logger.info("[%s] 货币资金合计=%.2f（演示，未写入单元格）", self.code, money)

        # 5) 真实实现应类似：
        #    hook.set(self.cell_map["货币资金_期末"], money)
        #
        # 【新手须知：为什么用 self.cell_map 而不直接写 "C8"？】
        #   万一模板调整、单元格位置变了，只需改 cell_map 一处，
        #   不用满代码找 "C8"。这叫"把易变的配置集中管理"。