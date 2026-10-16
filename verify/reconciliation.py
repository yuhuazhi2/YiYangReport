# -*- coding: utf-8 -*-
"""
勾稽关系核对模块
============================================================

【模块职责】
    实现模板"主要勾稽关系"表页中的核对公式（全表共 45 条）。
    每条勾稽关系表达为：某报表指标 = 其他报表指标的组合运算。

【什么是"勾稽关系"？】
    这是会计里的重要概念：不同报表之间的数字应当能对得上。
    例如：
        资产负债表的"货币资金"期末数
            = 现金流量相关报表的"现金 + 银行存款 + 其他货币资金"
    如果对不上，说明某张表取数错了，或者公式有问题。
    这是财务报表"自查"的核心手段，所以本模块非常重要。

【设计思路】
    - 定义 CheckItem：编号、描述、左边表达式、右边表达式。
    - 从 ReportContext.report_data 读取各报表登记的关键值。
    - 逐条比较，允许容差（默认 0.01 元，避免浮点误差误报）。
    - 输出核对报告（通过 / 不通过 / 无法核对）。

【关于表达式求值的安全说明】
    表达式形如 "LK01.货币资金_期末 + LK04.现金合计"。
    我们会先把它"翻译"成纯数字运算式（把指标替换成实际数字），
    再用"受限的 eval"求值——即禁用所有内置函数，只允许数字和四则运算。
    这样既灵活（公式改起来方便），又不会被执行恶意代码。
"""
from __future__ import annotations

import re                        # 正则，用于把 "CODE.指标" 替换成数字
from dataclasses import dataclass
from typing import Any, Optional

from core.context import ReportContext
from core.logging_setup import get_logger

logger = get_logger(__name__)

# 允许的误差（元）。两个数相差不超过它就认为是"相符"。
# 【为什么需要容差？】浮点数运算会有微小误差，比如 0.1+0.2 != 0.3。
#                    另外会计上也有"分位四舍五入"的情况。
DEFAULT_TOLERANCE = 0.01


@dataclass
class CheckItem:
    """
    一条勾稽关系。

    字段：
        no        : 编号（对应模板里勾稽关系表的序号）。
        desc      : 人类可读的描述。
        lhs       : 左边表达式，如 "LK01.货币资金_期末"。
        rhs       : 右边表达式，如 "LK04.现金合计 + LK04.银行存款合计"。
        tolerance : 本条的允许误差（默认 0.01）。
    """

    no: int
    desc: str
    lhs: str        # 左边表达式，如 "LK01.货币资金_期末"
    rhs: str        # 右边表达式，如 "LK04.现金合计 + LK04.银行存款合计"
    tolerance: float = DEFAULT_TOLERANCE


@dataclass
class CheckResult:
    """
    单条核对的结果。

    【新手难点：passed 为什么是 Optional[bool]？】
        它有"三种状态"，用 bool 装不下：
            True  → 相符
            False → 不符
            None  → 无法核对（因为某张相关报表没生成、或指标没登记）
        这种"三态"设计比"强行返回 False"更诚实：
        无法核对 != 核对失败，二者对用户的意义完全不同。
    """

    item: CheckItem
    lhs_value: Optional[float]
    rhs_value: Optional[float]
    passed: Optional[bool]     # None 表示无法核对（缺数据）
    message: str = ""


# ---------------------------------------------------------------------------
# 勾稽关系清单（示例，按模板"主要勾稽关系"表整理，可逐步补全到 45 条）
#
# 【重要】键名必须与各报表 verify() 登记到 ctx.report_data 的指标名一致，
#        否则会因"找不到数据"而显示"无法核对"。
# ---------------------------------------------------------------------------
DEFAULT_CHECKS: list[CheckItem] = [
    CheckItem(
        no=1,
        desc="LK01表货币资金期末数 = LK04表现金合计 + 银行存款合计 + 其他货币资金合计",
        lhs="LK01.货币资金_期末",
        rhs="LK04.现金合计 + LK04.银行存款合计 + LK04.其他货币资金合计",
    ),
    CheckItem(
        no=3,
        desc="LK01表应收账款期末数 = LK11表应收账款账龄分析表期末余额合计",
        lhs="LK01.应收账款_期末",
        rhs="LK11.期末余额合计",
    ),
    CheckItem(
        no=6,
        desc="LK01表存货期末数 = LK05表期末结存合计数",
        lhs="LK01.存货_期末",
        rhs="LK05.期末结存合计",
    ),
    CheckItem(
        no=21,
        desc="LK02表销售费用本期数 = LK27表销售费用当月合计",
        lhs="LK02.销售费用_本期",
        rhs="LK27.当月合计",
    ),
    CheckItem(
        no=24,
        desc="LK02表财务费用本期数 = LK28表财务费用当月合计",
        lhs="LK02.财务费用_本期",
        rhs="LK28.当月合计",
    ),
]


class Reconciler:
    """
    勾稽关系核对器。

    【典型用法】
        rec = Reconciler(ctx)
        results = rec.run()              # 拿到结构化结果列表
        text = rec.report_text()         # 拿到可直接显示的文本报告
    """

    def __init__(self, ctx: ReportContext, checks: Optional[list[CheckItem]] = None):
        """
        参数：
            ctx    : 报表上下文（内含各表登记的关键指标）。
            checks : 自定义核对清单。不传则用 DEFAULT_CHECKS。

        【新手须知：checks if checks is not None else DEFAULT_CHECKS】
            不能写 checks or DEFAULT_CHECKS！
            因为空列表 [] 在布尔判断里是"假"，
            传了空列表会被误判成"没传"而改用默认清单。
            所以要用 is not None 精确判断"到底传没传"。
        """
        self.ctx = ctx
        self.checks = checks if checks is not None else DEFAULT_CHECKS

    # ------------------------------------------------------------------
    def _eval(self, expr: str, results: dict) -> Optional[float]:
        """
        在受限命名空间下求值表达式。

        表达式形如 "LK01.货币资金_期末 + LK04.现金合计"。
        做法：把每个 "CODE.指标名" 替换成它对应的数值，再算结果。

        返回：计算出的数值；若缺数据或算不出来则返回 None。

        【新手难点：re.sub 配合"替换函数"】
            re.sub(模式, 替换, 字符串) 本来是把匹配到的内容换成文本，
            但第二个参数也可以传一个"函数"，每匹配到一处就调用它一次，
            用它的返回值作为替换内容。
            这里正好利用这一点：每匹配到一个 "CODE.指标名"，
            就去 ctx 里查它的实际值，返回 "(数值)" 用于参与运算。
        """
        # 【新手必学：嵌套函数（内部函数）】
        #   repl 定义在 _eval 内部，只能在这里用，外界访问不到。
        #   它的好处是能直接访问外层的 self（闭包特性）。
        def repl(m):
            """
            m 是正则匹配对象：
                m.group(1) → 代码部分，如 "LK01"
                m.group(2) → 指标名，如 "货币资金_期末"
            """
            code, key = m.group(1), m.group(2)
            data = self.ctx.get_report_data(code)
            if not data or key not in data:
                # 找不到数据：抛 KeyError，由外层捕获后返回 None
                raise KeyError(f"{code}.{key}")
            val = data[key]
            if val is None:
                raise KeyError(f"{code}.{key} 为空")
            # 【重要】用括号包住数值，避免替换后运算优先级出错
            #   例如 "a + b" 替换成 "1 + 2"，若原本是 "a * 2" 就成 "1 + 2 * 2" 了
            return f"({float(val)})"

        try:
            # 正则解释：([A-Za-z]\w*)\.([^\s+\-*/()]+)
            #   ([A-Za-z]\w*)        → 第1组：以字母开头，后跟字母/数字/下划线（代码）
            #   \.                   → 一个点号（需转义）
            #   ([^\s+\-*/()]+)      → 第2组：不含空格、运算符、括号的一串字符（指标名）
            py_expr = re.sub(r"([A-Za-z]\w*)\.([^\s+\-*/()]+)", repl, expr)
        except KeyError as exc:
            # 说明有指标查不到数据
            logger.debug("表达式缺数据：%s", exc)
            return None

        # 【新手必学：受限 eval 的安全做法】
        #   eval() 能执行任意 Python 代码，直接用很危险。
        #   但把 __builtins__ 设成空字典 {}，就禁用了所有内置函数
        #   （如 __import__、open、exec），只剩数字和四则运算能力。
        #   这就是"沙箱"思路：给能用的东西列白名单，其余一概不给。
        allowed = {"__builtins__": {}}
        try:
            return float(eval(py_expr, allowed))  # noqa: S307 受限命名空间
        except Exception as exc:
            # 表达式本身有问题（比如除零、语法错）也算"算不出来"
            logger.debug("表达式求值失败：%s | %s", expr, exc)
            return None

    # ------------------------------------------------------------------
    def run(self) -> list[CheckResult]:
        """
        执行全部核对，返回结果列表。

        【新手必学：三路分支的判断顺序】
            先判断 None，再判断 True/False。
            因为 None 在布尔判断里是"假"，若不先判 None，
            它会掉进 else 分支被误当成"不符"。
        """
        results: list[CheckResult] = []

        for item in self.checks:
            # 分别求出左右两边的值
            lhs_val = self._eval(item.lhs, {})
            rhs_val = self._eval(item.rhs, {})

            # 任一边算不出来 → 无法核对（注意：不是"不符"）
            if lhs_val is None or rhs_val is None:
                results.append(CheckResult(
                    item, lhs_val, rhs_val, None,
                    "数据不完整，无法核对（请确认相关报表已生成并登记指标）",
                ))
                continue

            # 【新手必学：abs() 取绝对值】
            #   只看差多少，不管谁大谁小。
            diff = abs(lhs_val - rhs_val)
            passed = diff <= item.tolerance

            # 【新手必学：条件表达式（三目）】
            #   "A" if 条件 else "B"
            #   条件成立取 A，否则取 B。
            #   这里格式化说明符 %.2f 表示保留两位小数。
            msg = (f"相符（差 {diff:.2f}）" if passed
                   else f"不符！左={lhs_val:.2f} 右={rhs_val:.2f} 差={diff:.2f}")
            results.append(CheckResult(item, lhs_val, rhs_val, passed, msg))

        return results

    # ------------------------------------------------------------------
    def report_text(self) -> str:
        """
        生成"文本版"核对报告，便于直接显示在界面或写入日志文件。

        返回一个多行字符串，形如：
            ============================================================
            勾稽关系核对报告
            ============================================================
            [ 1] √ 通过  LK01表货币资金期末数 = ...
                  相符（差 0.00）
            ...
            ------------------------------------------------------------
            合计：通过 3，不符 1，无法核对 1
        """
        results = self.run()

        # 【新手必学："字符串" * 60 重复拼接】
        #   "=" * 60 会得到 60 个等号组成的一条分割线。
        lines = ["=" * 60, "勾稽关系核对报告", "=" * 60]

        # 用三个计数器统计各种结果
        n_pass = n_fail = n_na = 0

        for r in results:
            # 【新手必学：is True / is False 精确比较】
            #   这里必须用 is 而不是 ==，
            #   因为 r.passed 可能是 None，而 None == False 是不成立的，
            #   但为了"一眼看出三态"，用 is 最清晰严谨。
            if r.passed is True:
                flag = "√ 通过"
                n_pass += 1
            elif r.passed is False:
                flag = "× 不符"
                n_fail += 1
            else:
                flag = "- 无法核对"
                n_na += 1

            # :>2 表示"右对齐、宽度至少 2"，让编号 1 和 10 对齐好看
            lines.append(f"[{r.item.no:>2}] {flag}  {r.item.desc}")
            lines.append(f"      {r.message}")

        lines.append("-" * 60)
        lines.append(f"合计：通过 {n_pass}，不符 {n_fail}，无法核对 {n_na}")

        # 【新手必学：join 把列表拼成一个字符串】
        #   "\n".join(lines) 用换行符把每一行连起来。
        #   比反复用 += 拼字符串高效得多。
        return "\n".join(lines)