# -*- coding: utf-8 -*-
"""
占位报表类
============================================================

【模块职责】
    对于尚未实现具体取数逻辑的表页，使用 StubReport 顶替。
    它会：
    - 不做数据填充，仅保证表页被正确复制出来（保持模板原样）。
    - 在日志里留下一句提示，方便排查。

【为什么要做"占位"？】
    项目是分阶段实现的。LK01 做好了，但 LK02~LK33 还在开发中。
    如果未实现的表页直接报错，用户就没法生成任何报表。
    有了占位类，用户可以先生成已实现的部分 + 未实现的空白页，
    后续再逐个补上。这就是"渐进式交付"。

【新手须知：什么是"桩"（Stub）？】
    Stub 是软件测试里的术语，指"用来顶替真实实现的假东西"。
    它实现了同样的接口（fetch/write），但内部什么都不做或者只做最简单的动作。
"""
from __future__ import annotations

from core.context import ReportContext
from core.excel_engine import _SheetHook
from core.logging_setup import get_logger
from reports.base import BaseReport

logger = get_logger(__name__)


class StubReport(BaseReport):
    """
    占位实现：仅复制表页，不填数。

    【新手必学：子类如何覆写 code/name？】
        父类 BaseReport 里 code = ""（空字符串）。
        子类在 __init__ 里通过 self.code = code 动态赋值，
        这样同一个类就能表示"任意一张未实现的表"。
        这也是为什么占位类能通用——它的代码是运行时传入的。
    """

    def __init__(self, code: str, name: str = ""):
        """
        参数：
            code: 表页代码，如 "LK05"。
            name: 报表名称，留空则用代码代替。
        """
        # 【新手必学：super().__init__() 必须调用】
        #   父类的 __init__ 会创建 self.data 和 self._hook，
        #   不调用的话这些属性就不存在，后面 write/verify 会出错。
        super().__init__()
        self.code = code
        # 若没传名称就用代码兜底，保证 name 不为空
        self.name = name or code

    def fetch(self, ctx: ReportContext) -> None:
        """
        占位取数：不做任何数据库查询。

        【新手须知：为什么还要设置 self.data？】
            父类的 verify() 会读 self.data 相关的东西（通过 cell_map）。
            这里给一个标记性的字典 {"_stub": True}，
            表示"这是占位报表"，方便调试时分辨。
            下划线开头也是"内部/标记字段"的约定。
        """
        self.data = {"_stub": True}

    def write(self, hook: _SheetHook, ctx: ReportContext) -> None:
        """
        占位填表：什么都不写，保留模板原样。

        【新手须知：logger.debug 而不是 print】
            debug 级别默认不显示（要 logging 设成 DEBUG 才看得到）。
            这样正式使用时不会刷屏，排查时才开启，非常实用。
        """
        logger.debug("[%s] 占位报表，未填数", self.code)


def make_stub(code: str, name: str = "") -> StubReport:
    """
    工厂函数：创建一个占位报表实例。

    【新手须知：为什么要有"工厂函数"？】
        调用方只需 make_stub("LK05", "应收账款明细表")，
        不用记住 StubReport 的构造细节。
        将来若占位逻辑变复杂（比如要记录哪些表没实现），
        也只改这里一处，调用方不受影响。
    """
    return StubReport(code, name)