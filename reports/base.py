# -*- coding: utf-8 -*-
"""
报表基类
============================================================

【模块职责】
    定义"所有报表"的统一规范（接口）。每张表页（LK01~LK33）对应一个报表类，
    继承 BaseReport，实现：
    - fetch(ctx)      : 从数据库/缓存取数，存入 self.data。
    - write(hook, ctx): 把 self.data 写入对应表页的单元格。

    可选实现：
    - verify(ctx): 把本表关键值登记到 ctx，供勾稽核对。
    - cell_map   : 声明本表关键指标的单元格映射（指标名 → 单元格地址）。

【新手必学：什么是"抽象基类"（ABC）？】
    想象"报表"这个抽象概念。它必须能"取数"和"填表"，
    但具体怎么取、怎么填，每张表都不一样。
    抽象基类 = "规定必须有哪些方法，但不实现它们"。
    好处：框架层（ReportGenerator）可以放心地调用 report.fetch()，
          因为任何一张报表都保证有这个方法。

【新手必学：继承与"多态"】
    LK01Report、LK02Report 都继承 BaseReport，
    对它们的调用方式完全一样：rpt.generate(hook, ctx)
    Python 会自动去调用各自的 fetch/write。
    这就是"多态"——同一个调用，不同对象做出不同行为。
"""
from __future__ import annotations

from abc import ABC, abstractmethod   # ABC=抽象基类，abstractmethod=抽象方法装饰器
from typing import Any, Optional

from core.context import ReportContext
from core.excel_engine import _SheetHook
from core.logging_setup import get_logger

logger = get_logger(__name__)


class BaseReport(ABC):
    """
    所有报表的基类。

    【新手须知：class BaseReport(ABC) 的括号】
        表示 BaseReport 继承自 ABC。
        ABC 是"抽象基类"的标记，配合 @abstractmethod 使用，
        能强制子类必须实现指定方法。
    """

    # =================================================================
    # 类属性：定义在类里、所有实例共享。
    # 子类必须（通过赋值）覆盖 code，比如 code = "LK01"。
    # =================================================================
    code: str = ""            # 表页代码，如 "LK01"
    name: str = ""            # 报表名称

    # 可选：关键指标单元格映射 {指标名: "单元格地址"，如 "C8"}
    # 【新手须知】这里用 {} 作为默认值是安全的，因为我们只读不改它；
    #            但如果打算修改，就必须用 default_factory（见 context.py 的说明）。
    cell_map: dict[str, str] = {}

    def __init__(self):
        """
        构造函数。

        【新手须知：子类为什么要调用 super().__init__()？】
            super() 代表"父类"。子类写了 __init__ 后，
            父类的 __init__ 不会自动执行，必须手动调用，
            否则 self.data / self._hook 这些基础属性就不会被创建。
        """
        self.data: dict[str, Any] = {}          # 本表取到的数据，键值自定义
        self._hook: Optional[_SheetHook] = None # 写入句柄，generate 时赋上

    # ------------------------------------------------------------------
    # 必须实现（抽象方法）
    # ------------------------------------------------------------------
    @abstractmethod
    def fetch(self, ctx: ReportContext) -> None:
        """
        取数：将本表所需数据存入 self.data。

        应优先使用 ctx.fetch(...) 以获得缓存与审计（避免重复查库）。

        【新手必学：@abstractmethod 的作用】
            在父类里只写"方法名和文档"，不写具体实现（或只写一点说明）。
            任何继承 BaseReport 的子类如果没实现 fetch，
            实例化时就会报错：
                TypeError: Can't instantiate abstract class LK02Report
                           with abstract method fetch
            这就强制了"每张报表都必须自己写取数逻辑"，防止遗漏。
        """

    @abstractmethod
    def write(self, hook: _SheetHook, ctx: ReportContext) -> None:
        """填表：将 self.data 写入表页。"""

    # ------------------------------------------------------------------
    # 可选（有默认实现，子类可以不用写）
    # ------------------------------------------------------------------
    def verify(self, ctx: ReportContext) -> dict:
        """
        核对：返回本表关键校验值（供勾稽核对）。

        默认行为：从 cell_map 声明的单元格里"读回"刚写入的值，
        然后登记到 ctx（这样勾稽核对模块能直接拿到）。

        【为什么是"读回"而不是"直接返回 self.data"？】
            读回的是"实际写进 Excel 的值"，能发现"写入失败/写错格"的问题，
            更接近"验证最终报表"的目的。
        """
        values = {}
        if self._hook is not None:
            # cell_map 形如 {"货币资金_期末": "C8"}，
            # 下面把 name 和 cell 分别取出（这叫"遍历字典的 items"）。
            for name, cell in self.cell_map.items():
                try:
                    values[name] = self._hook.get(cell)   # 读回单元格的值
                except Exception:
                    # 取不到就记 None，不让核对流程因为一格读失败而崩溃
                    values[name] = None
        ctx.register_report_data(self.code, values)
        return values

    # ------------------------------------------------------------------
    # 生成流程（由 ReportGenerator 调用，子类通常不用改）
    # ------------------------------------------------------------------
    def generate(self, hook: _SheetHook, ctx: ReportContext) -> None:
        """
        完整生成流程：取数 → 填表 → 核对登记。

        【新手必学：模板方法模式（Template Method Pattern）】
            父类把"固定的流程步骤"写死在这里：
                1. fetch
                2. write
                3. verify
            而每一步的具体做，由子类决定。
            这样"流程"统一、稳定，"细节"灵活、可变化。
            这是非常经典的设计模式，值得记住。
        """
        self._hook = hook   # 保存句柄，verify 时要用它读回单元格
        logger.info("[%s] 开始取数：%s", self.code, self.name)
        self.fetch(ctx)                        # 第1步：取数
        logger.info("[%s] 开始填表", self.code)
        self.write(hook, ctx)                  # 第2步：填表
        self.verify(ctx)                       # 第3步：核对登记
        logger.info("[%s] 生成完成", self.code)

    # ------------------------------------------------------------------
    # 辅助
    # ------------------------------------------------------------------
    def describe(self) -> str:
        """返回"代码 + 名称"的描述文本，便于日志与界面展示。"""
        return f"{self.code}  {self.name}"