# -*- coding: utf-8 -*-
"""
报表注册表
============================================================

【模块职责】
    管理"表页代码 → 报表类"的映射。
    生成流程通过本注册表获取每张表的生成器。

    - 已实现逻辑的报表在此登记具体类（如 LK01Report）。
    - 未登记的代码自动使用 StubReport（仅复制表页、不填数）。
    - 报表名称从模板解析结果获取，保证"代码 + 名称"准确。

【新手必学：什么是"注册表模式"？】
    问题：43 张报表，每张一个类。生成时给个代码 "LK05"，
          怎么知道该用哪个类？
    朴素做法：写一长串 if/elif：
        if code == "LK01": cls = LK01Report
        elif code == "LK02": cls = LK02Report
        ...
        （43 个分支，新增一张就要改这里，容易漏、容易错）

    注册表做法：
        用一张字典 {"LK01": LK01Report, "LK02": LK02Report, ...} 存映射。
        报表类自己"报到"（用 @register 装饰器），
        查询时 cls = _REGISTRY.get(code) 一句话搞定。
    好处：新增报表只需新建文件 + 加个 @register，不用改这段代码。
    这叫"开闭原则"：对扩展开放，对修改封闭。
"""
from __future__ import annotations

from typing import Optional

from core.logging_setup import get_logger
from core.template import SheetInfo
from reports.base import BaseReport
from reports.stub import make_stub

logger = get_logger(__name__)

# =====================================================================
# 已实现报表的登记表：{代码: 报表类}
# 随着开发推进，逐步加入 LK02、LK03 ……
#
# 【新手须知】字典的值是"类"本身（不是实例），
#   想用时再 cls() 实例化。
# =====================================================================
_REGISTRY: dict[str, type[BaseReport]] = {}


def register(cls: type[BaseReport]) -> type[BaseReport]:
    """
    装饰器：登记报表类。

    【新手必学：装饰器是什么？】
        装饰器本质是"接收一个函数（或类），返回它（或包装后的它）"。
        写法：
            @register
            class LK01Report(BaseReport):
                ...
        等价于：
            class LK01Report(BaseReport):
                ...
            LK01Report = register(LK01Report)

        也就是说：类定义好之后，Python 自动把它传给 register() 执行一次。
        register 把类存进 _REGISTRY 字典，最后把类原样返回（所以类还能正常用）。

    【为什么返回 cls？】
        装饰器如果不返回东西（或返回 None），
        那么 LK01Report 这个名字就会变成 None，后续就用不了了。
        必须原样返回，才能"顺便注册，不影响使用"。
    """
    if not cls.code:
        # 忘记写 code 是常见错误，早点报错比后面出诡异 bug 好
        raise ValueError(f"报表类 {cls.__name__} 未定义 code")
    # .upper() 统一大写，避免 "lk01" 和 "LK01" 被当成两个不同的键
    _REGISTRY[cls.code.upper()] = cls
    return cls


def get_report_class(code: str) -> Optional[type[BaseReport]]:
    """按代码获取报表类；未登记返回 None。"""
    return _REGISTRY.get(code.upper())


def create_report(code: str, name: str = "") -> BaseReport:
    """
    创建报表实例（工厂函数）。

    规则：
    - 若该代码"已实现"，实例化具体类（如 LK01Report()）。
    - 否则返回占位报表 StubReport（保证框架不至于因未实现而崩溃）。

    【新手必学：工厂函数（Factory）】
        调用方只说"给我 LK01 的报表对象"，不必关心是哪个类、怎么造。
        将来若要加缓存、加日志，也只需改这里一处。
    """
    cls = get_report_class(code)
    if cls is not None:
        return cls()                 # 已实现：直接用真正的报表类
    return make_stub(code, name)     # 未实现：用占位报表顶上


def build_report_set(codes: list[str],
                     catalog: list[SheetInfo]) -> list[BaseReport]:
    """
    根据选中的代码列表，构造有序的报表实例列表。

    名称从模板目录（catalog）中取得，并保持输入 codes 的顺序。

    【为什么顺序很重要？】
        报表之间有勾稽依赖（比如资产负债表要用到科目余额）。
        按 LK 编号顺序生成，能让前面表的数据先进缓存，后面表直接复用。
    """
    # 【新手必学：字典推导式 + 条件】
    #   {s.code: s.title for s in catalog if s.code}
    #   把"代码 → 标题"抽成字典，方便下面按代码查名称。
    name_map = {s.code: s.title for s in catalog if s.code}

    reports: list[BaseReport] = []
    for code in codes:
        key = code.upper()
        # get(key, "") 表示"取不到就给空字符串"，避免 KeyError
        name = name_map.get(key, "")
        reports.append(create_report(key, name))
    return reports


def implemented_codes() -> list[str]:
    """
    已实现取数逻辑的代码列表（按字母排序）。

    【用途】界面可以据此把"已实现"和"未实现"的表页区分显示，
          让用户知道哪些表能真正出数。
    """
    return sorted(_REGISTRY.keys())


def load_builtin_reports() -> None:
    """
    导入内置报表模块，以触发 @register 注册。

    【新手难点：为什么"仅仅 import"就能完成注册？】
        Python 首次 import 某模块时，会从头到尾执行该模块代码。
        而 LK01 模块里写着 @register class LK01Report，
        执行到这里就会调用 register() 把类塞进 _REGISTRY。
        所以我们"只要 import，就完成了注册"——这就是这种写法的巧妙之处。

    【为什么用 try/except 包住？】
        开发初期，某些报表模块可能还没写好、或引用了不存在的库。
        用 try 包住，保证"加载失败也只是少一张表"，不影响整体运行。
    """
    # 已实现的报表模块（逐步添加）
    try:
        from reports import lk01_sample  # noqa: F401
        # 【新手须知：noqa 注释】
        #   F401 是"导入了却没使用"的 lint 警告。
        #   但我们导入它本就是为了"触发注册"，不需要使用它，
        #   所以用 # noqa: F401 告诉检查工具"这里别报警告"。
    except ImportError as exc:
        # logger.debug 级别较低，平时不显示，排查问题时才看
        logger.debug("lk01_sample 未加载：%s", exc)