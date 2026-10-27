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


# =====================================================================
# 加载失败记录：{模块名: 失败原因}
#
# 【为什么框架层要专门记这个？】
#   本项目要分 43 个阶段逐步开发报表。若某一张表的模块写错了
#   （语法错、引用了不存在的库、字段名拼错），它绝不能拖累
#   其余已开发好的报表 —— 否则今天改 LK20 就把 LK01 弄坏了，
#   越往后越不敢动，项目必然失控。
#
#   所以这里采取"逐个隔离加载"：
#     ① 自动扫描 reports/ 目录下所有 lk*.py；
#     ② 每个模块单独 try/except 导入；
#     ③ 失败只记录、只降级为占位报表，不影响其他模块；
#     ④ 界面上能查"哪几张表加载失败了、为什么"。
# =====================================================================
_LOAD_FAILURES: dict[str, str] = {}


def load_failures() -> dict[str, str]:
    """返回本次运行中加载失败的报表模块及其原因（副本）。"""
    return dict(_LOAD_FAILURES)


def load_builtin_reports() -> None:
    """
    加载内置报表模块，触发 @register 注册（自动扫描 + 逐个隔离加载）。

    【新手难点：为什么"仅仅 import"就能完成注册？】
        Python 首次 import 某模块时，会从头到尾执行该模块代码。
        而报表模块里写着 @register class LKxxReport，
        执行到这里就会调用 register() 把类塞进 _REGISTRY。
        所以我们"只要 import，就完成了注册"。

    【本函数的设计要点：43 张表必须彼此独立】
        本项目分 43 个阶段逐步开发。如果采用"在文件里逐个手写 import"
        的老写法，会有三个致命问题：
          ① 每加一张表都要改本文件（改得越多，越容易碰坏别人的表）；
          ② 某张表写错语法/引用错库，会让整个 import 链中断，
             导致连已经做好的表都加载不出来；
          ③ 无法知道"到底哪几张没加载成功、为什么"。

        现在的做法：
          ① 自动扫描 reports/ 目录下的 lk*.py，无需改本文件；
          ② 每个模块在独立的 try/except 中导入，互不影响；
          ③ 失败原因记进 _LOAD_FAILURES，界面可查；
          ④ 失败的表现在没有实现类，会自动退化为占位报表
             （表页照样复制、只是不填数），整体流程不中断。

        这就是"新增一张表 = 新增一个文件"的落地方式，
        也是保证 43 张表长期并行开发不互相踩踏的基础。
    """
    import importlib
    import pkgutil
    from pathlib import Path

    # 清空上一轮的失败记录，允许重复调用（例如界面点"重新加载"）
    _LOAD_FAILURES.clear()

    # 定位 reports 包所在目录
    pkg_dir = Path(__file__).resolve().parent

    # 【新手必学：pkgutil.iter_modules】
    #   它列出某个目录下可被导入的子模块（不递归）。
    #   这里只取以 "lk" 开头的模块（约定：报表实现文件一律叫 lkXX_xxx.py），
    #   从而自动跳过 registry / base / stub / header 等框架文件。
    module_names: list[str] = []
    for info in pkgutil.iter_modules([str(pkg_dir)]):
        name = info.name
        if name.startswith("lk") and not name.startswith("lk_"):
            module_names.append(name)

    # 按名称排序，让加载顺序稳定、日志可读
    module_names.sort()

    loaded: list[str] = []
    for name in module_names:
        try:
            # 【关键】逐个独立导入。任何一张表出错，只影响它自己。
            importlib.import_module(f"reports.{name}")
            loaded.append(name)
        except Exception as exc:      # noqa: BLE001
            # 【为什么捕获 Exception 而不是只捕 ImportError？】
            #   报表模块里的错误五花八门：语法错(SyntaxError)、
            #   引用了不存在的名字(NameError)、类属性写错(ValueError,
            #   如 register 检查 code 为空)、甚至数据文件缺失(OSError)。
            #   想让"一张表坏掉不拖累其他表"，就必须把所有异常都接住。
            _LOAD_FAILURES[name] = f"{type(exc).__name__}: {exc}"
            # 用 warning 而非 debug：这类问题需要维护者知道
            logger.warning("报表模块 %s 加载失败（不影响其他表）：%s", name, exc)

    if loaded:
        logger.info("报表模块已加载 %d 个：%s", len(loaded), "、".join(loaded))
    if _LOAD_FAILURES:
        logger.warning("有 %d 个报表模块加载失败：%s",
                       len(_LOAD_FAILURES), "、".join(_LOAD_FAILURES))

    # 兼容旧版本：此前的实现是硬编码 import 这一个模块。
    # 现在由上面的自动扫描覆盖，但保留一次兜底导入，
    # 防止将来有人把文件改名成不符合 "lk*" 约定而意外漏载。
    try:
        from reports import lk01_sample  # noqa: F401
    except Exception as exc:      # noqa: BLE001
        logger.debug("lk01_sample 未额外加载：%s", exc)