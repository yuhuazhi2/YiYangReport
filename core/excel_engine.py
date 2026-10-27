# -*- coding: utf-8 -*-
"""
Excel 生成引擎
============================================================

【模块职责】
    1. 打开模板文件（只读方式，不修改原文件）。
    2. 按用户选择的表页代码，保留这些表页并删除未选中的报表页。
    3. 提供单元格读写接口（_SheetHook），供各报表填数。
    4. 另存为新文件（一个文件多表页）。

【为什么要用 win32com 驱动本机 Excel？】
    引擎基于 win32com（Python 通过 COM 接口操作 Windows 程序），
    相当于"用代码遥控打开 Excel 做操作"。
    好处：能完整保留 .xls 的格式、公式、合并单元格、列宽与字号。
    代价：运行环境必须安装 Microsoft Excel，而且会真的启动一个 Excel 进程。

    替代方案：若无 Excel，则回退到 xlrd+xlwt（会丢失部分格式，仅用于调试）。

【新手须知：COM 是什么？】
    COM 是 Windows 上让不同程序互相调用的技术。
    装了 pywin32 后，Python 就能像"按按钮"一样操作 Excel、Word 等软件。
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
import time
from pathlib import Path
from typing import Optional, Sequence

from core.logging_setup import get_logger
# _normalize_code 以下划线开头，属于"模块内工具函数"，
# 跨模块导入虽可行，但通常代表"内部约定"，此处复用是为了统一代码规范。
from core.template import SheetInfo, TemplateParser, _normalize_code

logger = get_logger(__name__)

# 【新手须知：为什么又要 try/except ImportError？】
#   win32com 是第三方库 pywin32 提供的，可能没装。
#   把它包起来并设置 HAS_WIN32 标志位，让程序能在"没有装 pywin32"的
#   机器上正常导入本模块，只是真正要生成报表时才提示错误。
try:
    import win32com.client as win32          # 用 as 起个别名，写起来短一些
    import win32com.client.gencache as gencache  # 类型库缓存（早期绑定用）
    import pythoncom                         # COM 初始化的辅助库
    HAS_WIN32 = True
except ImportError:
    win32 = None
    gencache = None
    pythoncom = None
    HAS_WIN32 = False


class ExcelEngineError(Exception):
    """Excel 引擎专用异常（自定义异常，便于上层精确捕获）。"""


# 【新手须知：为什么需要一个"修缓存"的函数？】
#   pywin32 的"早期绑定"会先在硬盘上生成一小段 Python 源码（缓存），
#   之后每次调用 Excel 都直接用它，速度更快。
#   缓存位置在 %TEMP%\gen_py\3.11\<Excel的类型库编号>\ 下。
#
#   但这个目录可能被"悄悄破坏"：
#     - 杀毒软件/电脑清理工具删掉了里面的 .py 文件；
#     - 异常关机或断电，只写了一半；
#     - 同一台机器装过多套 Python 或重装过 pywin32，互相覆盖。
#   一旦 .py 丢失、只剩 __pycache__ 里的 .pyc，Python 3 会把该目录
#   当成"空的命名空间包"——import 不报错，但里面什么都没有，
#   于是抛出 module has no attribute 'CLSIDToClassMap'。
#   这句报错极具误导性：用户会以为"电脑没装 Excel"，其实 Excel 好好的。
#
#   所以下面这个函数只做一件事：把"只有 .pyc、没有 .py"的坏目录找出来删掉，
#   让 pywin32 下次自动重新生成一份干净的。
#   判定依据非常保守——正常缓存目录里一定有 __init__.py，
#   只有它缺失时才认定损坏，绝不误删别的程序（Word、ADO 等）的缓存。
def _repair_gencache() -> int:
    """
    清理 pywin32 类型库缓存中"源文件丢失"的坏目录，并同步清掉内存里的残留。

    返回：实际删除的目录数量（0 表示没有发现损坏，或没有权限删）。

    【新手必学：为什么"只删文件夹"不够？】
        pywin32 判断"缓存有没有生成过"，靠两处记录：
          1. 硬盘上的 dicts.dat —— "我曾为这些类型库生成过文件"；
          2. 内存里的 sys.modules —— "这些模块我已经导入过了"。
        如果只删文件夹、不动这两处，pywin32 会认为"缓存还在"，
        于是既不去重新生成文件，又继续把那个空模块当成真的，
        结果就是"每次启动都失败，怎么删都删不好"。

        所以这里必须三件事一起做：
          删目录 → 清 sys.modules → 让 pywin32 重新扫描磁盘(Rebuild)。
        这三步做完，pywin32 才会真正重新生成一份完整的缓存。

    【新手必学：为什么只看第一层目录？】
        gen_py 目录下的结构是「一层程序文件夹 + 里面的模块文件」，
        我们只关心第一层的那些"类型库编号"文件夹，不需要钻得更深。
        所以用 os.listdir 就够了，比 os.walk 更直接。
    """
    if gencache is None:
        return 0

    try:
        root = gencache.GetGeneratePath()
    except Exception:      # 极少数环境拿不到路径，直接放弃修复
        return 0

    if not root or not Path(root).is_dir():
        return 0

    removed_names: list[str] = []
    for name in os.listdir(root):
        folder = Path(root) / name
        # 跳过文件（如 dicts.dat）和 __pycache__ 本身
        if not folder.is_dir() or name == "__pycache__":
            continue
        # 【关键判定】正常缓存一定有 __init__.py；
        #   只剩 __pycache__ 说明源文件已被删，属于坏缓存。
        if (folder / "__init__.py").is_file():
            continue
        try:
            shutil.rmtree(folder)
            removed_names.append(name)
            logger.warning("已清理损坏的 Excel 类型库缓存：%s", name)
        except Exception as exc:      # 文件被占用时可能删不掉，不能因此中断
            logger.warning("清理缓存目录失败（可忽略）：%s -> %s", name, exc)

    if not removed_names:
        return 0

    # 【第2件事】清掉内存里已导入的残留模块。
    #   否则下次 __import__ 会直接把那个"空壳"再拿出来，缓存永远修不好。
    prefix = "win32com.gen_py."
    for mod_name in list(sys.modules.keys()):
        if not mod_name.startswith(prefix):
            continue
        tail = mod_name[len(prefix):]
        if any(tail == n or tail.startswith(n + ".") for n in removed_names):
            del sys.modules[mod_name]

    # 【第3件事】让 pywin32 重新扫描磁盘、重写登记表。
    #   不加这一步，dicts.dat 里仍留着"已生成"的旧记录，
    #   EnsureDispatch 会以为缓存健在，从而拒绝重新生成。
    try:
        gencache.Rebuild(verbose=0)
        logger.info("已重建 pywin32 类型库登记表。")
    except Exception as exc:
        # 重建失败也不致命：后面还有后期绑定兜底
        logger.warning("重建类型库登记表失败（可忽略）：%s", exc)

    return len(removed_names)


def _dispatch_excel():
    """
    连接 Excel 程序，并自带"缓存坏了就修、修不好就换条路走"的能力。

    【三步策略，逐级兜底】
        第1步：正常走早期绑定（EnsureDispatch）——最快，能用就用。
        第2步：若报错，先修一次缓存（删掉坏目录），再试早期绑定。
        第3步：若仍失败，退回后期绑定（Dispatch）——它不读缓存，慢一点但一定能连。

    【为什么要分这么细，而不是一上来就用 Dispatch？】
        早期绑定速度快、还能享受代码提示，正常机器上应该优先用它；
        只有在"缓存确实坏了"这种少数情况下才需要降级。
        这样既保证速度，又保证"绝不会因为一个临时缓存问题就告诉用户没装 Excel"。

    【新手难点：EnsureDispatch 与 Dispatch 的区别】
        EnsureDispatch("Excel.Application")
            → 先确保硬盘上有一份类型信息（缓存），再按这份信息调用。
              好比先拿到"菜单目录"再点菜，点得快，但菜单坏了就点不了。
        Dispatch("Excel.Application")
            → 不查菜单，每次现问 Excel"你会不会这个操作"。
              慢一点，但菜单坏不坏都不影响。
        两者拿到的都是同一个 Excel 程序，后面的操作代码完全通用。
    """
    if not HAS_WIN32:
        raise ExcelEngineError("未安装 pywin32，无法使用 Excel 引擎。请执行: pip install pywin32")

    # 第1步：正常路径
    try:
        return win32.gencache.EnsureDispatch("Excel.Application")
    except Exception as exc:
        # 原始异常只进日志，不抛给界面（界面看到的是人话）
        logger.warning("Excel 早期绑定失败，准备修复类型库缓存：%s", exc)

    # 第2步：修缓存后重试
    #   注意：即使 _repair_gencache() 返回 0（没找到坏目录，可能已被外部删掉），
    #   也照样重试一次早期绑定 —— 因为此前的失败可能正是因为
    #   目录已被删、但登记表还留着旧记录，Rebuild 之后就能正常生成了。
    removed = _repair_gencache()
    try:
        app = win32.gencache.EnsureDispatch("Excel.Application")
        if removed:
            logger.info("类型库缓存已修复并重新生成，Excel 连接成功。")
        return app
    except Exception as exc:
        logger.warning("修复缓存后早期绑定仍失败，改用后期绑定：%s", exc)

    # 第3步：后期绑定兜底（不依赖任何缓存文件）
    try:
        app = win32.Dispatch("Excel.Application")
        logger.info("已改用后期绑定方式连接 Excel。")
        return app
    except Exception as exc:
        logger.error("连接 Excel 失败：%s", exc)
        raise ExcelEngineError("无法启动本机的 Microsoft Excel") from exc


class ExcelEngine:
    """
    Excel 生成引擎（win32com 实现）。

    【典型用法（with 写法，退出时自动关闭，杜绝 Excel 进程残留）】
        with ExcelEngine(template_path) as eng:
            eng.keep_only(["LK01", "LK02", "LK03"])   # 只留这几页
            hook = eng.sheet("LK01")                  # 拿到 LK01 页的写入句柄
            hook.set("A1", 123)                       # 往 A1 格写值
            eng.save_as("out.xls")                    # 另存为新文件
    """

    def __init__(self, template_path: str | Path):
        """
        参数 template_path：模板文件路径。

        【新手必学：初始化时的两次"提前校验"】
            1. 检查 pywin32 是否存在（没有就直接报错，别勉强）
            2. 检查模板文件是否存在
            这就是"快速失败"原则：早发现问题，报错信息才最清晰。
        """
        if not HAS_WIN32:
            raise ExcelEngineError(
                "未安装 pywin32，无法使用 Excel 引擎。请执行: pip install pywin32"
            )

        # .resolve() 把路径转成"绝对路径"，并解析掉 .. 之类的相对写法。
        # 原因：Excel 打开文件时按"当前工作目录"找相对路径容易出错，
        #       用绝对路径最稳妥。
        self.template_path = Path(template_path).resolve()
        if not self.template_path.exists():
            raise ExcelEngineError(f"模板文件不存在：{self.template_path}")

        self._excel = None    # Excel.Application 对象（整个 Excel 程序）
        self._wb = None       # 当前打开的工作簿（Workbook）
        self._opened = False  # 标记"是否已打开"，防止重复打开

        # 【新手必学：缓存字典】
        #   _sheet_cache 存放"代码 → _SheetHook 对象"的映射。
        #   同一个表页反复取用时，直接返回缓存的对象，
        #   不用每次都去 Excel 里找一遍（那很慢）。
        self._sheet_cache: dict[str, "_SheetHook"] = {}

        # 通过 xlrd 解析得到"代码 → 工作表名"映射。
        # 为什么需要它？Excel 里用"工作表名"定位表页，而用户/报表用"LK代码"，
        # 二者需要一个转换表。
        self._code_to_name = self._build_code_map()

    def _build_code_map(self) -> dict[str, str]:
        """
        用 xlrd 快速解析模板，建立"代码 → 工作表名"的映射。

        【新手必学：字典推导式】
            {s.code: s.sheet_name for s in parser.report_sheets}
            等价于：
                d = {}
                for s in parser.report_sheets:
                    d[s.code] = s.sheet_name
            一行顶四行，Python 里非常常用。

        【为什么要用 xlrd 而不是直接问 Excel？】
            xlrd 读取快得多，而 win32com 每次查询都要跨进程通信，很慢。
            用 xlrd 先把"地图"画好，后面 win32com 就能精准定位。
        """
        parser = TemplateParser(self.template_path)
        parser.open()
        mapping = {s.code: s.sheet_name for s in parser.report_sheets}
        # 顺带把"表页详情"和"全部表页（含说明类）"也缓存下来
        self._sheet_infos = {s.code: s for s in parser.report_sheets}
        self._all_sheets = parser.sheets
        parser.close()
        return mapping

    # ------------------------------------------------------------------
    # 生命周期（打开 / 关闭）
    # ------------------------------------------------------------------
    def open(self) -> None:
        """
        启动 Excel 并打开模板。

        【重要安全措施】
            用 ReadOnly=True 打开模板，即使程序出 bug 也不会写坏模板文件。

        【新手难点：CoInitialize / CoUninitialize 是什么？】
            COM 规定：每个线程使用 COM 前必须先"初始化"，
            用完后要"反初始化"。在单线程程序里这步常被省略，
            但在多线程（如 Qt 的 QThread）中必须成对出现，
            否则会报 "CoInitialize has not been called" 之类的错。
        """
        if self._opened:
            return   # 已打开就直接返回（这也是一种"幂等"设计）

        pythoncom.CoInitialize()   # 初始化 COM（与 close 里的 CoUninitialize 配对）

        # 【新手必学：为什么这里不直接写 EnsureDispatch？】
        #   连接 Excel 这一步"埋着雷"：pywin32 的类型库缓存可能被破坏，
        #   破坏后报的错是 "module has no attribute 'CLSIDToClassMap'"，
        #   极具误导性（看着像没装 Excel，其实 Excel 好好的）。
        #   所以统一走 _dispatch_excel()：它内部会自动修缓存、修不好就降级，
        #   保证"只要这台机器真的有 Excel，就一定能连上"。
        self._excel = _dispatch_excel()

        self._excel.Visible = False        # 不显示 Excel 窗口（后台跑）
        self._excel.DisplayAlerts = False  # 关闭"确定要删除吗"之类的弹窗，否则会卡住

        # UpdateLinks=0 表示"不更新外部链接"，避免弹窗或自动计算拖慢速度
        self._wb = self._excel.Workbooks.Open(
            str(self.template_path), ReadOnly=True, UpdateLinks=0
        )
        self._opened = True
        logger.info("Excel 已打开模板：%s（共 %d 页）",
                    self.template_path.name, self._wb.Sheets.Count)

    def close(self, save: bool = False) -> None:
        """
        关闭工作簿与 Excel 进程。

        参数 save：是否保存改动。默认 False，因为我们的模板本来就不该被改。

        【新手必学：try/finally 的意义】
            不管 try 里是否出错，finally 里的代码"一定会执行"。
            这里用它保证：哪怕关闭工作簿报错，也一定会去执行 Quit()，
            避免留下一个看不见的 Excel 进程占内存。
            用"任务管理器"查看进程，如果 taskkill 前有一堆 EXCEL.EXE，
            通常就是没做好这一步。
        """
        if not self._opened:
            return

        try:
            if self._wb is not None:
                self._wb.Close(SaveChanges=save)   # 关闭工作簿
        except Exception as exc:
            logger.warning("关闭工作簿异常：%s", exc)
        finally:
            self._wb = None    # 无论如何都置空

        try:
            if self._excel is not None:
                self._excel.Quit()   # 退出 Excel 程序
        except Exception:
            pass
        finally:
            self._excel = None
            self._opened = False
            self._sheet_cache.clear()   # 清空表页缓存
            try:
                pythoncom.CoUninitialize()   # 与 open() 里的 CoInitialize 配对
            except Exception:
                pass

    def __enter__(self):
        """支持 with 语句：进入时自动 open()。"""
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """
        支持 with 语句：退出时自动 close()。

        【重要】这里传 save=False，确保"模板不被保存"，
        即使程序中途报错，模板也绝不被修改。
        """
        self.close(save=False)
        return False

    # ------------------------------------------------------------------
    # 表页筛选
    # ------------------------------------------------------------------
    def keep_only(self, codes: Sequence[str]) -> list[str]:
        """
        仅保留指定代码的表页，删除其余报表页。

        说明类页（封面/目录/勾稽关系）默认一并保留——因为它们没有 LK 代码。

        参数 codes：用户勾选的表页代码列表，如 ["LK01", "LK05"]。
        返回：实际保留的报表页代码列表。

        【新手难点：为什么删除时要"从后往前"删？】
            假设要删第 2 页和第 5 页。如果先删第 2 页，
            原来的第 5 页就变成第 4 页了，再按"第 5 页"去找就找错了。
            从后往前删（先删第 5 页，再删第 2 页），
            前面的位置不受影响，就不会错位。
            代码里用 reversed(to_delete) 实现"倒序删除"。
        """
        self._ensure_open()

        # 用 set（集合）存放"要保留的表页名"。
        # 【新手必学：set 的特点】元素不重复、查找极快（比列表快很多）。
        keep_names = set()

        # 1) 说明类页全部保留（它们没有 code，是空字符串）
        for s in self._all_sheets:
            if not s.code:
                keep_names.add(s.sheet_name)

        # 2) 用户选中的报表页
        kept_codes = []
        for code in codes:
            key = _normalize_code(code)             # 统一成 "LK01" 这种格式
            name = self._code_to_name.get(key)      # 反查工作表名
            if name:
                keep_names.add(name)
                kept_codes.append(key)
            else:
                # 用户传了个模板里不存在的代码，记个警告但不中断
                logger.warning("模板中未找到表页代码：%s", code)

        # 3) 找出"所有需要删除的表页名"
        to_delete = []
        # 【新手须知：Excel 的 Sheets 从 1 开始计数，不是从 0】
        #   所以 range(1, Count + 1)，而不是 range(Count)。
        for i in range(1, self._wb.Sheets.Count + 1):
            sh = self._wb.Sheets(i)
            if sh.Name not in keep_names:
                to_delete.append(sh.Name)

        # 4) 从后往前删除，避免索引错位
        for name in reversed(to_delete):
            self._wb.Sheets(name).Delete()

        logger.info("已保留 %d 张报表页，删除 %d 张", len(kept_codes), len(to_delete))
        return kept_codes

    # ------------------------------------------------------------------
    # 表页访问
    # ------------------------------------------------------------------
    def sheet(self, code: str) -> "_SheetHook":
        """
        获取指定代码表页的写入句柄。

        返回一个 _SheetHook 对象，用它来写单元格。
        返回类型写 "_SheetHook"（带引号）是因为这个类在下面才定义，
        加引号可以让 Python 延迟解析（否则会报"名字未定义"）。

        【新手必学：返回字符串形式的前向引用】
            类型提示在运行时也会被求值，如果引用的类还没定义就会报错。
            写成带引号的字符串 "ClassName"，Python 只在做静态检查时才解析它，
            从而允许"先用后定义"。
        """
        self._ensure_open()
        key = _normalize_code(code)

        # 命中缓存就直接返回，省去跨进程查找
        if key in self._sheet_cache:
            return self._sheet_cache[key]

        name = self._code_to_name.get(key)
        if not name:
            raise ExcelEngineError(f"找不到表页：{code}")

        try:
            ws = self._wb.Sheets(name)
        except Exception as exc:
            raise ExcelEngineError(f"访问表页失败 {code}: {exc}") from exc

        # 包一层 _SheetHook，让填数代码更简洁（不用记 Excel 的复杂语法）
        hook = _SheetHook(ws, key)
        self._sheet_cache[key] = hook
        return hook

    # ------------------------------------------------------------------
    # 保存
    # ------------------------------------------------------------------
    def save_as(self, out_path: str | Path, fmt: str = "xls") -> Path:
        """
        另存为新文件（不改动模板）。

        参数：
            out_path : 输出文件路径。
            fmt      : 'xls' (97-2003, FileFormat=56) 或 'xlsx' (FileFormat=51)。

        【新手须知：FileFormat 的魔法数字】
            Excel 的另存为需要指定"格式代码"：
                56 = Excel 97-2003 工作簿 (.xls)
                51 = Excel 工作簿 (.xlsx)
            这些数字是微软定义好的常量，只能照抄。
            因为模板是 .xls，默认也用 56，避免格式转换引发兼容问题。

        【重要】用 SaveAs 而不是 Save，这样才能"写到新文件、不碰原文件"。
        """
        self._ensure_open()
        out = Path(out_path).resolve()

        # parents=True 表示"父目录不存在也一并创建"，
        # exist_ok=True 表示"目录已存在也不报错"。
        # 【新手必学】这两点如果漏写，多次运行程序会因目录已存在而报错。
        out.parent.mkdir(parents=True, exist_ok=True)

        file_format = 56 if fmt == "xls" else 51
        try:
            self._wb.SaveAs(str(out), FileFormat=file_format)
            logger.info("报表已保存：%s", out)
        except Exception as exc:
            raise ExcelEngineError(f"保存失败：{exc}") from exc
        return out

    def _ensure_open(self):
        """内部辅助：确保 Excel 已打开（没有就自动打开）。"""
        if not self._opened:
            self.open()


class _SheetHook:
    """
    单个表页的读写封装。

    类名以 _ 开头表示"本模块内部使用"，不打算给外部直接调用。

    【为什么要包这一层？】
        直接用 win32com 写单元格要记很多 Excel 专有语法，
        包一层后各报表的填数代码可以非常直观：
            hook.set("B5", 123)
            hook.set_row(10, 3, [1, 2, 3])
    """

    def __init__(self, ws, code: str):
        """
        参数 ws  : win32com 的工作表对象（Worksheet）。
        参数 code: 表页代码，如 "LK01"。
        """
        self.ws = ws
        self.code = code

    # ---- 写 ----
    def set(self, cell: str, value) -> None:
        """
        按 A1 记法写单元格（普通写入，交给 Excel 自行判断类型）。

        【新手必学：A1 记法】
            "B5" 表示第 B 列、第 5 行。
            字母是列，数字是行。这是 Excel 的通用写法。

        【警告：写文字请改用 set_text()】
            Excel 的 .Value 赋值会"猜类型"：写 "2026年5月" 会被当成日期，
            单元格内部存的其实是序列号 46143，若该格又是金额格式，
            就会显示成 46,143.00。详见 set_text() 的说明。
        """
        self.ws.Range(cell).Value = value

    def set_text(self, cell: str, value) -> None:
        """
        按 A1 记法写"纯文本"单元格，防止被 Excel 自动识别成日期或数字。

        【为什么必须要有这个方法？（真实踩坑记录）】
            需求要求把会计期间写进 LK01 的 C4。代码写的是字符串
            "2026年5月"，但生成出来的报表里 C4 显示的是 46,143.00。

            原因不是取数错，而是 Excel 的两步"自作聪明"：
              ① Range.Value 赋值会做类型识别，"2026年5月" 被理解成
                 日期 2026-05-01，内部存成日期序列号 46143
                 （Excel 里 1 = 1900-01-01，所以 2026-05-01 正好是 46143）；
              ② C4 这一格在模板里被设成了金额格式
                 （#,##0.00 千分位两位小数），于是 46143 被渲染成
                 "46,143.00" —— 看起来就像一笔莫名其妙的钱。

        【解决：两步，且顺序绝对不能反】
              ① 先把该格的数字格式设为文本 "@"
              ② 再写入字符串
            如果反过来（先写值、后改格式），值早已变成序列号 46143，
            改格式只能让它显示成 "46143"，仍然不是年月。

        【适用范围】
            页眉里的"会计期间""公司名称""制表""复核"这类纯文字信息，
            都应使用本方法写入，不要用 set()。
        """
        rng = self.ws.Range(cell)
        # 第①步：设为文本格式（"@" 是 Excel 里"文本"的格式代码）
        rng.NumberFormat = "@"
        # 第②步：写入字符串。None 统一写空串，避免 Excel 显示成 0
        rng.Value = "" if value is None else str(value)

    def set_text_rc(self, row: int, col: int, value) -> None:
        """
        按"行列号"写纯文本单元格（1-based），作用同 set_text()。

        【用途】需要按行列定位、又不想被识别成日期时用这个。
                例如页眉里的"会计期间："标签所在格。
        """
        cell = self.ws.Cells(row, col)
        cell.NumberFormat = "@"
        cell.Value = "" if value is None else str(value)

    def set_rc(self, row: int, col: int, value) -> None:
        """
        按"行列号"写单元格（行列都从 1 开始，遵循 Excel 习惯）。

        【新手易错点】这里的行列是 1-based，不是 Python 常见的 0-based！
            set_rc(1, 1, x) 写的是 A1，不是第一行第一列从0算。
            之所以这样设计，是为了跟 Excel 界面保持一致，减少换算错误。
        """
        self.ws.Cells(row, col).Value = value

    def set_many(self, items: Sequence[tuple]) -> None:
        """
        批量写单元格。

        支持两种格式的元组：
            (cell, value)         → 如 ("B5", 123)
            (row, col, value)     → 如 (5, 2, 123)

        【新手必学：靠元组长度自动判断类型】
            len(item) == 2 说明是"单元格地址+值"；
            否则（长度 3）就是"行+列+值"。
            这种写法让调用方更灵活。
        """
        for item in items:
            if len(item) == 2:
                self.set(item[0], item[1])
            else:
                self.set_rc(item[0], item[1], item[2])

    def set_row(self, row: int, start_col: int, values: Sequence) -> None:
        """
        从指定列开始，往右写一行数据。

        例：set_row(10, 3, [100, 200, 300])
            会在第 10 行的 C、D、E 列分别写入 100、200、300。

        【新手必学：enumerate 配合偏移量】
            enumerate(values) 给出 i 从 0 开始的序号，
            start_col + i 就是"第几列"。
        """
        for i, v in enumerate(values):
            self.ws.Cells(row, start_col + i).Value = v

    # ---- 读 ----
    def get(self, cell: str):
        """按 A1 记法读单元格的值。"""
        return self.ws.Range(cell).Value

    def get_rc(self, row: int, col: int):
        """按行列号读单元格的值（1-based）。"""
        return self.ws.Cells(row, col).Value

    # ---- 格式 ----
    def set_number_format(self, cell: str, fmt: str) -> None:
        """
        设置单元格的数字显示格式。

        例：hook.set_number_format("C8", "#,##0.00")
            让数字显示成带千分位、保留两位小数的形式（如 1,234.56）。

        【新手须知】这只是"显示格式"，不改变单元格里存的实际数值。
        """
        self.ws.Range(cell).NumberFormat = fmt

    @property
    def name(self) -> str:
        """当前工作表的名字。"""
        return self.ws.Name

    @property
    def used_rows(self) -> int:
        """已使用区域的行数。"""
        return self.ws.UsedRange.Rows.Count

    @property
    def used_cols(self) -> int:
        """已使用区域的列数。"""
        return self.ws.UsedRange.Columns.Count
    # ==================================================================
    # 【框架扩展区】为 43 张报表长期开发预留的通用能力
    # ------------------------------------------------------------------
    # 说明：下面这些方法不是"某一张表的特殊需求"，而是从 43 张表
    #      的实际结构中归纳出来的共性能力。放在框架层，好处是：
    #        · 后续 40 张表开发时不必各写一套；
    #        · 模板版式微调时，改一处即可适配全部报表。
    # ==================================================================

    # ---- 批量写（性能关键）----
    def set_block(self, row: int, col: int, rows: Sequence[Sequence]) -> int:
        """
        从 (row, col) 开始写入一个"二维表格块"（多行多列一次写完）。

        参数：
            row  : 起始行号（1-based）
            col  : 起始列号（1-based）
            rows : 二维序列，如 [[100, 200], [300, 400]]
        返回：实际写入的数据行数。

        【为什么需要它？】
            明细类报表（往来、存货、固定资产）常一次写几十上百行。
            逐格调用 set_rc() 会发起成百上千次跨进程 COM 调用，
            在"局域网 + 真实 Excel"环境下明显变慢。
            整块赋值只走一次 COM，快很多。

        【新手必学】
            self.ws.Range(左上, 右下).Value = 二维元组
            这是 Excel 最高效的写入方式。要求目标区域的
            行列数与数据完全一致，否则会错位或报错。
        """
        data = list(rows)
        if not data:
            return 0

        n_rows = len(data)
        n_cols = max((len(r) for r in data), default=0)
        if n_cols == 0:
            return 0

        # 【关键】补齐成矩形：Excel 要求每行列数一致
        matrix = []
        for r in data:
            row_vals = list(r) + [""] * (n_cols - len(r))
            matrix.append(tuple(row_vals))

        top_left = self.ws.Cells(row, col)
        bottom_right = self.ws.Cells(row + n_rows - 1, col + n_cols - 1)
        self.ws.Range(top_left, bottom_right).Value = tuple(matrix)
        return n_rows

    def set_formula(self, cell: str, formula: str) -> None:
        """
        写入 Excel 公式（不是值）。

        参数：
            cell    : 单元格地址，如 "C20"
            formula : 公式文本，必须以等号开头，如 "=SUM(C8:C19)"

        【使用场景】
            动态插行后，原来写死的合计公式（如 =SUM(C8:C14)）
            范围就不够了，需要按实际行数重写。

        【新手易错点】
            公式必须用 Formula 属性赋值。若用 Value，Excel 会把
            "=SUM(...)" 当纯文本显示出来。
        """
        self.ws.Range(cell).Formula = formula

    # ---- 批量读 ----
    def get_block(self, row: int, col: int, n_rows: int,
                  n_cols: int) -> list[tuple]:
        """
        读取一个矩形区域，返回二维数据。用于核对与调试。
        """
        top_left = self.ws.Cells(row, col)
        bottom_right = self.ws.Cells(row + n_rows - 1, col + n_cols - 1)
        raw = self.ws.Range(top_left, bottom_right).Value
        # 单个单元格时返回标量而非二维元组，统一包装
        if not isinstance(raw, tuple):
            return [(raw,)]
        return [tuple(r) if isinstance(r, tuple) else (r,) for r in raw]

    # ---- 按内容定位（应对模板版式微调）----
    def find_row(self, keyword: str, col: int = 1, start_row: int = 1,
                 max_row: int = 0, exact: bool = False) -> int:
        """
        在指定列中查找包含关键字的行，返回行号（1-based）；找不到返回 0。

        参数：
            keyword   : 要找的文字，如 "货币资金"
            col       : 在哪一列找（1-based，默认第 1 列）
            start_row : 起始行（默认 1）
            max_row   : 终止行；0 表示用已使用区域的行数
            exact     : True=完全相等；False=包含即算（默认）

        【为什么这是框架级能力？】
            43 张表的行项目位置各不相同，模板将来也可能微调。
            若每张表都写死"第 15 行是货币资金"，模板一挪行，
            全部报表同时失效，而且数字会悄悄写到别的行上。

            改为按项目名称定位后：
              · 模板挪行，代码不用改；
              · 代码自解释（找"货币资金"比找"第15行"直观）。

        【性能提示】
            单表几百行的字符串比较可忽略不计。
            若同一关键字要反复定位，建议报表内自行缓存结果。
        """
        key = str(keyword).strip()
        if not key:
            return 0

        end_row = max_row if max_row > 0 else self.used_rows
        for r in range(start_row, end_row + 1):
            v = self.get_rc(r, col)
            if v is None:
                continue
            text = str(v).strip()
            if not text:
                continue
            if (text == key) if exact else (key in text):
                return r
        return 0

    def find_cell(self, keyword: str, start_row: int = 1, max_row: int = 0,
                  max_col: int = 0) -> tuple[int, int]:
        """
        查找包含关键字的单元格，返回 (行, 列)；找不到返回 (0, 0)。

        【新手难点：为什么返回 (0,0) 而不是 None？】
            调用方通常紧接着解包：r, c = hook.find_cell(...)。
            若返回 None，解包会抛 TypeError；返回 (0,0) 后可以写
            "if r == 0: 跳过"，处理方式统一，不易漏判。
        """
        key = str(keyword).strip()
        if not key:
            return (0, 0)

        end_row = max_row if max_row > 0 else self.used_rows
        end_col = max_col if max_col > 0 else self.used_cols
        for r in range(start_row, end_row + 1):
            for c in range(1, end_col + 1):
                v = self.get_rc(r, c)
                if v is None:
                    continue
                if key in str(v):
                    return (r, c)
        return (0, 0)

    # ---- 动态行（数据量超出模板预留行数时）----
    def insert_rows(self, row: int, count: int = 1,
                    copy_from: int = 0) -> int:
        """
        在第 row 行之前插入 count 行；可把 copy_from 行的格式复制过来。

        参数：
            row       : 在第几行之前插入（1-based）
            count     : 插入几行
            copy_from : 格式样板行（1-based）；0 表示不复制格式
        返回：实际插入行数。

        【为什么必须复制格式？】
            Excel 插入空行时不带边框、不带数字格式。直接写数会导致
            报表"缺框线"，与模板风格不一致，财务人员一眼能看出异常。

            标准做法：插入空行 → 刷样板行格式 → 再写数据。
            本方法把前两步合并，调用方只需传 copy_from。

        【务必注意：插入后行号会整体下移】
            原本第 row 行及其后内容都会 +count。
            调用方必须按"新行号"计算后续所有位置（尤其是合计行）。
        """
        if count <= 0:
            return 0

        for _ in range(count):
            self.ws.Rows(f"{row}:{row}").Insert(Shift=1, CopyOrigin=0)

        if copy_from > 0:
            src = self.ws.Rows(f"{copy_from}:{copy_from}")
            dst = self.ws.Rows(f"{row}:{row + count - 1}")
            src.Copy()
            dst.PasteSpecial(Paste=-4122)   # xlPasteFormats：只贴格式
            # 【新手须知】必须清掉 Excel 的"剪切副本"状态，
            #   否则界面会残留闪烁虚线框，后续粘贴也可能受影响。
            try:
                self.ws.Application.CutCopyMode = False
            except Exception:
                pass

        return count

    def delete_rows(self, row: int, count: int = 1) -> int:
        """
        从第 row 行起删除 count 行（1-based）。返回实际删除行数。

        典型用途：模板预留 30 行空位、实际只有 5 条数据时，
                  删掉多余空行让报表干净。
        """
        if count <= 0:
            return 0
        self.ws.Rows(f"{row}:{row + count - 1}").Delete()
        return count

    # ---- 合并单元格 ----
    def merge(self, row1: int, col1: int, row2: int, col2: int) -> None:
        """把 (row1,col1) 到 (row2,col2) 的区域合并（1-based）。"""
        top_left = self.ws.Cells(row1, col1)
        bottom_right = self.ws.Cells(row2, col2)
        self.ws.Range(top_left, bottom_right).Merge()

    def unmerge(self, row1: int, col1: int, row2: int, col2: int) -> None:
        """取消指定区域的合并。"""
        top_left = self.ws.Cells(row1, col1)
        bottom_right = self.ws.Cells(row2, col2)
        self.ws.Range(top_left, bottom_right).UnMerge()

    # ---- 区域边界 ----
    def used_range(self) -> tuple[int, int, int, int]:
        """
        返回已使用区域 (起始行, 起始列, 结束行, 结束列)，均为 1-based。

        【踩坑提示：UsedRange 会"虚胖"】
            某区域即使内容已清空，只要曾被格式化过，Excel 仍算它已使用。
            所以不要用它判断"有没有数据"，它的用途是给出安全扫描边界。
        """
        ur = self.ws.UsedRange
        r1 = int(ur.Row)
        c1 = int(ur.Column)
        r2 = r1 + int(ur.Rows.Count) - 1
        c2 = c1 + int(ur.Columns.Count) - 1
        return (r1, c1, r2, c2)



def check_excel_available() -> tuple[bool, str]:
    """
    检查本机是否具备 Excel 引擎运行条件。

    返回 (是否可用, 提示消息)。

    【程序启动时的用途】
        main.py 会调用它，若不可用就在界面给出友好提示甚至阻止生成，
        避免用户点了按钮后才看到一堆晦涩的 COM 报错。

    【为什么这里要临时启动一次 Excel？】
        因为真正可靠的检测方式是"实际试着连一下 Excel"，
        光检查 pywin32 装没装还不够（Excel 本身可能缺失或损坏）。
        连上拿到版本号后立刻 Quit，开销很小。
    """
    if not HAS_WIN32:
        # 面向客户的措辞：不出现 pip 之类命令行指令，交给维护人员处理
        return False, "本机未安装 Python 的 pywin32 组件，无法生成 Excel 报表。"

    try:
        pythoncom.CoInitialize()
    except Exception:
        pass    # 已经初始化过也无妨，不影响后续

    try:
        # 【关键】走统一入口：自动修复损坏的类型库缓存，必要时降级为后期绑定。
        # 这样"缓存坏了"不会被误判成"没有装 Excel"。
        app = _dispatch_excel()
        ver = app.Version       # 顺便把 Excel 版本带回去，便于排查
        app.Quit()
        return True, f"检测到 Microsoft Excel，版本 {ver}"
    except Exception as exc:
        # 【重要】原始 COM 报错只进日志，不拼进界面文案。
        #   因为那些报错里满是 CLSID、类型库编号之类的术语，
        #   客户看不懂，反而会误以为是自己电脑的问题。
        logger.error("Excel 可用性检测失败：%s", exc, exc_info=True)
        return False, "未能启动本机的 Microsoft Excel。"
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass