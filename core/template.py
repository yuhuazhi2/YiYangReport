# -*- coding: utf-8 -*-
"""
模板解析模块
============================================================

【模块职责】
    1. 读取 `2026版新修订企业财务报模板.xls`（只读），解析每个表页的：
       - 工作表名（sheet name）
       - 表页代码（LK 编号，取工作表名中 "LKxx" 部分）
       - 报表标题（按字体大小定位：除 LK33 在第1行外，其余在第6/7行）
       - 标题所在行号
    2. 提供 43 张报表清单（代码 + 名称），供界面选择与报表注册。

【重要原则】
    本模块**绝不修改模板文件**，仅读取。
    所有"生成"操作都会另存到其他位置，保证模板始终干净。

【背景知识：为什么要用 xlrd 而不是 openpyxl？】
    .xls 是 Excel 97-2003 的老格式（二进制），openpyxl 只能读 .xlsx（新格式）。
    xlrd 能读 .xls，而且它"只能读不能写"，天然符合我们"不修改模板"的要求。
"""
from __future__ import annotations

import re                       # 正则表达式模块，从文本中"找规律"
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import xlrd                     # 用来读取 .xls（老版 Excel）文件

from core.logging_setup import get_logger

logger = get_logger(__name__)

# =====================================================================
# 【新手必学：正则表达式】
#   re.compile() 把一段"模式"预先编译好，之后可反复 search，效率更高。
#   模式 r"(LK\s*\d+(?:-\d+)?)" 逐段解释：
#     LK          → 先匹配字面量 "LK"
#     \s*         → 允许中间有 0 个或多个空白字符（空格/制表符）
#     \d+         → 一个或多个数字（表页序号）
#     (?:-\d+)?   → 可选的"-再加数字"部分，如 "-1"
#     (...)       → 最外层的括号表示"捕获组"，
#                   search() 后用 .group(1) 就能取出括号里匹配到的内容
#   r"..." 前缀表示"原始字符串"，反斜杠不会被转义，正则里强烈建议这么写。
#   re.IGNORECASE 表示大小写不敏感（lk / LK / Lk 都行）。
#
#   举例："LK12-1(B15-1)" → search 后 group(1) 得到 "LK12-1"
# =====================================================================
CODE_PATTERN = re.compile(r"(LK\s*\d+(?:-\d+)?)", re.IGNORECASE)

# 标题开头"年月"前缀的清洗规则。
# 【典型例子】LK33 在模板中的标题写成了：
#     "2026年05月产品销售收入及销售成本比较亏损情况明细表"
#   这个"2026年05月"是模板制作当时写死的期间。
#   但本程序是"每个月都要用"的：这个月生成 5 月的，下个月生成 6 月的。
#   如果把年月固化成表名的一部分，就会出现两个问题：
#     1) 换个月份生成，列表里还显示"05月"，与实际报表不符；
#     2) 同名报表堆积，用户分不清哪张是哪个月。
#   所以统一把标题开头的年月去掉，变成
#     "产品销售收入及销售成本比较亏损情况明细表"，
#   真实会计期间交给页眉/参数在生成时动态填入，表名则永久稳定。
#
# 【新手必学：^ 锚定开头】
#   模式里的 ^ 表示"只从字符串最开头匹配"。
#   因此只会去掉标题最前面的年月，不会误伤出现在
#   正文中间的"……2025年12月……"这类文字。
#
# 【逐段解释】^\s*\d{4}\s*年(?:\s*\d{1,2}\s*月)?\s*
#   ^            从开头
#   \s*          允许前面有空白
#   \d{4}        四位年份，如 2026
#   \s*年        年字，允许中间有空格
#   (?:...)?     整个"月"部分是可选的（"2026年"也能被去掉）
#   \d{1,2}\s*月 一到两位月份 + 月字，如 5月 / 05月
#   \s*          把年月后面残留的空格一并吃掉
TITLE_PERIOD_PREFIX = re.compile(r"^\s*\d{4}\s*年(?:\s*\d{1,2}\s*月)?\s*")

# 说明类表页（没有 LK 代码，是封面/目录/勾稽关系等辅助页）
INFO_SHEETS = ["报表封面", "报表目录及报送时间", "主要勾稽关系", "报表目录"]


@dataclass
class SheetInfo:
    """
    模板中"一个表页"的描述信息。

    【新手必学：@dataclass 数据类】
        当一个类主要是"装数据"（像一张表格的一行）时，
        用 @dataclass 装饰可以自动生成 __init__、以及方便的打印格式。
        否则你要手写一大堆 self.xxx = xxx。

        对比一下，不用 @dataclass 要写成：
            class SheetInfo:
                def __init__(self, sheet_name, code, title, ...):
                    self.sheet_name = sheet_name
                    self.code = code
                    ...
        用 @dataclass 只需按顺序声明"字段: 类型"，Python 帮你搞定。

    【字段说明（冒号后面是类型提示）】
        sheet_name : 工作表名，如 "LK12-1(B15-1)"
        code       : 表页代码，如 "LK12-1"；说明类表页为空字符串 ""
        title      : 报表标题（按字体大小定位出来的）
        title_row  : 标题所在行号（0 开始计数，即第1行是0）
        order      : 该表页在模板中的先后顺序
        nrows      : 总行数
        ncols      : 总列数
    """

    sheet_name: str          # 工作表名（如 LK12-1(B15-1)）
    code: str                # 表页代码（如 LK12-1），说明类为 ""
    title: str               # 报表标题（按字体定位）
    title_row: int           # 标题所在行（0-based）
    order: int               # 在模板中的顺序
    nrows: int
    ncols: int

    @property
    def display_name(self) -> str:
        """
        界面显示名：代码 + 名称。

        【新手须知：@property 让"方法"能当"属性"用】
            调用时不用加括号：sheet.display_name
            这里根据是否有 code，返回不同的展示形式：
              有代码 → "LK01  资产负债表"
              无代码 → 用标题，标题也空就用工作表名
        """
        if self.code:
            return f"{self.code}  {self.title}"
        return self.title or self.sheet_name


def _clean_title(raw: str) -> str:
    """
    清洗报表标题：去掉开头"2026年05月"这类年月前缀。

    参数 raw：从单元格里原样读到的标题文字。
    返回   ：去掉年月前缀后的标题；若清洗后为空则返回原文。

    【为什么要单独写成函数？】
        清洗规则只有一处定义，将来若要再加"去掉【】"、
        "去掉多余空格"之类规则，只改这里，所有表页同时生效。

    【新手注意：兜底很重要】
        万一某张表的标题整条就只有"2026年05月"，
        清洗后会变成空字符串。空标题在界面上就是一个空白条目，
        用户根本不知道那是什么表。所以这里做兜底：
        清洗结果为空，就退回原文，宁可留着年月也不能变空白。
    """
    if not raw:
        return raw
    cleaned = TITLE_PERIOD_PREFIX.sub("", raw).strip()
    return cleaned or raw


def _normalize_code(raw: str) -> str:
    """
    规范化代码：去除空格，统一大写。

    例如：'LK 12-1' → 'LK12-1'，'lk01' → 'LK01'
    这样不管模板里写的是大写小写、中间有没有空格，
    我们在代码里比较时都能统一，避免"看起来一样却不相等"的坑。
    """
    # str.replace(" ", "") 删除所有空格；.upper() 转大写
    return raw.replace(" ", "").upper()


class TemplateParser:
    """
    模板解析器。

    【典型用法（with 写法，用完自动关闭）】
        with TemplateParser("模板.xls") as parser:
            for sheet in parser.report_sheets:
                print(sheet.code, sheet.title)
    """

    def __init__(self, template_path: str | Path):
        """
        参数 template_path：模板文件路径。
        str | Path 的类型提示表示"这两个类型都行"。

        【新手须知：assert 与提前校验】
            文件不存在就没必要继续了，直接抛出异常（FileNotFoundError），
            这叫"快速失败"（fail fast），比后面出错再来找原因好得多。
        """
        self.template_path = Path(template_path)
        if not self.template_path.exists():
            raise FileNotFoundError(f"模板文件不存在：{self.template_path}")

        # self._wb 保存 xlrd 打开的工作簿对象；还没打开时是 None
        # 下划线开头是"内部属性"的命名约定
        self._wb: Optional[xlrd.book.Book] = None
        # self._sheets 保存解析出来的表页清单
        self._sheets: list[SheetInfo] = []

    # ------------------------------------------------------------------
    # 打开与关闭
    # ------------------------------------------------------------------
    def open(self) -> None:
        """
        打开模板（只读）。

        【新手须知：formatting_info=True 是什么？】
            xlrd 默认不读取"格式信息"（字体、颜色等）以节省内存。
            但我们定位标题要靠"字体大小"，所以必须打开这个开关，
            让它连格式一起读进来。代价是打开稍慢一点。
        """
        self._wb = xlrd.open_workbook(str(self.template_path), formatting_info=True)
        self._parse()   # 打开后立刻解析

    def close(self) -> None:
        """关闭模板（释放内存）。"""
        self._wb = None

    def __enter__(self):
        """支持 with 语句：进入时自动 open()。"""
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """支持 with 语句：退出时自动 close()。返回 False 表示不吞异常。"""
        self.close()
        return False

    # ------------------------------------------------------------------
    # 解析
    # ------------------------------------------------------------------
    def _parse(self) -> None:
        """解析全部工作表，生成 SheetInfo 清单。"""
        # assert 断言：确保前面已 open（否则继续下去会出错）。
        # 这是给"开发者"看的自检，不是给用户用的错误提示。
        assert self._wb is not None

        self._sheets = []   # 清空，允许重复解析

        # 【新手必学：enumerate】
        #   enumerate(列表) 会同时给出"下标"和"元素"：
        #       for idx, sh in enumerate(self._wb.sheets()):
        #   等价于：
        #       idx = 0
        #       for sh in self._wb.sheets():
        #           ...
        #           idx += 1
        #   这里 idx 正好可以当"表页顺序 order"。
        for idx, sh in enumerate(self._wb.sheets()):
            code = self._extract_code(sh.name)          # 从表名提取 LK 代码
            title, title_row = self._locate_title(sh)   # 按字体定位标题
            info = SheetInfo(
                sheet_name=sh.name,
                code=code,
                title=title,
                title_row=title_row,
                order=idx,
                nrows=sh.nrows,      # 该表总行数
                ncols=sh.ncols,      # 该表总列数
            )
            self._sheets.append(info)

        logger.info("模板解析完成：共 %d 个表页", len(self._sheets))

    @staticmethod
    def _extract_code(sheet_name: str) -> str:
        """
        从工作表名提取表页代码。

        【新手必学：@staticmethod 静态方法】
            用 @staticmethod 装饰的方法不需要 self，
            它就相当于"挂在这个类里的普通函数"，用来做与此类相关的工具活儿。
            调用方式：TemplateParser._extract_code("LK01...")

        【新手必学：正则搜索 flag】
            CODE_PATTERN.search() 会在字符串里"找第一个匹配"。
            找到返回 Match 对象，找不到返回 None。
            因此这里用 if m else "" 做空值保护。
        """
        m = CODE_PATTERN.search(sheet_name)
        # m.group(1) 取出第1个捕获组（括号里匹配到的内容）
        return _normalize_code(m.group(1)) if m else ""

    def _locate_title(self, sh) -> tuple[str, int]:
        """
        通过字体大小定位标题。

        思路：
        - 扫描前 10 行，找出"字体最大的文本单元格"作为标题。
        - 特殊：某些表（如 LK33）标题在第 1 行，也会被这套逻辑覆盖。

        返回 (标题文本, 标题所在行号)。

        【新手须知：为什么要"字体最大"？】
            财务模板里，表标题通常用最大字号（16~18pt），
            正文和表头字号小一些。所以"字体最大"是可靠的标题特征。
        """
        wb = self._wb
        best = None  # 用来记录当前找到的"最佳候选"，格式：(字号, 行号, 文本)

        # min(10, sh.nrows)：最多只看前 10 行，效率高且够用
        for r in range(min(10, sh.nrows)):
            for c in range(sh.ncols):
                cell = sh.cell(r, c)

                # 只关心文本类型单元格（数字、日期等跳过）
                # xlrd.XL_CELL_TEXT 是 xlrd 定义的"文本"类型常量
                if cell.ctype != xlrd.XL_CELL_TEXT:
                    continue

                text = str(cell.value).strip()
                # 跳过空文本，以及公司名这类"固定出现在每页但不该当标题"的文字
                if not text or text in ("力劲集团",):
                    continue

                # 取这个单元格的字体字号（可能需要跨表查询，用 try 包住防意外）
                try:
                    xf = wb.xf_list[sh.cell_xf_index(r, c)]   # 取"格式记录"
                    fh = wb.font_list[xf.font_index].height   # 由格式拿到字体，再取字号
                except Exception:
                    continue    # 取不到字号就跳过这个单元格

                # 【新手必学：打擂台式比较】
                #   best 为空，或当前字号比记录的大，就更新 best
                if best is None or fh > best[0]:
                    best = (fh, r, text)

        if best is None:
            return "", 0    # 一个字都没找到，返回空
        # 【关键】返回前统一清洗标题：
        #   去掉"2026年05月"这类年月前缀，让表名不绑定具体期间。
        return _clean_title(best[2]), best[1]   # 返回 (清洗后文本, 行号)

    # ------------------------------------------------------------------
    # 访问（对外提供的"只读"接口）
    # ------------------------------------------------------------------
    @property
    def sheets(self) -> list[SheetInfo]:
        """
        全部表页（含说明类）。

        【新手须知：为什么返回 list(self._sheets) 而不是 self._sheets？】
            这是"防御性编程"：返回一份副本，调用方即使修改这个列表，
            也不会破坏解析器内部的原始数据。相当于"给你复印件，别动我原件"。
        """
        return list(self._sheets)

    @property
    def report_sheets(self) -> list[SheetInfo]:
        """
        仅正式报表（有 LK 代码的表页）。

        【新手必学：列表推导式 + 条件过滤】
            [s for s in self._sheets if s.code]
            等价于：
                result = []
                for s in self._sheets:
                    if s.code:            # 有代码才要
                        result.append(s)
            注意：非空字符串在判断时视为 True，所以 if s.code 就是"有代码"。
        """
        return [s for s in self._sheets if s.code]

    def get_sheet(self, code: str) -> Optional[SheetInfo]:
        """
        按代码查找表页，找不到返回 None。

        先规范化 code 再比较，避免因大小写/空格导致"应该找到却找不到"。
        """
        key = _normalize_code(code)
        for s in self._sheets:
            if s.code == key:
                return s
        return None

    def codes(self) -> list[str]:
        """返回全部正式报表的代码列表，如 ["LK01", "LK02", ...]。"""
        return [s.code for s in self.report_sheets]


def load_template_catalog(template_path: str | Path) -> list[SheetInfo]:
    """
    便捷函数：读取模板并返回表页清单。

    把"创建解析器 → 打开 → 取数据 → 关闭"这套固定流程封装起来，
    调用方一行搞定：
        sheets = load_template_catalog("模板.xls")

    【新手须知：模块级函数 vs 类方法】
        这种"一次性便捷操作"写成模块级函数更好用，
        不需要调用方自己管理 parser 对象的生命周期。
    """
    parser = TemplateParser(template_path)
    parser.open()
    sheets = parser.sheets   # 注意：这里已经拿到副本，可以安全 close 了
    parser.close()
    return sheets


if __name__ == "__main__":
    # =================================================================
    # 【新手必学：if __name__ == "__main__" 是什么意思？】
    #   每个 .py 文件被运行时，Python 会设一个内置变量 __name__：
    #     - 如果这个文件被"直接运行"（python template.py），__name__ == "__main__"
    #     - 如果这个文件被"import 导入"，__name__ 等于模块名（如 "core.template"）
    #   所以下面这段代码只在"直接运行本文件"时执行，被 import 时不会执行。
    #   通常把"测试/演示代码"放这里，方便单独调试模块。
    # =================================================================
    import sys
    # Windows 终端默认编码可能不是 utf-8，会导致中文乱码；
    # 这行把标准输出强制设为 utf-8，保证中文打印正常。
    sys.stdout.reconfigure(encoding="utf-8")

    from config.settings import DEFAULT_TEMPLATE, PROJECT_ROOT

    cats = load_template_catalog(PROJECT_ROOT / DEFAULT_TEMPLATE)
    print(f"共 {len(cats)} 个表页")
    for s in cats:
        # :<14 表示"左对齐占 14 个字符宽"，让输出整齐
        print(f"  {s.code or '(说明类)':<14} R{s.title_row + 1:<3} {s.title or s.sheet_name}")