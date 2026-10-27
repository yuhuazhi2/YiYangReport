# -*- coding: utf-8 -*-
"""
通用页眉填写
============================================================

【模块职责】
    所有表页共享统一的页眉结构
    （力劲集团 / 公司名称 / 币种 / 会计期间 / 制表 / 复核）。
    本模块通过在前若干行中搜索关键字单元格，把公司名称、会计期间等写入，
    避免各报表重复实现，也适应个别表页页眉行位置的差异。

【设计思路：为什么用"搜关键字"而不是"写死单元格地址"？】
    ❌ 写死地址：hook.set("B3", "公司名称:XXX")
       问题：万一某张表的页眉在 C4 而不是 B3，就写错了。
    ✅ 搜关键字：先找哪个格子里写着"公司名称"，再改它。
       好处：不管页眉在哪一行哪一列，都能自动找到，鲁棒性强。
    这是"用内容定位，而非用坐标定位"的典型思路。
"""
from __future__ import annotations

from core.context import ReportContext
from core.excel_engine import _SheetHook
from core.logging_setup import get_logger

logger = get_logger(__name__)

# 关键字 -> 用来在页眉区搜索定位
_KEYWORD_COMPANY = "公司名称"
_KEYWORD_PERIOD = "会计期间"
_KEYWORD_CURRENCY = "币种"


def _find_cell(hook: _SheetHook, keyword: str, max_row: int = 8) -> tuple[int, int] | None:
    """
    在前若干行中查找"包含关键字的单元格"，返回 (row, col)（1-based）。

    参数：
        hook     : 表页写入句柄。
        keyword  : 要找的关键字，如 "公司名称"。
        max_row  : 最多往下找多少行（页眉一般在前几行，限定范围可提速）。

    返回：找到则返回 (行, 列)，找不到返回 None。

    【新手必学：返回类型写成 tuple[int, int] | None】
        意思是"要么是一个二元组，要么是 None"。
        | 表示"或"（Python 3.10+ 的写法，等价于 Union）。

    【新手难点：为什么要 try/except 包住 hook.get_rc？】
        用 UsedRange 之外的空单元格，某些 Excel 版本会抛异常。
        这里逐格试探，遇到异常就当这格是空的，跳过即可。
    """
    # 【注意】Excel 行列都从 1 开始，所以 range 写 1..max_row
    for r in range(1, max_row + 1):
        # 页眉一般在前几列（A~G），不必扫到很远
        for c in range(1, 8):
            try:
                v = hook.get_rc(r, c)   # 读该格的值
            except Exception:
                continue                # 读不到就跳过
            if v is None:
                continue                # 空格子跳过
            # 把值转成字符串再判断"是否包含关键字"
            if keyword in str(v):
                return (r, c)
    return None


def fill_common_header(hook: _SheetHook, ctx: ReportContext,
                       company_name: str = "",
                       fill_period: bool = True) -> None:
    """
    填写通用页眉。

    做法：
    - 公司名称：<公司名>
    - 会计期间：<年>年<月>月
    保留原标识文字，仅在其后追加取值。

    【参数 fill_period：给个别表页自己填会计期间的机会】
        默认 True = 按通用规则，把"会计期间：2026年5月"写回标签所在格。

        但有些表页有明确的需求指定了另一个位置。例如 LK01 要求把
        会计期间写到 C4 格（而"会计期间："标签在 C1）。若这里仍按通用
        规则写一遍 C1、报表里又写一遍 C4，同一页就会出现两处年月 ——
        财务人员看到会以为数据重复或程序出错。

        所以这类表页调用本函数时传 fill_period=False，
        把"写哪里"的决定权完全交给该报表自己。

    【新手必学：位置解包的写法】
        pos = _find_cell(...)
        if pos:
            r, c = pos      # 把二元组拆成 r 和 c 两个变量
        这比写 pos[0]、pos[1] 更易读。

    【注意】if pos 判断的是"pos 不是 None"。
           因为 None 在布尔判断里为假，元组 (r, c) 为真。
    """
    # 公司名称
    pos = _find_cell(hook, _KEYWORD_COMPANY)
    # company_name 为空时不动它，避免把原有文字冲掉
    if pos and company_name:
        r, c = pos
        # 保留"公司名称:"这个标识，冒号后追加实际公司名
        #
        # 【为什么用 set_text_rc 而不是 set_rc？】
        #   页眉里很多格子被模板设成了金额/日期格式。用普通写入时，
        #   Excel 会把 "2026年5月" 识别成日期并存成序列号 46143，
        #   再按金额格式显示成 46,143.00（LK01 的 C4 就是这么踩的坑）。
        #   公司名一般不触发这种识别，但同一段代码统一用文本写入更稳妥，
        #   不必逐个字段去猜"这个值会不会被 Excel 误判"。
        hook.set_text_rc(r, c, f"公司名称:{company_name}")

    # 会计期间
    # 【新手须知】fill_period 为 False 时整段跳过，交回给报表自己写。
    if fill_period:
        pos = _find_cell(hook, _KEYWORD_PERIOD)
        if pos:
            r, c = pos
            # 用 ctx.period 里的年月拼出"会计期间：2026年5月"
            # 【务必用 set_text_rc】年月文字若用普通写入，会被 Excel
            #   当成日期、存成序列号，再按该格原有格式显示成一串数字。
            hook.set_text_rc(
                r, c, f"会计期间：{ctx.period.year}年{ctx.period.month}月"
            )