# -*- coding: utf-8 -*-
"""通过字体大小提取每个工作表的标题"""
import sys
import xlrd

sys.stdout.reconfigure(encoding='utf-8')

FILE = '2026版新修订企业财务报模板.xls'
wb = xlrd.open_workbook(FILE, formatting_info=True)

print('%-18s %-6s %-6s %s' % ('工作表', '行', '字体', '标题内容'))
print('-' * 80)

for sh in wb.sheets():
    best = None  # (font_height, row, col, text)
    # 只扫描前 10 行
    for r in range(min(10, sh.nrows)):
        for c in range(sh.ncols):
            cell = sh.cell(r, c)
            if cell.ctype != xlrd.XL_CELL_TEXT:
                continue
            text = str(cell.value).strip()
            if not text:
                continue
            xf_index = sh.cell_xf_index(r, c)
            xf = wb.xf_list[xf_index]
            font = wb.font_list[xf.font_index]
            if best is None or font.height > best[0]:
                best = (font.height, r, c, text)
    if best:
        # 字体高度以 1/20 磅为单位
        pt = best[0] / 20.0
        print('%-18s R%-5d %-6s %s' % (sh.name, best[1] + 1, '%.1fpt' % pt, best[3]))
    else:
        print('%-18s %s' % (sh.name, '(未找到标题)'))