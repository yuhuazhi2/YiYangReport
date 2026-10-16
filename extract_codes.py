# -*- coding: utf-8 -*-
"""提取每个工作表内实际标注的表页代码（LK编号）"""
import sys
import re
import xlrd

sys.stdout.reconfigure(encoding='utf-8')

FILE = '2026版新修订企业财务报模板.xls'
wb = xlrd.open_workbook(FILE)

pattern = re.compile(r'LK\s*\d', re.IGNORECASE)

print('%-18s %-12s %-6s %s' % ('工作表名', '页内代码', '位置', '上下文'))
print('-' * 90)

for sh in wb.sheets():
    found = []
    for r in range(sh.nrows):
        for c in range(sh.ncols):
            cell = sh.cell(r, c)
            if cell.ctype != xlrd.XL_CELL_TEXT:
                continue
            text = str(cell.value).strip()
            if pattern.search(text):
                found.append((r + 1, c + 1, text))
    if found:
        for (r, c, text) in found[:5]:
            print('%-18s %-12s R%-5dC%-3d %s' % (sh.name, '', r, c, text.replace('\n', ' / ')[:60]))
        if len(found) > 5:
            print('    ... 共 %d 处' % len(found))
    else:
        print('%-18s %s' % (sh.name, '(页内未标注LK代码)'))