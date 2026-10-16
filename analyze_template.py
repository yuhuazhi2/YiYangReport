# -*- coding: utf-8 -*-
"""分析财务报表模板 .xls 文件结构，并输出完整报告到 report_structure.txt"""
import sys
import xlrd

FILE = '2026版新修订企业财务报模板.xls'
OUT = 'report_structure.txt'

wb = xlrd.open_workbook(FILE, formatting_info=True)

lines = []
def p(s=''):
    lines.append(str(s))

p('=' * 70)
p('文件: %s' % FILE)
p('工作表数量: %d' % wb.nsheets)
p('工作表列表: %s' % wb.sheet_names())
p('=' * 70)

for sh in wb.sheets():
    p()
    p('#' * 70)
    p('工作表名称: %s' % sh.name)
    p('行数: %d  列数: %d' % (sh.nrows, sh.ncols))
    p('-' * 70)
    merged = sh.merged_cells
    if merged:
        p('合并单元格数量: %d' % len(merged))
    # 遍历所有行
    for r in range(sh.nrows):
        row_vals = []
        for c in range(sh.ncols):
            v = sh.cell_value(r, c)
            if v != '' and v is not None:
                row_vals.append('C%d=%r' % (c, v))
        if row_vals:
            p('R%d: %s' % (r, ' | '.join(row_vals)))

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

print('报告已写入: %s' % OUT)
print('总行数: %d' % len(lines))
print('工作表摘要:')
for sh in wb.sheets():
    print('  %-20s 行数=%-5d 列数=%-5d 合并单元格=%d' % (sh.name, sh.nrows, sh.ncols, len(sh.merged_cells)))