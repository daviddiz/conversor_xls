#    para lanzar script de ejemplo ejecutar
#    python /home/david/Escritorio/xlrd-0.9.3/scripts/runxlrd.py 3rows *blah*.xls

import xlrd
book = xlrd.open_workbook("/home/david/Escritorio/openinnova/desglose nomina.xls")
print "The number of worksheets is", book.nsheets
print "Worksheet name(s):", book.sheet_names()
sh = book.sheet_by_index(0)
print sh.name, sh.nrows, sh.ncols
print "Cell D30 is", sh.cell_value(rowx=29, colx=3)
for rx in range(sh.nrows):
    print sh.row(rx)
# Refer to docs for more details.
# Feedback on API is welcomed.