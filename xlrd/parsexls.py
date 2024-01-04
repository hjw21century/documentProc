import xlrd
book = xlrd.open_workbook("test.xls", formatting_info=True)
print("The number of worksheets is {0}".format(book.nsheets))
for name in book.sheet_names():
    print(f"Worksheet name: {name}")
sh = book.sheet_by_index(1)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
for rx in range(sh.nrows):
    xfx = sh.cell_xf_index(rx, 0)
    xf = book.xf_list[xfx]
    print(xf.background.pattern_colour_index)