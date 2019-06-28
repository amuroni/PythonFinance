import xlwings as xw

wb = xw.Book("Model.xlsx")  # load the excel file

summary = wb.sheets["Summary"]  # select a sheet
irr = summary.range("H22").value # select a cell, the IRR in this case
print("{:.2%}".format(irr))  # limitation in openpyxl, switched to xlwings

