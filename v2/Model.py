"""Financial model class. This will hold all basic variables that will be editable from Main.py"""
import xlwings as xw


class Model:
    def __init__(self):
        self.wb = xw.Book("Model.xlsx")
        self.data = self.wb.sheets("DATA")
        self.summary = self.wb.sheets("Summary")
        self.project_start = self.data.range("D6")
        self.project_duration_months = self.data.range("D7")
        self.project_duration_years = self.data.range("D8")
        self.project_end = self.data.range("D9")
        self.construction_start = self.data.range("D10")
        self.construction_duration_months = self.data.range("D11")
        self.construction_end = self.data.range("D12")
        self.operation_start = self.data.range("D13")
        self.operation_end = self.data.range("D14")
        self.costs = self.data.range("D24")
        self.revenues = self.data.range("D25")
        self.payables = self.data.range("D28")
        self.receivables = self.data.range("D29")
        self.VAT_costs = self.data.range("D32")
        self.VAT_revenues = self.data.range("D33")
        self.VAT_investment = self.data.range("D34")
        tax_it = input("Would you like to apply Italian tax law? (Y/N")
        if tax_it == "Y":
            self.ires = self.data.range("D37")
            self.irap = self.data.range("D38")
        else:
            self.tax_avg = self.data.range("D40")
