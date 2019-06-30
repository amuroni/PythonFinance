"""Financial model class. This will hold all the input assumptions of the Model, passed on Main.py"""
import xlwings as xw


class Model:
    def __init__(self):
        self.wb = xw.Book("Model.xlsx")
        self.data = self.wb.sheets("DATA")
        self.summary = self.wb.sheets("Summary")
        self.CP = self.wb.sheets["CP"]
        self.project_start = self.data.range("D6")
        self.project_duration_months = self.data.range("D7")
        self.project_duration_years = self.data.range("D8")
        self.project_end = self.data.range("D9")
        self.construction_start = self.data.range("D10")
        self.construction_duration_months = self.data.range("D11")
        self.construction_end = self.data.range("D12")
        self.operation_start = self.data.range("D13")
        self.operation_end = self.data.range("D14")
        self.investment = self.data.range ("D16")
        self.costs = self.data.range("D24")
        self.revenues = self.data.range("D25")
        self.payables = self.data.range("D28")
        self.receivables = self.data.range("D29")
        self.VAT_costs = self.data.range("D32")
        self.VAT_revenues = self.data.range("D33")
        self.VAT_investment = self.data.range("D34")
        self.ires = self.data.range("D37")
        self.irap = self.data.range("D38")
        self.tax_avg = self.data.range("D40")

    def input(self):
        self.project_start.value = input("Please input a project start date (dd/mm/yyyy)")
        self.project_duration_years.value = input("Please set a project duration (n. of years)")
        self.project_duration_months.value = input("Please set a project duration (n. of months)")
        self.construction_start.value = input("Please set a construction start date (dd/mm/yyyy)")
        self.construction_duration_months = input("Please set a construction duration (n. of months)")
        self.investment.value = input("Please set and investment value to be realised in the construction phase")
        self.costs.value = input("Please set the avg amount of operating costs")
        self.revenues.value = input("Please set the avg amount of operating revenues")
        self.payables.value = input("Please set the expected amount of time for payables (n. of days)")
        self.receivables.value = input("Please set the expected amount of time for receivables (n. of days)")
        self.VAT_investment = input("Please set the % of VAT on the investment")
        self.VAT_revenues = input("Please set the % of VAT on revenues")
        self.VAT_costs = input("Please set the % of VAT on costs")
        tax_it = input("Would you like to apply Italian tax law? (Y/N)")
        if tax_it == "Y" or "y":
            self.ires.value = 0.039
            self.irap.value = 0.24
        else:
            self.tax_avg.value = input("Please set an avg amount of tax paid in your country of origin")
        self.wb.save()

    def check(self):
        if self.CP.range("H2").value == "OK":
            pass
        else:
            print("Errors were found, please run again the script and double check input values")
