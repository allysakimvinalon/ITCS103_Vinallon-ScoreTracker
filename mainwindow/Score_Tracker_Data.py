import openpyxl 
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Score-Tracker-Data"

ws.append(["Name", "Grade", "Remarks"])

wb.save("userdata.xlsx")
print("Excel file created successfully.")