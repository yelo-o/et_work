import xlwings
import pandas as pd
import openpyxl
from openpyxl import load_workbook



excel_app = xlwings.App(visible=False)
excel_book = excel_app.books.open('KT&G.xlsx')
excel_book.save('KT&G_refined.xlsx')
excel_book.close()
excel_app.quit()

wb = load_workbook(filename='KT&G.xlsx', data_only=True)

wb.save('KT&G_refined2.xlsx')