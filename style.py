from datetime import date, time
import xlsxwriter

workbook = xlsxwriter.Workbook(r"../data entry spread sheet_new1.xlsx")

worksheet = workbook.add_worksheet()

worksheet.data_validation('H2', {'validate': 'list','source': ['complete', 'inprogress', 'not complete']})


workbook.close()