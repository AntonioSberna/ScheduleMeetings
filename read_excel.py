


import openpyxl as opyxl



excel_file = "input_data.xlsx"



wb = opyxl.load_workbook(excel_file, data_only=True)
sh = wb["Global constraints"]

month = sh["A2"].value.split()[0]
year = sh["A2"].value.split()[1]


cell = "D6"
i = sh[cell].fill.start_color.index[2:]


wb.close()