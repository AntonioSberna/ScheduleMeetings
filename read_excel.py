


import openpyxl as opyxl
import utility



excel_file = "input_data_filled.xlsx"

# Read attendees names
attendees = utility.read_attendees()


wb = opyxl.load_workbook(excel_file, data_only=True)
sh = wb["Global constraints"]

# I don't know if they are needed
month = sh["A2"].value.split()[0]
year = sh["A2"].value.split()[1]

# Read global constraints
cell = "H6"
i = sh[cell].fill.start_color.index[2:]


# Read meetings and attendees
sh = wb["Groups compositions"]
meet_attend = {}
for row in sh.iter_rows(min_row=4, max_row=sh.max_row, min_col=2, max_col=2):
    for cell in row:
        meet_attend[cell.value] = [sh.cell(row=cell.row, column=col).value for col in range(3, sh.max_column + 1)]


# Read attendees constraints
sh = wb["Attendees constraints"]


wb.close()