


import openpyxl as opyxl
import utility



excel_file = "input_data_filled.xlsx"

# Read attendees names
attendees = utility.read_attendees()


wb = opyxl.load_workbook(excel_file, data_only=True)
sh = wb["Settings"]

# Read colors legend
color_holidays = sh["AS42"].fill.start_color.index[2:]
color_workday = sh["AS43"].fill.start_color.index[2:]
color_error = sh["AS44"].fill.start_color.index[2:]


sh = wb["Global constraints"]

# I don't know if they are needed
month = sh["A2"].value.split()[0]
year = sh["A2"].value.split()[1]




# Read global constraints and creation of time slots
time_slots = []
max_col = sh[3][-1].column
for row in sh.iter_rows(min_row=5, max_row=sh.max_row, min_col=2, max_col=max_col):
    for cell in row:
        col = cell.fill.start_color.index[2:]
        if col == color_workday and cell.value != "X":
            time_slots.append(sh.cell(row=3, column=cell.column).value + " / " + sh.cell(row=cell.row, column=1).value) # così è un casino, ma non so come fare meglio
        if cell.value is not None and cell.value != "X":
            raise ValueError(f"Error in the Global Constraints - day {sh.cell(row=3, column=cell.column).value}, timeslot {sh.cell(row=cell.row, column=1).value}")




# Read meetings and attendees
sh = wb["Groups compositions"]
meet_attend = {}
for row in sh.iter_rows(min_row=4, max_row=sh.max_row, min_col=2, max_col=2):
    for cell in row:
        meet_attend[cell.value] = [sh.cell(row=cell.row, column=col).value for col in range(3, sh.max_column + 1)]


# Read attendees constraints
sh = wb["Attendees constraints"]


wb.close()