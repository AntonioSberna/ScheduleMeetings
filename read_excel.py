


import openpyxl as opyxl
import utility



excel_file = "input_data_filled.xlsx"

# Read attendees names
attendees = utility.read_attendees()


wb = opyxl.load_workbook(excel_file, data_only=True)

# Read colors legend
sh = wb["Settings"]
color_holidays = sh["AS42"].fill.start_color.index[2:]
color_workday = sh["AS43"].fill.start_color.index[2:]
color_error = sh["AS44"].fill.start_color.index[2:]


# Read global constraints 
sh = wb["Global constraints"]

# I don't know if they are needed
month = sh["A2"].value.split()[0]
year = sh["A2"].value.split()[1]

# Creation of time slots
time_slots = []
for row in sh.iter_rows(min_row=5, max_row=sh.max_row, min_col=2, max_col=sh.max_column):
    for cell in row:
        col = cell.fill.start_color.index[2:]
        if col == color_workday and cell.value != "X":
            time_slots.append(sh.cell(row=3, column=cell.column).value + " / " + sh.cell(row=cell.row, column=1).value) # così è un casino, ma non so come fare meglio
        if cell.value is not None and cell.value.capitalize() != "X":
            raise ValueError(f"Error in the Global Constraints - day {sh.cell(row=3, column=cell.column).value}, timeslot {sh.cell(row=cell.row, column=1).value}")




# Read meetings and attendees
sh = wb["Groups compositions"]
meet_attend = {}
for row in sh.iter_rows(min_row=4, max_row=sh.max_row, min_col=2, max_col=2):
    for cell in row:
        meet_attend[cell.value] = [sh.cell(row=cell.row, column=col).value for col in range(3, sh.max_column + 1)]


# Read attendees constraints
sh = wb["Attendees constraints"]
attendee_constraints = {}
for row in sh.iter_rows(min_row=4, max_row=sh.max_row, min_col=3, max_col=sh.max_column):
    attendee_constraints[sh.cell(row=row[0].row, column=2).value] = []
    for cell in row:
        if cell.value is not None and cell.value.capitalize() == "X":
            attendee_constraints[sh.cell(row=cell.row, column=2).value].append(sh.cell(row=2, column=cell.column).value)

wb.close()




from ortools.sat.python import cp_model
model = cp_model.CpModel()

# Creation of DVs
meeting_slot = {}
for meeting in meet_attend.keys():
    for j in range(len(time_slots)):
        meeting_slot[(meeting, j)] = model.NewBoolVar(f'riunione_{meeting}_slot_{time_slots[j]}')

# Constraint 1: Each meeting must be assigned to exactly one slot
for meeting in meet_attend.keys():
    model.Add(sum(meeting_slot[(meeting, j)] for j in range(len(time_slots))) == 1)

# Constraint 2: A person cannot attend two meetings in the same slot
for attendee in attendees:
    for j in range(len(time_slots)):
        model.Add(sum(meeting_slot[(meeting, j)] for meeting in meet_attend
                        if attendee in meet_attend[meeting]) <= 1)

# Constraint 3: Personal constraints

solver = cp_model.CpSolver()
status = solver.Solve(model)


if status == cp_model.OPTIMAL:
    print('Meetings arrangement found:')
    for meeting in meet_attend.keys():
        for j in range(len(time_slots)):
            if solver.Value(meeting_slot[(meeting, j)]):
                print(f'{meeting} on {time_slots[j]}')
else:
    print('No optimal solution found.')
