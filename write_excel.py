import openpyxl as opyxl

import calendar
from datetime import datetime
from dateutil.relativedelta import relativedelta
import holidays

def read_attendees(filename="attendees.xlsx"):
    wb = opyxl.load_workbook(filename)
    ws = wb.active

    attendees = []
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=3, max_col=3):
        for cell in row:
            attendees.append(cell.value)
    return attendees
    

# to install: holidays, openpyxl

# Next month
next_month = datetime.now() + relativedelta(months=1)
days_in_month = calendar.monthrange(next_month.year, next_month.month)[1]

# Slot ranges
slot_ranges = [[7, 12], [14, 16]]
slot_duration = 1

# Weekly day off
weekly_day_off = {"Friday", "Saturday", "Sunday"}

time_slots = []
for start_hour, end_hour in slot_ranges:
    current_time = start_hour
    while current_time < end_hour:
        hours = int(current_time)
        minutes = int((current_time - hours) * 60)
        time_slots.append(f"{hours:02d}:{minutes:02d}")
        current_time += slot_duration



# Excel
wb = opyxl.Workbook()

# Global constraints
ws = wb.active
ws.title = f"Global constraints"

color_holidays = "9a0000"
color_workday = "009a00"
column_width = 6.75

ws.cell(row=2, column=1).value = f"{next_month.strftime('%B')} {next_month.year}"
ws.cell(row=2, column=1).font = opyxl.styles.Font(bold=True)

for day in range(1, days_in_month + 1):
    ws.column_dimensions[opyxl.utils.get_column_letter(day + 1)].width = column_width

    # Day of the month
    ws.cell(row=3, column=day+1).value = f"{day} {next_month.strftime('%b')}"
    ws.cell(row=3, column=day+1).alignment = opyxl.styles.Alignment(horizontal="center", vertical="center")
    ws.cell(row=3, column=day+1).font = opyxl.styles.Font(bold=True)
    

    # Week day
    ws.cell(row=4, column=day+1).value = datetime(next_month.year, next_month.month, day).strftime('%a') #short weekday name 
    ws.cell(row=4, column=day+1).alignment = opyxl.styles.Alignment(horizontal="center", vertical="center")
    ws.cell(row=4, column=day+1).font = opyxl.styles.Font(bold=True)
    
    # Color weekend and holidays
    start_row = 4
    date = f"{next_month.year}-{next_month.month:02d}-{day:02d}"
    for n_time_slot in range(len(time_slots) + 1):
        if calendar.day_name[calendar.weekday(next_month.year, next_month.month, day)] in weekly_day_off or date in holidays.US(years=next_month.year): #weekday >= 5
            ws.cell(row=start_row+n_time_slot, column=day+1).fill = opyxl.styles.PatternFill(start_color=color_holidays, end_color=color_holidays, fill_type="solid")
        else:
            ws.cell(row=start_row+n_time_slot, column=day+1).fill = opyxl.styles.PatternFill(start_color=color_workday, end_color=color_workday, fill_type="solid")
    

# Insert timeslots
ws.column_dimensions["A"].width = 11.5
for i, time_slot in enumerate(time_slots, start=1):
    ws.cell(row=start_row + i, column=1).value = f"{time_slot} - {time_slots[i] if i < len(time_slots) else 'End'}"

# Conditional formatting
red_fill = opyxl.styles.PatternFill(start_color=color_holidays, end_color=color_holidays, fill_type="solid")
gray_fill = opyxl.styles.PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
rule_red = opyxl.formatting.rule.CellIsRule(operator='equal', formula=['"X"'], fill=red_fill)
rule_gray = opyxl.formatting.rule.FormulaRule(formula=['AND(B5<>"X", B5<>"")'], fill=gray_fill)
ws.conditional_formatting.add(f"B5:{opyxl.utils.get_column_letter(days_in_month + 1)}{len(time_slots) + 4}", rule_red)
ws.conditional_formatting.add(f"B5:{opyxl.utils.get_column_letter(days_in_month + 1)}{len(time_slots) + 4}", rule_gray)


thin_side = opyxl.styles.Side(border_style="thin", color="000000")
border = opyxl.styles.Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
for row in ws[f"A3:{opyxl.utils.get_column_letter(days_in_month + 1)}{len(time_slots) + 4}"]:
    for cell in row:
        cell.border = border



# Team dayoff
attendees = read_attendees()
print(attendees)
ws_new = wb.create_sheet(title="Attendees contraints")
ws_new["A3"] = "Name"



# Salva il file
wb.save("./input_data.xlsx")

wb.close()


