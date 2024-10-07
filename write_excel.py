from openpyxl import Workbook
from openpyxl.styles import PatternFill
import calendar
from datetime import datetime
from dateutil.relativedelta import relativedelta
import holidays


# to install: holidays, openpyxl

# Next month
next_month = datetime.now() + relativedelta(months=1)
days_in_month = calendar.monthrange(next_month.year, next_month.month)[1]





# Excel
wb = Workbook()

# Global constraints
ws = wb.active
ws.title = f"Global constraints"

# Day of the month
for day in range(1, days_in_month + 1):
    ws.cell(row=1, column=day).value = f"{day} {next_month.strftime('%B')}"

# Week day
for day in range(1, days_in_month + 1):
    weekday = calendar.weekday(next_month.year, next_month.month, day)
    weekday_name = calendar.day_name[weekday]
    ws.cell(row=2, column=day).value = weekday_name


    # Color weekend and holidays
    color = "9a0000"
    date = f"{next_month.year}-{next_month.month:02d}-{day:02d}"
    if weekday_name in {"Friday", "Saturday", "Sunday"} or date in holidays.US(years=next_month.year):
        ws.cell(row=3, column=day).fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

# Insert timeslots
ws.cell(row=4, column=1).value = "Timeslots"






# Team dayoff
ws_new = wb.create_sheet(title="Saluto")
ws_new["A3"] = "ciao"



# Salva il file
wb.save("./input_data.xlsx")

wb.close()


