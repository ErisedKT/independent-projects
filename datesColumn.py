#! /usr/bin/python3
# datesColumn.py - Puts leading dates, complete months and trailing dates between entered dates 
# in a column in a spreadsheet.

import openpyxl
from datetime import date, timedelta
from calendar import month_name, monthrange

# Get start and end dates from user.

fromDate = input('FROM (dd/mm/yyyy): ')
toDate = input('TO (dd/mm/yyyy): ')

# Convert input into Date objects.

fromDate = date(int(fromDate[6:]), int(fromDate[3:5]), int(fromDate[:2]))
toDate = date(int(toDate[6:]), int(toDate[3:5]), int(toDate[:2]))

start = [fromDate.month, fromDate.year] # start of complete months
leadingDates = []   

# Find dates left in month after fromDate.

if fromDate.day != 1:   # check if it is a complete month
    try:
        # Find next month's first day.
        nextMonth = fromDate.replace(month=fromDate.month+1, day=1)
        start[0] += 1

    except ValueError:
        # If month is December, go to next year.
        if fromDate.month == 12:
            nextMonth = date(year=fromDate.year+1, month=1, day=1)
            start[0] = 1
            start[1] += 1 
        else:
            raise Exception('Invalid input!')

    # Find number of days left.       
    delta = nextMonth - fromDate
    # Add remaining days to list.
    for i in range(delta.days):
        leadingDates.append((fromDate + timedelta(days=i)).strftime('%d %B %Y'))

trailingDates = []
end = [toDate.month, toDate.year]   # end of complete months

# Find dates passed in month on toDate.

if monthrange(toDate.year, toDate.month)[1] != toDate.day:  # check if month is complete
    # Find this month's first day.
    currMonth = toDate.replace(day=1)
    # Find number of days passed in the month.
    delta = toDate - currMonth

    # Add passed days to list.
    for i in range(delta.days + 1):
        trailingDates.append((currMonth + timedelta(days=i)).strftime('%d %B %Y'))
    end[0] -= 1

months = []
initialMonth = start[0] # month of fromDate
stop = 13

# Find complete months from fromDate to toDate.

for y in range(start[1], end[1] + 1):
    if y == end[1]:
        stop = end[0] + 1

    # Add complete months to list.
    for i in range(initial, stop):
        months.append(month_name[i] + ' ' + str(y))

    initial = 1 # go back to January after year complete

# Open a workbook.

wb = openpyxl.Workbook()
sheet = wb.active
row = 1

# Put all dates and months in a single list.
column = leadingDates + months + trailingDates
# Add all data to worksheet.
for elem in column:
    sheet.cell(row=row, column=1).value = elem
    row += 1
    
# Save workbook.
wb.save('dates.xlsx')