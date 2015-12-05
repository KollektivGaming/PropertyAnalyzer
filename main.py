"""
Hello World! Project started 12/3/2015 by Bret Lien
First goal will be to create some useful .xlsx spreadsheets using something like XlsxWriter
"""
from datetime import datetime

import xlsxwriter

wb = xlsxwriter.Workbook('/home/matroska/testing/test.xlsx')
loanTab = wb.add_worksheet('Loan Amort.')
leasesTab = wb.add_worksheet('Leases')
leaseExpirationsTab = wb.add_worksheet('Lease Expirations')

bold = wb.add_format({'bold': True})
bolditalic = wb.add_format({'bold': True, 'italic': True})
money_format = wb.add_format({'num_format': '$#,##0'})
date_format = wb.add_format({'num_format': 'mmmm d yyyy'})

loanTab.set_column('B:B', 20)

loanTab.write('A1', 'Loan Amount', bold)
loanTab.write('B1', 'Date', bolditalic)
loanTab.write('C1', 'Rate', bolditalic)

expenses = (
    ['Rent', '2013-01-13', 1000],
    ['Gas', '2013-01-14', 100],
    ['Food', '2013-02-16', 300],
    ['Gym', '2013-01-20', 50],
)

row = 1
col = 0

for item, buydate, cost in expenses:
    date = datetime.strptime(buydate, "%Y-%m-%d")

    loanTab.write_string(row, col, item)
    loanTab.write_datetime(row, col + 1, date, date_format)
    loanTab.write_number(row, col + 2, cost, money_format)
    row += 1

loanTab.write_string(row, 0, 'Total')
loanTab.write_formula(row, 2, '=SUM(C2:C5)')


wb.close()
