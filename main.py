"""
Hello World! Project started 12/3/2015 by Bret Lien
First goal will be to create some useful .xlsx spreadsheets using something like XlsxWriter
12/6/2015: Finished grinding it this weekend - the loan tab is complete
"""
import datetime
import xlsxwriter

wb = xlsxwriter.Workbook('/home/matroska/testing/test.xlsx')
devProfileTab = wb.add_worksheet('Dev Profile')
loanTab = wb.add_worksheet('Loan Amort.')
leasesTab = wb.add_worksheet('Leases')
leaseExpirationsTab = wb.add_worksheet('Lease Expirations')

number_format = wb.add_format({'num_format': 0x01, 'font_name': 'Arial', 'font_size': 10})
money_format = wb.add_format({'num_format': 0x07, 'font_name': 'Arial', 'font_size': 10})
percent_format = wb.add_format({'num_format': 0x0a, 'font_name': 'Arial', 'font_size': 10})
date_format = wb.add_format({'num_format': 0x0e, 'font_name': 'Arial', 'font_size': 10})

# Dev Profile area
devProfileTab.set_first_sheet()
devProfileTab.set_column('A:E', 20)
devProfileTab.write_number('D7', 13500000, money_format)
# end Dev Profile Tab area

# Loan Tab area
loanTab.activate()

commencement_date = datetime.date(2016, 1, 1)
amort_period = 25
loan_term = 10
int_rate1 = .045

table_row_offset = 8

loanTab.set_column('A:H', 10)
loanTab.set_row(0, 25)
loanTab.set_row(1, 13)
loanTab.set_row(2, 13)

loantitle_format = wb.add_format({
    'font_name': 'Baskerville Old Face',
    'font_size': 18,
    'bold': True,
    'align': 'center',
    'valign': 'vcenter'
})

loanlabel_format = wb.add_format({
    'font_name': 'Arial',
    'font_size': 10,
    'align': 'right'
})

loanlabel_format3 = wb.add_format({
    'font_name': 'Arial',
    'font_size': 10,
    'text_wrap': True,
    'align': 'center',
    'valign': 'vcenter',
    'bottom': True
})

loanTab.merge_range('A1:G1', 'AMORTIZATION SCHEDULE', loantitle_format)
loanTab.write_formula('C4', '=\'Dev Profile\'!D7', money_format)

labelset1 = ('Principal:', 'Terms (Years)', 'Annual Rate:', 'Payment:', 'Debt Service:')
labelset2 = ('Commencement Date', 'Month:', 'Day:', 'Year:')
labelset3 = ('Date of Payment', 'Interest', 'Principal', 'Total Payment', 'Ending Principal',
             'Annual Principal', 'Annual Interest', 'Total Annual')

for i, label in enumerate(labelset1):
    loanTab.write(3 + i, 1,  label, loanlabel_format)

for i, label in enumerate(labelset2):
    loanTab.write(3 + i, 4,  label, loanlabel_format)

for i, label in enumerate(labelset3):
    loanTab.merge_range(table_row_offset, i, table_row_offset + 1, i, label, loanlabel_format3)

loanTab.write_datetime('F4', commencement_date, date_format)
loanTab.write_formula('F5', '=MONTH(F4)', number_format)
loanTab.write_formula('F6', '=DAY(F4)', number_format)
loanTab.write_formula('F7', '=YEAR(F4)', number_format)

loanTab.write_number('C5', amort_period, number_format)
loanTab.write_number('C6', int_rate1, percent_format)
loanTab.write_formula('C7', '=PMT(C6/12, C5*12, -C4)', money_format)
loanTab.write_formula('C8', '=+C7*12', money_format)

for i in range(0, loan_term * 12):
    loanTab.write_formula(table_row_offset + 3 + i, 0, ('=EDATE(A' +
                                                        str(table_row_offset + 3 + i) + ', 1)'), date_format)
    loanTab.write_formula(table_row_offset + 3 + i, 1, ('=ROUND(E' +
                                                        str(table_row_offset + 3 + i) + '*C6/12,2)'), money_format)
    loanTab.write_formula(table_row_offset + 3 + i, 2, ('=IF((E' + str(table_row_offset + 3 + i) +
                                                        '-(C$7-B' + str(table_row_offset + 4 + i) +
                                                        '))<0,C$7-B' + str(table_row_offset + 4 + i) +
                                                        '+(E' + str(table_row_offset + 3 + i) +
                                                        '-(C$7-B' + str(table_row_offset + 4 + i) +
                                                        ')),C$7-B' + str(table_row_offset + 4 + i) + ')'), money_format)
    loanTab.write_formula(table_row_offset + 3 + i, 3, ('=B' +
                                                        str(table_row_offset + 4 + i) + '+C' +
                                                        str(table_row_offset + 4 + i)), money_format)
    loanTab.write_formula(table_row_offset + 3 + i, 4, ('=E' + str(table_row_offset + 3 + i) + '-C' +
                                                        str(table_row_offset + 4 + i)), money_format)
    if ((i + 1) % 12 == 0) and i > 0:
        loanTab.write_formula(table_row_offset + 3 + i, 5, ('=SUM(C' + str(table_row_offset + 4 - 11 + i) + ':C' +
                                                            str(table_row_offset + 4 + i) + ')'), money_format)
        loanTab.write_formula(table_row_offset + 3 + i, 6, ('=SUM(B' + str(table_row_offset + 4 - 11 + i) + ':B' +
                                                            str(table_row_offset + 4 + i) + ')'), money_format)
        loanTab.write_formula(table_row_offset + 3 + i, 7, ('=SUM(F' + str(table_row_offset + 4 + i) + ':G' +
                                                            str(table_row_offset + 4 + i) + ')'), money_format)


loanTab.write_formula(table_row_offset + 2, 4, '=C4', money_format)
loanTab.write_formula(table_row_offset + 3, 0, '=F4', date_format)


# end Loan Tab area

# start Leases Tab area
# end Leases Tab area

# start Lease Expirations Tab area
# end Lease Expirations Tab area


wb.close()
