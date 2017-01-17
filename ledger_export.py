#!/bin/python
#
# This script parses the 'Bankbuch' / 'Kassenbuch' as used by the German Pirate
# Party
#
# To generate the plaintext / html table, simply run this script with the `.ods`
# file as the first argument:
#
# > python ./ledger_export.py Vorlage_Bankbuch.ods

import sys
from pyexcel_ods import get_data

# read the first sheet
data = get_data(sys.argv[1])
sheet0 = data[list(data.keys())[0]]

# let's define some html for the table
table_head = '''
<table style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
<thead>
<tr>
    <th>Datum</th>
    <th>Beschreibung</th>
    <th style="text-align: right;">Betrag</th>
    <th style="text-align: right;">Saldo</th>
</tr>
</thead>
<tbody>
'''

table_row = '''<tr>
    <td>{date:%d.%m.%Y}</td>
    <td>{title}</td>
    <td style="text-align: right;{style}">{amount:.2f} €</td>
    <td style="text-align: right;">{balance:.2f} €</td>
</tr>\n'''

table_footer = '''
</tbody>
</table>
'''

table_body = ''

# iterate over the table rows and sum up the balance
balance = 0

for i in range(3, len(sheet0)):
    if sheet0[i][2] == '': continue;
    date = sheet0[i][2]
    title = sheet0[i][5]
    amount = sheet0[i][6]
    balance += amount
    balance = round(balance, 2)
    style = 'color: #b00;' if (amount<0) else ''

    print(date, title, amount, balance)
    table_body = table_row.format(date=date, title=title, amount=amount, balance=balance, style=style) + table_body

print(table_head + table_body + table_footer)
