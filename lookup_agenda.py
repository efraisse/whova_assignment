#lookup_agenda.py
import xlrd
import sys

loc = ("agenda.xls")

sys.path.append(".")

from db_table import db_table
from import_agenda

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

args1, args2 = sys.argv[0], sys.argv[1] #through command line arguments

print(table.select(where = {args1: args2}))

cols = dict()
cols["date"] = 0
cols["start_time"] = 1
cols["end_time"] = 2
cols["session_or_sub"] = 3
cols["session_time"] = 4
cols["location"] = 5
cols["description"] = 6
cols["speakers"] = 7

index = 15
session_index = 0
while (index < sheet.nrows):
    if args2 in sheet.cell_value(index, cols[args1]):
        session_index = index;
        break;
    index += 1


while (session_index < sheet.nrows):
    if (sheet.cell_value(session_index + 1, cols["session_or_sub"]) == "Sub"):
        des = sheet.cell_value(session_index + 1, cols["description"])
        print(table.select(where = {"description": des})) #I use description because it contains the most characters
        session_index += 1
    else:
        break;
