#import_agenda.py
import xlrd
import sys
from db_table import db_table

sys.path.append(".")

loc = (sys.argv[0]) #through command line arguments

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

table = db_table("table_events", {"date": 'text',
                           "start_time": 'text',
                          "end_time": 'text',
                          "session_or_sub": 'text',
                          "session_title": 'text',
                          "location": 'text',
                          "description": 'text',
                          "speakers": 'text'})


index = 15
while (index < sheet.nrows):
    table.insert({"date": sheet.cell_value(index, 0),
                  "start_time": sheet.cell_value(index, 1),
                  "end_time": sheet.cell_value(index, 2),
                  "session_or_sub": sheet.cell_value(index, 3),
                  "session_title": sheet.cell_value(index, 4),
                  "location": sheet.cell_value(index, 5),
                  "description": sheet.cell_value(index, 6),
                  "speakers": sheet.cell_value(index, 7)})
    index += 1
