# -*- coding: utf-8 -*-
__title__   = "Import IES Data"
__highlight__ = "new"
__doc__     = """Version = 1.0
Date    = 2025.07.28
_________________________________________________________________
Description:
Import IES data and update filled region parameters.
_________________________________________________________________
How-to:
-> Click button
-> Select IES Room Zone Loads Excel workbook
-> If any rooms are not updated, ensure the filled region room number and room number in the excel sheet match
-> Done!
_________________________________________________________________
Last update:
- [2025.07.28] - 1.0 RELEASE
_________________________________________________________________
Author: Simeon Neese"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
import re
import clr
import math
from pyrevit import forms
from pyrevit.forms import alert

from Autodesk.Revit.DB import *

from types import NoneType

# Add reference to Microsoft Office Interop
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================
from Autodesk.Revit.UI import UIDocument
app    = __revit__.Application                  #type: Application
uidoc  = __revit__.ActiveUIDocument             #type: UIDocument
doc    = __revit__.ActiveUIDocument.Document    #type: Document

# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
# ╠╣ ║ ║║║║║   ║ ║║ ║║║║╚═╗
# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝╚═╝
#==================================================


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

#  __
# /  |
# `| |
#  | |
# _| |_
# \___/
# Import Excel
# Ask user to select Excel file
alert("Select Room Zone Loads excel report",warn_icon=False)
file_path = forms.pick_file(file_ext='xlsx')

if not file_path:
    forms.alert("No file selected.", exitscript=True)

# Ask user for starting cell
start_cell = "A10"
if not start_cell or not re.match(r"^[A-Za-z]+[0-9]+$", start_cell):
    forms.alert("Invalid cell format.", exitscript=True)

# Convert Excel-style cell reference to row and column index
def excel_cell_to_row_col(cell_ref):
    match = re.match(r"^([A-Za-z]+)([0-9]+)$", cell_ref)
    if not match:
        raise ValueError("Invalid cell reference")
    col_letters, row_number = match.groups()
    col = 0
    for char in col_letters.upper():
        col = col * 26 + (ord(char) - ord('A') + 1)
    return int(row_number), col

start_row, start_col = excel_cell_to_row_col(start_cell)

# Start Excel COM object
excel = Excel.ApplicationClass()
excel.Visible = False
excel.DisplayAlerts = False

# Open workbook
wb = excel.Workbooks.Open(file_path)
sheet = wb.Sheets[1]  # Use first worksheet

# Get used range
used_range = sheet.UsedRange
last_row = used_range.Row + used_range.Rows.Count - 1
last_col = used_range.Column + used_range.Columns.Count - 1

# Read data starting from user-defined cell
data = []
for r in range(start_row, last_row + 1):
    row_data = []
    for c in range(start_col, last_col + 1):
        value = sheet.Cells[r, c].Value2
        row_data.append(value)
    data.append(row_data)

# Close workbook and Excel
wb.Close(False)
excel.Quit()

# Release COM objects
def release(obj):
    import System
    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)

release(sheet)
release(wb)
release(excel)

#  _____
# / __  \
# `' / /'
#   / /
# ./ /___
# \_____/
# Organize data
keys = []
region_data = {}
for row in data:
        keys.append(row[2])

for x_idx, key in enumerate(keys):
        region_data[key] = {}
        region_data[key]["supply"] = data[x_idx][13]
        region_data[key]["outsideair"] = data[x_idx][30]

# region_data organized as a dictionary of dictionaries with
# the upper level keys being the room number/name and the lower level of keys being 'supply' and 'outsideair'

#  _____
# |____ |
#     / /
#     \ \
# .___/ /
# \____/
# Get all filled regions in the project

regions = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_DetailComponents).OfClass(FilledRegion).ToElements()
regiondic = {}
for region in regions:
        roomnumber = region.LookupParameter("MEPCE Room Number")
        if roomnumber.AsString() != "None":
            roomnumber = roomnumber.AsString()
            regiondic[roomnumber] = region

t = Transaction(doc,"Import IES Data")
t.Start()
for key in keys:
    keynumber = key.split(" ")
    keynumber = keynumber[0]
    try:
        keyregion = regiondic[keynumber]
    except:
        continue
    roomairflow = region_data[key]['supply'] / 60
    roomventilation = region_data[key]['outsideair'] / 60

    setroomairflow = keyregion.LookupParameter("MEPCE Room Airflow").Set(float(roomairflow))
    setroomventilation = keyregion.LookupParameter("MEPCE Room Ventilation").Set(float(roomventilation))
t.Commit()