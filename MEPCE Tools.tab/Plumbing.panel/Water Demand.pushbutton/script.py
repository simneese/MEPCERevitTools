# -*- coding: utf-8 -*-
__title__ = "Water Demand"
#__highlight__ = "new"
__doc__ = """Version = 1.0
Date    = 2025.07.24
_________________________________________________________________
Description:
Counts plumbing fixtures and exports to Waste and Water Demand Calc excel workbook.
_________________________________________________________________
How-to:
-> Click button
-> Either manually select fixtures or automatically select all fixtures in the project
-> Ensure that Excel is closed
-> Select Waste and Water Demand Calc excel workbook for the project
-> Done!
_________________________________________________________________
Last update:
- [2025.07.24] - 1.0 RELEASE
_________________________________________________________________
Author: Michael Opperman"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import *
from pyrevit import revit, forms, script
from pyrevit.forms import alert

import os
from collections import defaultdict
from types import NoneType

# Excel Interop (.NET) setup
from System.Runtime.InteropServices import Marshal
import clr
clr.AddReference('Microsoft.Office.Interop.Excel')
import Microsoft.Office.Interop.Excel as Excel

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#==================================================
from Autodesk.Revit.UI import UIDocument
from collections import defaultdict

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application


# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
# ╠╣ ║ ║║║║║   ║ ║║ ║║║║╚═╗
# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝╚═╝
#==================================================
from Snippets._selection import get_elements_of_categories


def is_workbook_open(filepath):
    try:
        # Try to connect to existing Excel instance
        excel_app_type = System.Type.GetTypeFromProgID("Excel.Application")
        excel_app = System.Activator.GetObject(None, "Excel.Application")

        # Check open workbooks
        for wb in excel_app.Workbooks:
            if wb.FullName.lower() == file_path.lower():
                return True
    except Exception:
        pass  # Excel probably not running, or COM error
    return False

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
# Get plumbing fixtures
selection_mode = forms.alert(
    msg="Select plumbing fixtures manually or process entire project?",
    options=["Manual Selection", "Full Project"],
    title="Selection Mode",
    warn_icon=False
)

if selection_mode == "Manual Selection":
    try:
        # Prompt user to manually select plumbing fixtures in the model
        alert("Draw Rectangle Over Desired Area",warn_icon=False)
        selected_elements = uidoc.Selection.PickElementsByRectangle()
        plumbing_fixtures = [el for el in selected_elements if el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_PlumbingFixtures)]

        
    except:
        alert("No elements selected or selection canceled.", exitscript=True)
else:
    # Full project plumbing fixture collection
    plumbing_fixtures = get_elements_of_categories(
        categories=[BuiltInCategory.OST_PlumbingFixtures],
        systemtypes=None,
        view=None,
        readout=False
    )

if not plumbing_fixtures:
    alert("No plumbing fixtures found in project. Exiting script.", exitscript=True)
    
#print("Fixture Counts:", dict(fixture_counts))

#  _____
# / __  \
# `' / /'
#   / /
# ./ /___
# \_____/
# Excel export path
alert("Close Excel before continuing.")
alert("Select location of Water and Waste Demand Calculator.",warn_icon=False)
export_path = forms.pick_file(file_ext='xlsx', title='Select Excel File to Update')

if not export_path or not os.path.exists(export_path):
    alert("No valid Excel file selected. Exiting.", exitscript=True)

# Launch Excel via COM
excel_app = Excel.ApplicationClass()
excel_app.Visible = False
excel_app.DisplayAlerts = False

# Open workbook
workbook = excel_app.Workbooks.Open(export_path)

# Use the first worksheet (assume user will control where T7 lives)
summary_sheet = workbook.Sheets.Item["Dynamo"]

# Define output start positions (T7 = col 20, row 7)
start_row = 7
col_tag = 20     # Column T
col_qty = 21     # Column U

# Build quantity dictionary from Revit plumbing fixtures
fixture_counts = defaultdict(int)

for fixture in plumbing_fixtures:
    try:
        symbol = fixture.Symbol
        if symbol is None:
            continue

        type_comments_param = symbol.LookupParameter("Type Comments")
        type_comment = type_comments_param.AsString() if type_comments_param and type_comments_param.HasValue else "None"

        fixture_counts[type_comment] += 1
    except Exception as e:
        print("Error processing fixture ID {}: {}".format(fixture.Id, str(e)))
        continue

# Write to Excel starting at T7 and U7
current_row = start_row
for tag, qty in sorted(fixture_counts.items()):
    summary_sheet.Cells(current_row, col_tag).Value2 = tag
    summary_sheet.Cells(current_row, col_qty).Value2 = qty
    current_row += 1

# Save and close
workbook.Save()
workbook.Close(False)
excel_app.Quit()

# Release COM objects
Marshal.ReleaseComObject(summary_sheet)
Marshal.ReleaseComObject(workbook)
Marshal.ReleaseComObject(excel_app)

alert("Excel file updated successfully.",warn_icon=False)