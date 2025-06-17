# -*- coding: utf-8 -*-
__title__   = "Room Schedule"
__highlight__ = "new"
__doc__     = """Version = 1.0
Date    = 2025.06.17
_________________________________________________________________
Description:

Export a room schedule with associated zoning data.
_________________________________________________________________
How-to:
-> Ensure HVAC zones have been created using the filled
   region method
-> Click button
-> Done! Schedule has been exported to your Desktop
_________________________________________________________________
Last update:
- [2025.06.17] - 1.0 RELEASE
_________________________________________________________________
Author: Simeon Neese"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
# pyrevit
from pyrevit import revit, DB, forms, script

# libraries
import os
import sys
import xlsxwriter
from datetime import datetime

# revit
from Autodesk.Revit.DB import *
from types import NoneType

# .NET Imports (You often need List import)
import clr
from pyrevit.forms import alert

clr.AddReference("System")
from System.Collections.Generic import List

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
from Snippets._filledregions import get_rooms, get_room_data

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
# Get date
now = datetime.now()
date = now.strftime("%y%m%d")

#  _____
# / __  \
# `' / /'
#   / /
# ./ /___
# \_____/
# Get filled regions with room numbers
try:
    rooms = get_rooms()
except:
    alert("Could not get rooms. Possible causes:\n- No filled regions created\n- Parameters not imported\n- No room numbers assigned")
    print "Could not get rooms. Possible causes:\n- No filled regions created\n- Parameters not imported\n- No room numbers assigned"
    sys.exit()

if not rooms:
    alert("Could not get rooms. Possible causes:\n- No filled regions created\n- Parameters not imported\n- No room numbers assigned")
    print "Could not get rooms. Possible causes:\n- No filled regions created\n- Parameters not imported\n- No room numbers assigned"
    sys.exit()

#  _____
# |____ |
#     / /
#     \ \
# .___/ /
# \____/
# Get room data
roomdata = get_room_data(rooms)[0]
paramsnames = get_room_data(rooms)[1]

#    ___
#   /   |
#  / /| |
# / /_| |
# \___  |
#     |_/
# Format lists for excel
dataexport = []
for param in paramsnames:
    if param == "Room Number" or param == "Room Name":
        continue
    else:
        paramvals = []
        paramvals.append(param)
        for room in roomdata:
            paramvals.append(room[param])
        dataexport.append(paramvals)
export = dataexport

#  _____
# |  ___|
# |___ \
#     \ \
# /\__/ /
# \____/
# Create Excel Workbook
worksheetName = "Room Schedule"

try:
    desktop_path = os.path.join(os.path.expanduser("~"),"Desktop")
except:
    alert("Could not find Desktop path!")
    print "Could not find Desktop path!"
    sys.exit()

file_name = "{} - Room Schedule.xlsx".format(date)
output_file = os.path.join(desktop_path,file_name)
try:
    workbook = xlsxwriter.Workbook(output_file, {'strings_to_numbers':False})
except:
    alert("Could not open workbook! Possible causes:\n- Duplicate Room Schedule workbook exists on Desktop and is open")
    print "Could not open workbook! Possible causes:\n- Duplicate Room Schedule workbook exists on Desktop and is open"
    sys.exit()
worksheet = workbook.add_worksheet(worksheetName)

row = 0
col = 0

for item in export:
    for it in item:
        worksheet.write(row,col,it)
        row += 1
    col += 1
    row = 0

try:
    workbook.close()
except:
    alert("Could not open workbook! Possible causes:\n- Duplicate Room Schedule workbook exists on Desktop and is open")
    print "Could not open workbook! Possible causes:\n- Duplicate Room Schedule workbook exists on Desktop and is open"
    sys.exit()

forms.alert("Schedule saved to Desktop as:\n{}.".format(file_name),title="Script complete",warn_icon=False)