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
from Snippets._filledregions import get_rooms

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
rooms = get_rooms()

if not rooms:
    alert("There are no filled regions created with the required parameters!")
    print"There are no filled regions created with the required parameters!"
    exit

#  _____
# |____ |
#     / /
#     \ \
# .___/ /
# \____/
# Get room data
roomdata = []
paramslookup = ["MEPCE Room Number","MEPCE Room Name","MEPCE HVAC Zone","Area","MEPCE Room Airflow","MEPCE Room Ventilation"]
paramsnames = ["Room Number","Room Name","HVAC Zone","Area (ft^2)","Supply Airflow (CFM)","Ventilation (CFM)"]
for room in rooms:
    roomparams = dict()
    for i in range(len(paramslookup)):
        param = paramslookup[i]
        exists = room.LookupParameter(param)
        if type(exists) is not NoneType:
            paramtype = str(exists.StorageType)
            if paramtype == "String":
                roomparams.update({paramsnames[i]: room.LookupParameter(param).AsString()})
            if paramtype == "Double":
                if i == 4 or i == 5:
                    roomparams.update({paramsnames[i]: room.LookupParameter(param).AsDouble() * 60})
                else:
                    roomparams.update({paramsnames[i]: room.LookupParameter(param).AsDouble()})
        else:
            roomparams.update({paramsnames[i]: None})
    areaflow = roomparams['Supply Airflow (CFM)']/roomparams['Area (ft^2)']
    roomparams.update({'Supply / Area (CFM/ft^2)': areaflow})
    roomdata.append(roomparams)
paramsnames.append('Supply / Area (CFM/ft^2)')

#    ___
#   /   |
#  / /| |
# / /_| |
# \___  |
#     |_/
# Format lists for excel
dataexport = []
for param in paramsnames:
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
    exit

file_name = "{} - Room Schedule.xlsx".format(date)
output_file = os.path.join(desktop_path,file_name)
workbook = xlsxwriter.Workbook(output_file, {'strings_to_numbers':False})
worksheet = workbook.add_worksheet(worksheetName)

row = 0
col = 0

for item in export:
    for it in item:
        worksheet.write(row,col,it)
        row += 1
    col += 1
    row = 0

workbook.close()

forms.alert("Schedule saved to Desktop as:\n{}.".format(file_name),title="Script complete",warn_icon=False)