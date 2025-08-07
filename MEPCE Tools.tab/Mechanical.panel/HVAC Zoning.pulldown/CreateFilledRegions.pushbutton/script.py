# -*- coding: utf-8 -*-
__title__   = "Create Filled Regions"
#__highlight__ = "new"
__doc__     = """Version = 1.0
Date    = 2025.07.28
_________________________________________________________________
Description:
Create filled regions for each room.
_________________________________________________________________
How-to:
-> Click button
-> Manually check filled regions to ensure they match the room boundaries, and place filled regions for any missed rooms
-> Done! Move on to zoning script
_________________________________________________________________
Last update:
- [2025.07.28] - 1.0 RELEASE
_________________________________________________________________
Author: Simeon Neese"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
# pyrevit
from pyrevit import revit, DB, forms, script

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
from Snippets._selection import get_elements_of_categories
from Snippets._filledregions import create_filled_region_from_room
from Snippets._parameters import import_all_shared_parameters

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
# Get rooms in active view level
allrooms = get_elements_of_categories(
    categories = [BuiltInCategory.OST_Rooms],
    linked = True,
    readout = False
)

active_view = doc.ActiveView
try:
    active_level = active_view.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL).AsString()
except:
    print "Could not get active view level!"
    alert(msg="Could not get active view level!",sub_msg="Ensure you have a plan view open",exitscript=True)

active_rooms = {}
for room in allrooms:
    if room.get_Parameter(BuiltInParameter.LEVEL_NAME).AsString() == active_level:
        room_number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
        room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
        room_tag = room_number + " " + room_name
        active_rooms[room_tag] = room
room_keys = active_rooms.keys()

#  _____
# / __  \
# `' / /'
#   / /
# ./ /___
# \_____/
# Get room boundaries
options = SpatialElementBoundaryOptions()
room_boundaries = {}
for room in room_keys:
    room_boundaries[room] = active_rooms[room].GetBoundarySegments(options)

#  _____
# |____ |
#     / /
#     \ \
# .___/ /
# \____/
# Import Shared Parameters
shared_param_path = "R:\\05-Studios\\03-Education\\3_Dynamo\\Create HVAC Zones & Update Diffuser CFM\\Shared Params\\FilledRegions.txt"
shared_params = import_all_shared_parameters(
    categories      = [BuiltInCategory.OST_DetailComponents],
    pgroup          = BuiltInParameterGroup.PG_GENERAL,
    param_file_path = shared_param_path,
    instance        = True
)

#    ___
#   /   |
#  / /| |
# / /_| |
# \___  |
#     |_/
# Create filled regions (with room name and number updated)
region_type = None
region_types = FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements()
if region_types:
    region_type = region_types[0]
else:
    print "⚠️ No FilledRegionType found in document."
    alert(msg="⚠️ No FilledRegionType found in document.",exitscript=True)

t = Transaction(doc,"Create Filled Region from Room Boundary")

t.Start()
for key in room_keys:
    room = active_rooms[key]
    room_number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
    room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
    region = create_filled_region_from_room(room,active_view,region_type)

    if region is False:
        continue
    else:
        regionkeys = region.keys()
        for key in regionkeys:
            subregion = region[key]
            nameparam = subregion.LookupParameter("MEPCE Room Name")
            nameparam.Set(room_name)
            numberparam = subregion.LookupParameter("MEPCE Room Number")
            numberparam.Set(room_number)
t.Commit()