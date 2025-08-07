# -*- coding: utf-8 -*-
__title__   = "Add Regions To Zones"
#__highlight__ = "new"
__doc__     = """Version = 1.0
Date    = 2025.07.28
_________________________________________________________________
Description:
Adds selected filled regions to a zone
_________________________________________________________________
How-to:
-> Click button
-> Enter zone prefix
-> Select regions to add to a zone
-> Continue until finished adding regions to zones
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
from Snippets._selection import select_multiple
from Snippets._math import get_lowest_available, update_zone_colors

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
# Ask for naming scheme
alert(msg='Input prefix for zone naming scheme',sub_msg='i.e.\nVAV-1-2- will result in zones named VAV-1-2-1, VAV-1-2-2, etc.',warn_icon=False)
from rpw.ui.forms import (FlexForm, Label, TextBox, Separator, Button)

components =    [Label('Prefix:'), TextBox('prefix'), Separator(), Button('Apply')]

form = FlexForm('Zone Naming', components)
form.show()

user_inputs = form.values
prefix      = user_inputs['prefix']

if prefix == "":
    alert("Prefix cannot be empty!",exitscript=True)

#  _____
# / __  \
# `' / /'
#   / /
# ./ /___
# \_____/
# Select Filled Regions and Apply Zone Names
active_view = doc.ActiveView

zonenum = 1
createdzones = []

t = Transaction(doc,"Add Regions to Zone")

while True:
    # Get pre-existing zone numbers with same prefix
    filledregions = FilteredElementCollector(doc, active_view.Id).OfCategory(
        BuiltInCategory.OST_DetailComponents).OfClass(FilledRegion).ToElements()

    existingzonenums = []
    for region in filledregions:
        regionzone = region.LookupParameter("MEPCE HVAC Zone").AsString()
        if regionzone and prefix in regionzone:
            zoneparts = regionzone.split('-')
            zonenum = int(zoneparts[-1])
            existingzonenums.append(zonenum)

    # Get next available zone number
    if existingzonenums:
        nextzone = get_lowest_available(existingzonenums)
    else:
        nextzone = 1

    zone = prefix + str(nextzone)

    # Check to continue
    check = alert("Create new zone {}?".format(zone),sub_msg="If yes, select regions to add to zone.",warn_icon=False,options=['Yes','No'])
    if check == 'No':
        break

    t.Start()

    # Update HVAC Zone parameter value
    regions = select_multiple([BuiltInCategory.OST_DetailComponents],"Select Filled Regions")
    if not regions:
        t.Commit()
        break
    ZONEPARAMS = [region.LookupParameter("MEPCE HVAC Zone").Set(zone) for region in regions]

    #  _____
    # |____ |
    #     / /
    #     \ \
    # .___/ /
    # \____/
    # Update zone colors
    update_zone_colors(active_view,filledregions)

    t.Commit()

    createdzones.append(zone)

alert("Finished!",sub_msg="Created New Zones:\n{}".format(createdzones),warn_icon=False)