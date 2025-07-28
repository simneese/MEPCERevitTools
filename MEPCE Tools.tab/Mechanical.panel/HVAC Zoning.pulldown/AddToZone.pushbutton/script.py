# -*- coding: utf-8 -*-
__title__   = "Add Regions To Zones"
__highlight__ = "new"
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
from Snippets._math import get_lowest_available, create_color_grad

def color_picker(index):
    colorsets = [
        [Color(104,14,14),Color(255,0,0)],      # Reds
        [Color(20,20,77),Color(0,0,255)],       # Blues
        [Color(69,43,12),Color(252,140,3)],     # Oranges
        [Color(34,69,23),Color(65,252,3)],      # Greens
        [Color(74,10,69),Color(242,5,222)],     # Purples
        [Color(22,79,77),Color(0,252,242)]      # Cyans
    ]

    return colorsets[index]
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
    # Get all existing zones organized by zone name
    allexistingprefixes = []
    allexistingzones = {}
    for region in filledregions:
        regionzone = region.LookupParameter("MEPCE HVAC Zone").AsString()
        if regionzone:
            zoneparts = regionzone.split('-')
            del zoneparts[-1]
            deriveprefix = ""
            for part in zoneparts:
                deriveprefix = deriveprefix + part + "-"
            allexistingprefixes.append(deriveprefix)
            try:
                lookup = allexistingzones[regionzone]
            except:
                allexistingzones[regionzone] = []
            allexistingzones[regionzone].append(region)

    # Get unique prefixes
    alluniqueprefixes = list(set(allexistingprefixes))
    alluniqueprefixes.sort()

    # Get the Solid Fill pattern (used for solid fills in Revit)
    solid_fill = None
    collector = DB.FilteredElementCollector(revit.doc).OfClass(FillPatternElement)
    for fpe in collector:
        pattern = fpe.GetFillPattern()
        if pattern.IsSolidFill:
            solid_fill = fpe
            break

    # Assign colors per zone group
    for z_idx,zoneprefix in enumerate(alluniqueprefixes):
        color = color_picker(z_idx)
        keyswithpre = []
        for key in allexistingzones.keys():
            split = key.split("-")
            del split[-1]
            keyprefix = ""
            for part in split:
                keyprefix = keyprefix + part + "-"
            if keyprefix == zoneprefix:
                keyswithpre.append(key)

        zonecount = len(keyswithpre)
        gradient = create_color_grad(color,zonecount)

        for keystocolor in keyswithpre:
            if len(gradient) != len(keyswithpre):
                alert("Error creating color gradients for zone group {}!".format(zoneprefix))
                continue

            for c_idx,zone in enumerate(keyswithpre):
                for region in allexistingzones[zone]:
                    ogs = OverrideGraphicSettings()
                    icolor = gradient[c_idx]
                    ogs.SetProjectionLineColor(Color(0,0,0))
                    ogs.SetSurfaceForegroundPatternColor(icolor)
                    ogs.SetSurfaceForegroundPatternId(solid_fill.Id)
                    ogs.SetSurfaceBackgroundPatternId(DB.ElementId.InvalidElementId)
                    active_view.SetElementOverrides(region.Id,ogs)


    t.Commit()

    createdzones.append(zone)

alert("Finished!",sub_msg="Created New Zones:\n{}".format(createdzones),warn_icon=False)