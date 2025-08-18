# -*- coding: utf-8 -*-
__title__   = "Update Diffuser CFM"
__highlight__ = "new"
__doc__     = """Version = 1.1
Date    = 2025.08.14
_________________________________________________________________
Description:
Update airflow for all diffusers in active view. Requires zone setup with filled regions and IES data imported using HVAC Zoning buttons.
_________________________________________________________________
How-to:
-> Set up filled regions using the "Create Filled Regions" button
-> Update filled region parameters using the "Import IES Data" button and the "Calculate Exhaust" button
-> Click button
-> Select systems to update
-> Enter whole number to round CFM value up to
-> Done!
_________________________________________________________________
Last update:
- [2025.08.14] - 1.1 Added rounding
- [2025.08.07] - 1.0 RELEASE
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

from rpw.ui.forms import (FlexForm, Label, TextBox, Separator, Button)

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
from Snippets._filledregions import get_faces

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

DEBUGMODE = False

#  __
# /  |
# `| |
#  | |
# _| |_
# \___/
# Get all air terminals in view
active_view = doc.ActiveView
view_level = active_view.GenLevel

terminals = FilteredElementCollector(doc,active_view.Id).OfCategory(BuiltInCategory.OST_DuctTerminal).ToElements()
terminals = [terminal for terminal in terminals if not "No Plenum" in terminal.LookupParameter("Family and Type").AsValueString()]

if not terminals:
    alert("No air terminals in the active view!",
          sub_msg="Only air terminals in the active view will be updated. Exiting Script.",
          exitscript=True,
          title="Update Diffusers"
          )

sortedterminals = {}
for terminal in terminals:
    systemclass = terminal.LookupParameter("System Classification").AsString()
    if systemclass == "Supply Air":
        connectors = terminal.MEPModel.ConnectorManager.Connectors
        for connector in connectors:
            if connector.MEPSystem:
                systemtype = connector.MEPSystem.LookupParameter("Type").AsValueString()
                systemclass = systemtype
    if systemclass not in sortedterminals:
        sortedterminals[systemclass] = [terminal]
    else:
        sortedterminals[systemclass].append(terminal)
systemnames = sortedterminals.keys()

#  _____
# / __  \
# `' / /'
#   / /
# ./ /___
# \_____/
# User input system types
selsystems = forms.SelectFromList.show(
    systemnames,
    multiselect = True,
    title='Select system types to update',
    button_name='Select'
)

if not selsystems:
    alert("No systems selected. Exiting script.",exitscript=True)

components = [Label('Round to the nearest:'), TextBox('rounding'),
              Separator(), Button('Continue')]
form = FlexForm('Input Rounding', components)
form.show()
user_inputs = form.values
rounding = float(user_inputs['rounding'])/60

#  _____
# |____ |
#     / /
#     \ \
# .___/ /
# \____/
# Get Filled Regions On Level With Room Airflow Parameter
allregions = FilteredElementCollector(doc).OfClass(FilledRegion).ToElements()
regions = []
for region in allregions:
    if type(doc.GetElement(region.OwnerViewId).GenLevel) is NoneType:
        continue
    else:
        regions.append(region)

regions = [region for region in regions if doc.GetElement(region.OwnerViewId).GenLevel.Name == view_level.Name]
faces = get_faces([regions])[0]

t = Transaction(doc,"Update Diffuser CFM's")
t.Start()
for idx_f,face in enumerate(faces):
    roomgroups = {}
    room = regions[idx_f]
    supplyair = room.LookupParameter("MEPCE Room Airflow").AsDouble()
    exhaustair = room.LookupParameter("MEPCE Room Exhaust").AsDouble()
    ventair = room.LookupParameter("MEPCE Room Ventilation").AsDouble()
    update = room.LookupParameter("MEPCE Update Region").AsInteger()
    number = room.LookupParameter("MEPCE Room Number").AsString()
    if not number:
        continue
    if update == 0:
        continue
    if DEBUGMODE is True:
        print "\nRoom: {} {}".format(number, room.LookupParameter("MEPCE Room Name").AsString())
        print "Supply: {} CFM\nVentilation: {} CFM\nExhaust: {} CFM".format(supplyair, ventair, exhaustair)
    for system in selsystems:
        roomgroup = []
        systerminals = sortedterminals[system]
        for term in systerminals:
            location = term.Location.Point
            if face.Project(location):
                roomgroup.append(term)
            else:
                continue
        termcount = len(roomgroup)
        if DEBUGMODE is True:
            print "{} Terminals: {}".format(system,termcount)
        if not termcount > 0:
            continue
        if system == "Supply Air" :
            for term in roomgroup:
                term.LookupParameter("Air Terminal Air Flow").Set(math.ceil(supplyair/termcount/rounding)*rounding)
        if system == "Return Air":
            for term in roomgroup:
                term.LookupParameter("Air Terminal Air Flow").Set(math.ceil(supplyair/termcount/rounding)*rounding)
                term.LookupParameter("Comments").Set("-")
        if system == "Exhaust Air" :
            for term in roomgroup:
                term.LookupParameter("Air Terminal Air Flow").Set(math.ceil(exhaustair/termcount/rounding)*rounding)
        if system == "Outside Air" :
            for term in roomgroup:
                term.LookupParameter("Air Terminal Air Flow").Set(math.ceil(ventair/termcount/rounding)*rounding)
t.Commit()