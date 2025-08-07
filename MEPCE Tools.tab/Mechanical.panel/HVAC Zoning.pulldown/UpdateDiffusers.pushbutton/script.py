# -*- coding: utf-8 -*-
__title__   = "Update Diffuser CFM"
__highlight__ = "new"
__doc__     = """Version = 1.0
Date    = 2025.07.28
_________________________________________________________________
Description:
Update airflow for all diffusers in active view. Requires zone setup with filled regions and IES data imported using HVAC Zoning buttons.
_________________________________________________________________
How-to:
-> Set up filled regions using the "Create Filled Regions" button
-> Update filled region parameters using the "Import IES Data" button
-> Click button
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
# Get Filled Regions On Level With Room Airflow Parameter
allregions = FilteredElementCollector(doc).OfClass(FilledRegion).ToElements()
regions = [region for region in allregions if doc.GetElement(region.OwnerViewId).GenLevel.Name == view_level.Name and region.LookupParameter("MEPCE Room Airflow").HasValue]
faces = get_faces([regions])[0]

t = Transaction(doc,"Update Diffuser CFM's")
t.Start()
for idx_f,face in enumerate(faces):
    roomgroups = {}
    room = regions[idx_f]
    supplyair = room.LookupParameter("MEPCE Room Airflow").AsDouble()
    exhaustair = room.LookupParameter("MEPCE Room Exhaust").AsDouble()
    ventair = room.LookupParameter("MEPCE Room Ventilation").AsDouble()
    for system in systemnames:
        roomgroup = []
        systerminals = sortedterminals[system]
        for term in systerminals:
            location = term.Location.Point
            if face.Project(location):
                roomgroup.append(term)
            else:
                continue
        termcount = len(roomgroup)
        if not termcount > 0:
            continue
        if system == "Supply Air" :
            for term in roomgroup:
                term.LookupParameter("Air Terminal Air Flow").Set(supplyair/termcount)
        if system == "Return Air":
            for term in roomgroup:
                term.LookupParameter("Air Terminal Air Flow").Set(supplyair/termcount)
                term.LookupParameter("Comments").Set("-")
        if system == "Exhaust Air" :
            for term in roomgroup:
                term.LookupParameter("Air Terminal Air Flow").Set(exhaustair/termcount)
        if system == "Outside Air" :
            for term in roomgroup:
                term.LookupParameter("Air Terminal Air Flow").Set(ventair/termcount)
t.Commit()