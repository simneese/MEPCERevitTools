# -*- coding: utf-8 -*-
__title__   = "Halftone Elements"
__highlight__ = "new"
__doc__     = """Version = 1.0
Date    = 2025.08.28
________________________________________________________________
Description:
It halftones stuff
________________________________________________________________
How-to:
-> Select Elements to Halftone
-> Push button
-> Done!
________________________________________________________________
Last update:
- [2025.08.28] - 1.0 RELEASE
________________________________________________________________
Author: Michael Opperman"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
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
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================
app    = __revit__.Application
uidoc  = __revit__.ActiveUIDocument
doc    = __revit__.ActiveUIDocument.Document #type:Document
view = doc.ActiveView

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
# Draw your rectangle !

# forms.alert("Draw Rectangle Over Desired Area",warn_icon=False)
# selected_elements = uidoc.Selection.PickElementsByRectangle()

#  _____
# / __  \
# `' / /'
#   / /
# ./ /___
# \_____/
# Half tone 

# Get currently selected elements
selection = uidoc.Selection.GetElementIds()

if not selection or len(selection) == 0:
    forms.alert("No elements selected. Please select elements first.",warn_icon=False)

# Define the half-tone override
override = OverrideGraphicSettings()
override.SetHalftone(True)

# Apply overrides
with revit.Transaction("Apply Half-Tone to Selected Elements"):
    for el_id in selection:
        view.SetElementOverrides(el_id, override)



