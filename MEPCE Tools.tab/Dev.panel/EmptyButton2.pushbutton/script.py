# -*- coding: utf-8 -*-
__title__ = "Empty Button 2"
__doc__ = """Version = 1.0
Date    = 2025.05.28
_________________________________________________________________
Description:
This script will create circuits and power lighting zones.
_________________________________________________________________
How-to:
-> Draw filled regions around lighting zones
-> Click Button
-> Done!
_________________________________________________________________
Last update:
- [2025.05.28] - 1.0 RELEASE
_________________________________________________________________
To-Do:
- 
_________________________________________________________________
Author: Simeon Neese"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
# Regular + Autodesk
from Autodesk.Revit.DB import *

# pyRevit
from pyrevit import revit, forms

# .NET Imports (You often need List import)
import clr
clr.AddReference("System")
from System.Collections.Generic import List

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#==================================================
from Autodesk.Revit.UI import UIDocument
doc   = __revit__.ActiveUIDocument.Document #type: Document
uidoc = __revit__.ActiveUIDocument          #type: UIDocument
app   = __revit__.Application               #type: Application

# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
# ╠╣ ║ ║║║║║   ║ ║║ ║║║║╚═╗
# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝╚═╝
#==================================================

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ MAIN
#==================================================
# START CODE HERE
print('Template has been developed by Erik Frits.')
