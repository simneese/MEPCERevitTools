# -*- coding: utf-8 -*-
__title__ = "Pipe Volumes"
__doc__ = """Version = 1.0
Date    = 2025.07.22
_________________________________________________________________
Description:
Calculates the total pipe volume for each piping system in the project, then exports the results as an excel workbook on the user's desktop.
_________________________________________________________________
Last update:
- [2025.07.22] - 1.0 RELEASE
_________________________________________________________________
Author: Michael Opperman
"""
# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
# Regular + Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import PipingSystemType
from Autodesk.Revit.UI import UIDocument
import xlsxwriter
import os
from types import NoneType

from pyrevit import revit, forms
from pyrevit.forms import alert

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#==================================================
from Autodesk.Revit.UI import UIDocument
from Snippets._selection import get_elements_of_categories

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application

# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
# ╠╣ ║ ║║║║║   ║ ║║ ║║║║╚═╗
# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝╚═╝
#==================================================
# Calculate volume using pipe diameter and length
def get_pipe_volume(pipe):
    try:
        length = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()
        diameter = pipe.Diameter  # diameter in feet
        radius = diameter / 2.0
        return 3.14159265359 * radius * radius * length
    except:
        return 0.0

# Extract base system name, e.g. "CHWR 1" -> "CHWR"
def get_base_system_name(name):
    if not name:
        return "Unassigned"
    return name.split()[0] if " " in name else name

# Collect all pipe elements in the project
pipe_elements = get_elements_of_categories(
    categories=[BuiltInCategory.OST_PipeCurves]
)

if not pipe_elements:
    alert("No pipe elements found in project. Exiting script.", exitscript=True)

# Group volume and elements by base system type
volume_by_base_system = {}
elements_by_base_system = {}

for pipe in pipe_elements:
    try:
        sys_name = pipe.MEPSystem.LookupParameter("Type").AsValueString()
        if not sys_name:
            sys_name = mep.Name
    except:
        sys_name = "Unassigned"

    vol = get_pipe_volume(pipe)

    volume_by_base_system[sys_name] = volume_by_base_system.get(sys_name, 0.0) + vol

    if sys_name not in elements_by_base_system:
        elements_by_base_system[sys_name] = []
    elements_by_base_system[sys_name].append(pipe)

# Excel export path
export_dir = os.path.expanduser("~\\Desktop\\")
export_path = os.path.join(export_dir, "Pipe System Volumes.xlsx")

# Create workbook
workbook = xlsxwriter.Workbook(export_path)
summary_sheet = workbook.add_worksheet("Volume Summary")
details_sheet = workbook.add_worksheet("Pipe Details")

# Write volume summary
summary_sheet.write("A1", "System Type")
summary_sheet.write("B1", "Volume (ft³)")
summary_sheet.write("C1", "Volume (gal)")

row = 2
for system, ft3 in sorted(volume_by_base_system.items()):
    gal = ft3 * 7.48052
    summary_sheet.write("A{}".format(row), system)
    summary_sheet.write("B{}".format(row), round(ft3, 2))
    summary_sheet.write("C{}".format(row), round(gal, 2))
    row += 1

# Write pipe details
details_sheet.write("A1", "System Type")
details_sheet.write("B1", "Pipe ID")
details_sheet.write("C1", "Length (ft)")

row = 2
for system, elems in sorted(elements_by_base_system.items()):
    for el in elems:
        el_id = el.Id.IntegerValue
        length_param = el.LookupParameter("Length")
        length = length_param.AsDouble() if length_param else 0
        length_ft = round(length, 2)

        details_sheet.write("A{}".format(row), system)
        details_sheet.write("B{}".format(row), el_id)
        details_sheet.write("C{}".format(row), length_ft)
        row += 1

# Finalize Excel file
workbook.close()

# Notify user
forms.alert("Pipe Volume Summary exported to:\n\n{}".format(export_path), title="Export Complete", warn_icon=False)

#This last bit of code will display a report within revit instead of making an excel workbook. 

# Build output string
#lines = ["PIPE VOLUME PER SYSTEM TYPE", "-" * 40]
#for system, ft3 in sorted(volume_by_base_system.items()):
    #gal = ft3 * 7.48052
    #lines.append("{:<30} : {:,.2f} ft³ / {:,.2f} gal".format(system, ft3, gal))

#lines.append("\nPIPE ELEMENTS BY SYSTEM TYPE")
#lines.append("-" * 40)

#for system, elems in sorted(elements_by_base_system.items()):
    #lines.append("\n{} ({} pipes):".format(system, len(elems)))
    #for el in elems:
        #el_id = el.Id.IntegerValue
        #length_param = el.LookupParameter("Length")
        #length = length_param.AsDouble() if length_param else 0
        #length_ft = round(length, 2)
        #lines.append("  - ID {} | Length: {:.2f} ft".format(el_id, length_ft))

# Show results
#forms.alert("\n".join(lines), title="Pipe Volume Summary")