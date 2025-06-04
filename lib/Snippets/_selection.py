# -*- coding: utf-8 -*-

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
from Autodesk.Revit.DB import *
from types import NoneType

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#==================================================
doc   = __revit__.ActiveUIDocument.Document #type: Document
uidoc = __revit__.ActiveUIDocument          #type: UIDocument
app   = __revit__.Application               #type: Application

#  _____        _     _____        _              _              _   _____  _                                _
# |  __ \      | |   /  ___|      | |            | |            | | |  ___|| |                              | |
# | |  \/  ___ | |_  \ `--.   ___ | |  ___   ___ | |_   ___   __| | | |__  | |  ___  _ __ ___    ___  _ __  | |_  ___
# | | __  / _ \| __|  `--. \ / _ \| | / _ \ / __|| __| / _ \ / _` | |  __| | | / _ \| '_ ` _ \  / _ \| '_ \ | __|/ __|
# | |_\ \|  __/| |_  /\__/ /|  __/| ||  __/| (__ | |_ |  __/| (_| | | |___ | ||  __/| | | | | ||  __/| | | || |_ \__ \
#  \____/ \___| \__| \____/  \___||_| \___| \___| \__| \___| \__,_| \____/ |_| \___||_| |_| |_| \___||_| |_| \__||___/
# Get elements selected by user

def get_selected_elements(filter_types=None):
    """Get Selected Elements in Revit UI.
    You can provide a list of types for filter_types parameter (optional)

    e.g.
    sel_wall = get_selected_elements([Wall])"""
    selected_element_ids = uidoc.Selection.GetElementIds()
    selected_elements    = [doc.GetElement(e_id) for e_id in selected_element_ids]

    # Filter Selection (Optionally)
    if filter_types:
        return [el for el in selected_elements if type(el) in filter_types]
    return selected_elements