# -*- coding: utf-8 -*-

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
from Autodesk.Revit.DB import *

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#==================================================
doc   = __revit__.ActiveUIDocument.Document #type: Document
uidoc = __revit__.ActiveUIDocument          #type: UIDocument
app   = __revit__.Application               #type: Application

# ______                           _      _         _____         _                      _
# | ___ \                         | |    | |       /  ___|       (_)                    | |
# | |_/ /  ___  _   _  ___   __ _ | |__  | |  ___  \ `--.  _ __   _  _ __   _ __    ___ | |_  ___
# |    /  / _ \| | | |/ __| / _` || '_ \ | | / _ \  `--. \| '_ \ | || '_ \ | '_ \  / _ \| __|/ __|
# | |\ \ |  __/| |_| |\__ \| (_| || |_) || ||  __/ /\__/ /| | | || || |_) || |_) ||  __/| |_ \__ \
# \_| \_| \___| \__,_||___/ \__,_||_.__/ |_| \___| \____/ |_| |_||_|| .__/ | .__/  \___| \__||___/
#                                                                   | |    | |
#                                                                   |_|    |_|

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