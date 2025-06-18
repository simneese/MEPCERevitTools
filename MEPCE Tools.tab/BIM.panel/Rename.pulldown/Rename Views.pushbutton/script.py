# -*- coding: utf-8 -*-
__title__ = "Rename Views"
__highlight__ = "updated"
__doc__ = """Version = 1.1
Date    = 2025.06.18
_________________________________________________________________
Description:
Rename View in Revit using find/replace logic.
_________________________________________________________________
How-to:
(Optional) -> Select views in the project browser
-> Click on the button
-> Select Views
-> Define Renaming Rules
-> Rename Views
_________________________________________________________________
Last update:
- [2025.06.18] - Will now automatically add a space after the prefix
- [2025.05.28] - 1.0 RELEASE
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

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ MAIN
#==================================================
# START CODE HERE

#  __    _____        _              _     _   _  _
# /  |  /  ___|      | |            | |   | | | |(_)
# `| |  \ `--.   ___ | |  ___   ___ | |_  | | | | _   ___ __      __ ___
#  | |   `--. \ / _ \| | / _ \ / __|| __| | | | || | / _ \\ \ /\ / // __|
# _| |_ /\__/ /|  __/| ||  __/| (__ | |_  \ \_/ /| ||  __/ \ V  V / \__ \
# \___/ \____/  \___||_| \___| \___| \__|  \___/ |_| \___|  \_/\_/  |___/

# Get views selected in the project browser
sel_el_ids = uidoc.Selection.GetElementIds()
sel_elem = [doc.GetElement(e_id) for e_id in sel_el_ids]
sel_views = [el for el in sel_elem if issubclass(type(el), View)]

# If none selected- prompt selectviews from pyrevit.forms.select_views()
if not sel_views:
    sel_views = forms.select_views()

# Ensure Views Selected
if not sel_views:
    forms.alert('No Views Selected. Please Try Again', exitscript=True)

#  _____  ______         __  _               ______                                 _                ______         _
# / __  \ |  _  \       / _|(_)              | ___ \                               (_)               | ___ \       | |
# `' / /' | | | |  ___ | |_  _  _ __    ___  | |_/ /  ___  _ __    __ _  _ __ ___   _  _ __    __ _  | |_/ / _   _ | |  ___  ___
#   / /   | | | | / _ \|  _|| || '_ \  / _ \ |    /  / _ \| '_ \  / _` || '_ ` _ \ | || '_ \  / _` | |    / | | | || | / _ \/ __|
# ./ /___ | |/ / |  __/| |  | || | | ||  __/ | |\ \ |  __/| | | || (_| || | | | | || || | | || (_| | | |\ \ | |_| || ||  __/\__ \
# \_____/ |___/   \___||_|  |_||_| |_| \___| \_| \_| \___||_| |_| \__,_||_| |_| |_||_||_| |_| \__, | \_| \_| \__,_||_| \___||___/
#                                                                                              __/ |
#                                                                                             |___/

from rpw.ui.forms import (FlexForm, Label, TextBox, Separator, Button)
components =    [Label('Prefix:'),  TextBox('prefix'),
                 Label('Find:'),    TextBox('find'),
                 Label('Replace:'), TextBox('replace'),
                 Label('Suffix:'),  TextBox('suffix'),
                 Separator(),       Button('Rename Views')]

form = FlexForm('Title', components)
form.show()

user_inputs = form.values
prefix      = user_inputs['prefix'] + " "
find        = user_inputs['find']
replace     = user_inputs['replace']
suffix      = user_inputs['suffix']

# Start Transaction to make changes in project
t = Transaction(doc, 'py-Rename Views')

t.Start()
for view in sel_views:

    #  _____   _____                     _                                   _   _  _                   _   _
    # |____ | /  __ \                   | |                                 | | | |(_)                 | \ | |
    #     / / | /  \/ _ __   ___   __ _ | |_   ___   _ __    ___ __      __ | | | | _   ___ __      __ |  \| |  __ _  _ __ ___    ___
    #     \ \ | |    | '__| / _ \ / _` || __| / _ \ | '_ \  / _ \\ \ /\ / / | | | || | / _ \\ \ /\ / / | . ` | / _` || '_ ` _ \  / _ \
    # .___/ / | \__/\| |   |  __/| (_| || |_ |  __/ | | | ||  __/ \ V  V /  \ \_/ /| ||  __/ \ V  V /  | |\  || (_| || | | | | ||  __/
    # \____/   \____/|_|    \___| \__,_| \__| \___| |_| |_| \___|  \_/\_/    \___/ |_| \___|  \_/\_/   \_| \_/ \__,_||_| |_| |_| \___|
    old_name = view.Name
    new_name = prefix + old_name.replace(find, replace) + suffix

    #    ___  ______                                        _   _  _
    #   /   | | ___ \                                      | | | |(_)
    #  / /| | | |_/ /  ___  _ __    __ _  _ __ ___    ___  | | | | _   ___ __      __ ___
    # / /_| | |    /  / _ \| '_ \  / _` || '_ ` _ \  / _ \ | | | || | / _ \\ \ /\ / // __|
    # \___  | | |\ \ |  __/| | | || (_| || | | | | ||  __/ \ \_/ /| ||  __/ \ V  V / \__ \
    #     |_/ \_| \_| \___||_| |_| \__,_||_| |_| |_| \___|  \___/ |_| \___|  \_/\_/  |___/ (ensure unique names)
    for i in range(20):
        try:
            view.Name = new_name
            print('{} -> {}'.format(old_name, new_name))
            break
        except:
            new_name += '*'

t.Commit()

print('-'*50)
print('Done!')