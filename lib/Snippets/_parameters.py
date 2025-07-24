# -*- coding: utf-8 -*-
from symbol import continue_stmt

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Mechanical import MechanicalSystemType, DuctSystemType
from Autodesk.Revit.DB.Plumbing import PipingSystemType
from Autodesk.Revit.ApplicationServices import Application
from types import NoneType

from pyrevit import revit, forms
from pyrevit.forms import alert

import os


# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#==================================================
doc   = __revit__.ActiveUIDocument.Document #type: Document
uidoc = __revit__.ActiveUIDocument          #type: UIDocument
app   = __revit__.Application               #type: Application

#  _____                                 _      ___   _  _  ______                                       _
# |_   _|                               | |    / _ \ | || | | ___ \                                     | |
#   | |   _ __ ___   _ __    ___   _ __ | |_  / /_\ \| || | | |_/ /  __ _  _ __   __ _  _ __ ___    ___ | |_   ___  _ __  ___
#   | |  | '_ ` _ \ | '_ \  / _ \ | '__|| __| |  _  || || | |  __/  / _` || '__| / _` || '_ ` _ \  / _ \| __| / _ \| '__|/ __|
#  _| |_ | | | | | || |_) || (_) || |   | |_  | | | || || | | |    | (_| || |   | (_| || | | | | ||  __/| |_ |  __/| |   \__ \
#  \___/ |_| |_| |_|| .__/  \___/ |_|    \__| \_| |_/|_||_| \_|     \__,_||_|    \__,_||_| |_| |_| \___| \__| \___||_|   |___/
#                   | |
#                   |_|
# Imports all shared parameters from a shared parameter file and assigns them to the same categories / group

def import_all_shared_parameters(categories, pgroup, param_file_path, instance=True, allow_vary_between_groups=False):
    """
    Imports all parameters from the shared parameter file and binds them to their categories.

    Parameters:
    - categories: list of built in categories to which the shared parameters will be applied
    - group: the group to which the parameters will be added
    - param_file_path: full path to the shared parameter file
    - instance: True for instance parameters, False for type parameters
    - allow_vary_between_groups: allow varying between groups (only applies to instance params)
    """

    if not os.path.exists(param_file_path):
        raise Exception("Shared parameter file not found at: {}".format(param_file_path))

    # Backup current shared parameter file
    original_file = app.SharedParametersFilename
    app.SharedParametersFilename = param_file_path

    shared_file = app.OpenSharedParameterFile()
    if shared_file is None:
        raise Exception("Failed to open shared parameter file.")

    with Transaction(doc, "Import Shared Parameters") as t:
        t.Start()

        binding_map = doc.ParameterBindings
        created_params = []

        for group in shared_file.Groups:
            for definition in group.Definitions:
                param_name = definition.Name

                # Skip if already bound
                if binding_map.Contains(definition):
                    continue

                # Determine categories to bind to (optional: customize this logic)
                # For now, bind to common MEP and Arch categories
                category_set = CategorySet()
                default_categories = categories

                for bic in default_categories:
                    cat = doc.Settings.Categories.get_Item(bic)
                    if cat:
                        category_set.Insert(cat)

                # Bind parameter
                binding = InstanceBinding(category_set) if instance else TypeBinding(category_set)
                binding_group = pgroup  # or customize this per parameter/group

                success = binding_map.Insert(definition, binding, binding_group)
                if success:
                    # Set vary between groups if applicable
                    if instance and allow_vary_between_groups:
                        param_elem = next((e for e in FilteredElementCollector(doc).OfClass(ParameterElement)
                                           if e.Name == param_name), None)
                        if param_elem:
                            param_elem.SetAllowVaryBetweenGroups(doc, True)

                    created_params.append(param_name)

        t.Commit()

    # Restore the original shared parameter file path
    app.SharedParametersFilename = original_file

    return created_params
