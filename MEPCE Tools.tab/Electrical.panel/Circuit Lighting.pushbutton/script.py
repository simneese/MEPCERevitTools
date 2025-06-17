# -*- coding: utf-8 -*-
__title__ = "Circuit Lighting"
__doc__ = """Version = 1.1
Date    = 2025.06.09
_________________________________________________________________
Description:
This button will circuit lights based on filled regions.
_________________________________________________________________
How-to:
-> Create/update filled regions around grouped lights
-> Click button
_________________________________________________________________
Last update:
- [2025.06.09] - 1.1 Fixed bug which caused elements to be
                     fetched incorrectly due to variable names.
                     Error messages now print in addition to
                     sending an alert.
- [2025.06.04] - 1.0 RELEASE
_________________________________________________________________
To-Do:
- Add method to add circuits to panels
_________________________________________________________________
Author: Simeon Neese"""


# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
# Regular + Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import ElectricalSystem, ElectricalSystemType
from types import NoneType

# pyRevit
from pyrevit import revit, forms

# .NET Imports (You often need List import)
import clr
from pyrevit.forms import alert

clr.AddReference("System")
from System.Collections.Generic import List

from enum import Enum

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
from Snippets._filledregions import get_filled_regions
from Snippets._filledregions import get_faces

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ MAIN
#==================================================

#  __
# /  |
# `| |
#  | |
# _| |_
# \___/
# Get all views
allviews = list(FilteredElementCollector(doc).OfClass(View))

#  _____
# / __  \
# `' / /'
#   / /
# ./ /___
# \_____/
# Get faces from filled regions
dic = get_filled_regions(allviews,'lighting')
regions = dic["filledregions"]  # List of regions organized by view like [[],[],[],...,[]]
filtviews = dic["filtviews"]    # List of lighting views
faces = get_faces(regions)      # List of list of faces. Organized by view like [[],[],[],...,[]]

#  _____
# |____ |
#     / /
#     \ \
# .___/ /
# \____/
# Get Light Locations grouped by view
alllights = []  # list of lights structured like allpoints list
allpoints = []  # Highest level list of light locations - should be structured like [[],[],[],...,[]] where each sub-list is a list of light locations for a view
allids = []     # All light fixture element Ids grouped by view

for i in range(len(filtviews)):
    viewpoints = [] # Sublist which will be appended to allpoints - is a list of light locations for the specific view
    viewlights = list(FilteredElementCollector(doc,filtviews[i].Id).OfCategory(BuiltInCategory.OST_LightingFixtures)) # Get lights in current view

    for light in viewlights:
        location = light.Location
        if type(location) is NoneType:
            continue
        else:
            viewpoints.append(location.Point)
    allpoints.append(viewpoints)
    alllights.append(viewlights)

#    ___
#   /   |
#  / /| |
# / /_| |
# \___  |
#     |_/
# Project light locations onto faces (only do for ones in the same view)
circuitgroups = []
for i in range(len(faces)):
    viewpoints = allpoints[i]   # All points in associated view
    viewlights = alllights[i]   # All lights in associated view

    for face in faces[i]:
        lightgroup = []         # List of lights inside this face
        for k in range(len(viewpoints)):
            if face.Project(viewpoints[k]):
                lightgroup.append(viewlights[k])
            else:
                continue
        circuitgroups.append(lightgroup)    # Will group lights by associated face

#  _____
# |  ___|
# |___ \
#     \ \
# /\__/ /
# \____/
# Check for existing circuits
allregions = []     # Get a flattened list of regions
comparam = []       # Get list of comment parameters for filled regions
markparam = []      # Get list of mark parameters for filled regions
for view in regions:
    for region in view:
        allregions.append(region)
        comparam.append(region.LookupParameter('Comments'))
        markparam.append(region.LookupParameter('Mark'))

allcircuits = list(FilteredElementCollector(doc).OfClass(ElectricalSystem))     # Get all circuits
assoccircuits = []

if allcircuits:
    for param in markparam:
        assoc = None
        for circuit in allcircuits:
            if param.AsString() == str(circuit.Id):
                assoc = circuit
        assoccircuits.append(assoc)

#   ____
#  / ___|
# / /___
# | ___ \
# | \_/ |
# \_____/
# Assign Lights to Circuits
with Transaction(doc,'Create Lighting Circuits') as t:
    t.Start()

    for group_index in range(len(circuitgroups)):
        group = circuitgroups[group_index]                                      # Get group
        #print 'Starting Group {}'.format(group_index)
        if not group:                                                           # Check if the filled region is empty
            #print "[Group {}] Empty group, skipping.".format(group_index)
            comparam[group_index].Set('Group {}'.format(group_index))
            markparam[group_index].Set('')
            continue

        valid_elements = []                                                     # List of elements which can be added to the circuit
        for elem in group:
            if not isinstance(elem, FamilyInstance):
                continue

            mep = elem.MEPModel
            if mep is None:
                continue

            conn_mgr = mep.ConnectorManager
            if conn_mgr is None or conn_mgr.Connectors.Size == 0:
                continue

            """# Optional: skip if already in circuit
            if hasattr(mep, "ElectricalSystems") and mep.ElectricalSystems and list(mep.ElectricalSystems):
                continue"""
            a = list(mep.GetElectricalSystems())

            if a:
                if str(a[0].Id) == markparam[group_index].AsString():
                    continue
                else:
                    setelement1 = ElementSet()
                    setelement1.Insert(elem)
                    try:
                        a[0].RemoveFromCircuit(setelement1)
                    except:
                        alert("Could not remove element ID {} from circuit ID {}".format(elem.Id.IntegerValue, a[0].Id))
                        print("Could not remove element ID {} from circuit ID {}".format(elem.Id.IntegerValue, a[0].Id))
                        continue

            valid_elements.append(elem)

        circ = assoccircuits[group_index]
        if circ:
            for element in valid_elements:
                setelement = ElementSet()
                setelement.Insert(element)
                try:
                    circ.AddToCircuit(setelement)
                except:
                    alert("[Group {}] Could not add element ID {}".format(group_index, element.Id.IntegerValue))
                    print("[Group {}] Could not add element ID {}".format(group_index, element.Id.IntegerValue))
                    continue

        else:
            if len(valid_elements) < 1:
                alert("[Group {}] No valid elements found.".format(group_index))
                print("[Group {}] No valid elements found.".format(group_index))
                continue

            main_elem = valid_elements[0]
            net_ids = List[ElementId]()
            net_ids.Add(main_elem.Id)

            try:
                circuit = ElectricalSystem.Create(doc, net_ids, ElectricalSystemType.PowerCircuit)
                if type(circuit) is NoneType:
                    net_ids.Add(valid_elements[1].Id)
                    circuit = ElectricalSystem.Create(doc, net_ids, ElectricalSystemType.PowerCircuit)
                    index_start = 2
                    if type(circuit) is NoneType:
                        alert('[Group {}] Could not create circuit. Element ID: {}, SystemType: {}'.format(group_index, main_elem.Id, ElectricalSystemType.PowerCircuit))
                        print('[Group {}] Could not create circuit. Element ID: {}, SystemType: {}'.format(group_index, main_elem.Id, ElectricalSystemType.PowerCircuit))
                        continue
                else:
                    index_start = 1

                comparam[group_index].Set("Group {}".format(group_index))
                markparam[group_index].Set(str(circuit.Id))

                for elem in valid_elements[index_start:]:
                    element_set = ElementSet()
                    element_set.Insert(elem)
                    try:
                        circuit.AddToCircuit(element_set)
                    except:
                        alert("[Group {}] Could not add element ID {}".format(group_index, elem.Id.IntegerValue))
                        print("[Group {}] Could not add element ID {}".format(group_index, elem.Id.IntegerValue))
                        continue
            except Exception as e:
                alert("[Group {}] Exception: {}".format(group_index, e))
                print("[Group {}] Exception: {}".format(group_index, e))

    t.Commit()

alert("Done!", title="Circuit Lighting", ok=True)