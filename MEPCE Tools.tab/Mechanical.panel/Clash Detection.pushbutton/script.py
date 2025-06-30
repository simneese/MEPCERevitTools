# -*- coding: utf-8 -*-
__title__ = "Clash Detection"
__highlight__ = "new"
__doc__ = """Version = 1.0
Date    = 2025.06.30
_________________________________________________________________
Description:
This button will detect if selected element categories
are clashing and will draw detail lines around clashing geometry.
_________________________________________________________________
How-to:
→ Click button
→ Select 'Yes' or 'No' to only check on active view
→ Select main categories to check
→ Select system types to filter, or select none
→ Select categories to check against
→ Select system types to filter, or select none
_________________________________________________________________
Last update:
- [2025.06.30] - 1.0 RELEASE
_________________________________________________________________
To-Do:
- Add more filter options
_________________________________________________________________
Author: Simeon Neese"""


# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
# Regular + Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import PipingSystemType
from Autodesk.Revit.DB.Mechanical import MechanicalSystemType, DuctSystemType
from types import NoneType

# pyRevit
from pyrevit import revit, forms

# .NET Imports (You often need List import)
import clr
from pyrevit.forms import alert

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
from Snippets._selection import get_elements_of_categories

def get_solid_geometry(element,transform=None):
    """Get solid geometry from an input element"""
    opt = Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = False
    geo_elem = element.get_Geometry(opt)

    solids = []

    for geo_obj in geo_elem:
        if isinstance(geo_obj, Solid) and geo_obj.Volume > 0:
            solids.append(geo_obj)
        elif isinstance(geo_obj, GeometryInstance):
            inst_geo = geo_obj.GetInstanceGeometry()
            for inst_obj in inst_geo:
                if isinstance(inst_obj, Solid) and inst_obj.Volume > 0:
                    solids.append(inst_obj)

    if transform:
        solids = [SolidUtils.CreateTransformed(s, transform) for s in solids if s]

    return solids

def bounding_boxes_intersect(bb1, bb2):
    """Check if two bounding boxes intersect"""
    if not bb1 or not bb2:
        return False
    return (bb1.Max.X >= bb2.Min.X and bb1.Min.X <= bb2.Max.X and
            bb1.Max.Y >= bb2.Min.Y and bb1.Min.Y <= bb2.Max.Y and
            bb1.Max.Z >= bb2.Min.Z and bb1.Min.Z <= bb2.Max.Z)

def flatten_solids(items):
    """Flatten a list of solids to only contain solid objects, and no nested lists."""
    flat = []
    if isinstance(items, Solid):
        flat.append(items)
    elif isinstance(items, list) or isinstance(items, tuple):
        for i in items:
            flat.extend(flatten_solids(i))  # recursion
    return flat

def select_types_form(selected_categories=[]):
    """Open form to select duct and pipe system types."""
    categoryobjs = []
    stypes = []
    for selection in selected_categories:
        if selection == 'Pipes':
            spipetypes = forms.SelectFromList.show(
                l_pipingsnames,
                multiselect=True,
                title='Filter Pipe System Types?',
                button_name='Select'
            )
            categoryobjs.append(BuiltInCategory.OST_PipeCurves)
            for typ in spipetypes:
                stypes.append(typ)
        if selection == 'Ducts':
            sducttypes = forms.SelectFromList.show(
                l_ductsnames,
                multiselect=True,
                title='Filter Duct System Types?',
                button_name='Select'
            )
            categoryobjs.append(BuiltInCategory.OST_DuctCurves)
            for typ in sducttypes:
                stypes.append(typ)
    return [stypes,categoryobjs]

def project_point_to_plane(point, plane_origin, plane_normal):
    # vector from origin to point
    vec = point - plane_origin
    distance = vec.DotProduct(plane_normal)
    return point - distance * plane_normal

def project_curve_to_plane(_curve, plane_origin, plane_normal):
    start = _curve.GetEndPoint(0)
    end = _curve.GetEndPoint(1)
    ps = project_point_to_plane(start, plane_origin, plane_normal)
    pe = project_point_to_plane(end, plane_origin, plane_normal)
    return Line.CreateBound(ps, pe)


def build_closed_curveloop(curves, tolerance=1e-6):
    loop = CurveLoop()
    if not curves:
        return loop

    # Start with the first curve
    current = curves.pop(0)
    loop.Append(current)
    current_end = current.GetEndPoint(1)

    while curves:
        found = False
        for i, c in enumerate(curves):
            start = c.GetEndPoint(0)
            end = c.GetEndPoint(1)
            if current_end.IsAlmostEqualTo(start, tolerance):
                loop.Append(c)
                current_end = end
                curves.pop(i)
                found = True
                break
            elif current_end.IsAlmostEqualTo(end, tolerance):
                loop.Append(c.CreateReversed())
                current_end = start
                curves.pop(i)
                found = True
                break
        if not found:
            raise Exception("Cannot form a continuous loop with provided curves.")

    return loop

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
# Get piping and duct system types
pipingsystemtypes = list(FilteredElementCollector(doc).OfClass(PipingSystemType).ToElements())
ductsystemtypes = list(FilteredElementCollector(doc).OfClass(MechanicalSystemType).ToElements())

# Create a dictionary for piping system types
d_pipingstypes = {}
l_pipingsnames = []
for pst in pipingsystemtypes:
    try:
        name = pst.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        d_pipingstypes.update({name: pst})
        l_pipingsnames.append(name)
        print name
    except Exception as e:
        print("Error accessing name:", e)
        print("Offending object:", pst)
        continue

# Create a dictionary for duct system types
d_ductstypes = {}
l_ductsnames = []
for dct in ductsystemtypes:
    try:
        name = dct.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        d_ductstypes.update({name: dct})
        l_ductsnames.append(name)
        print name
    except Exception as e:
        print("Error accessing name:", e)
        print("Offending object:", dct)
        continue

#  _____
# / __  \
# `' / /'
#   / /
# ./ /___
# \_____/
# Get user inputs
# Ask user if they would like to filter elements by active view
sviewfilt = forms.alert(msg="Filter elements by Active View?",options=["Yes","No"],warn_icon=False)

# Ask user to select a category to check
availcategories = ['Ducts','Pipes']
scategories = forms.SelectFromList.show(
    availcategories,
    multiselect=True,
    title='Select First Categories',
    button_name='Select'
)

# Alert user if no category is selected and exit script
if not scategories:
    alert(msg="No category selected.", title='Category Selection',
          sub_msg='At least one category should be selected. Exiting script.',exitscript=True)

# Ask user to select system types from each of the selected categories
form1out = select_types_form(scategories)
stypes = form1out[0]
categoryobjs = form1out[1]

#  _____
# |____ |
#     / /
#     \ \
# .___/ /
# \____/
# Get User Inputs II: The Movie
# Ask user to select a category to check aginst
availcategories2 = ['Structural','Ducts','Pipes']
scategories2 = forms.SelectFromList.show(
    availcategories2,
    multiselect=True,
    title='Select Categories to Check Against',
    button_name='Select'
)

# Alert user if no category is selected and exit script
if not scategories2:
    alert(msg="No category selected.", title='Category Selection',
          sub_msg='At least one category should be selected to check against. Exiting script.',exitscript=True)

# Ask user to select system types from each of the selected categories
structural_categories = [
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFoundation,
    BuiltInCategory.OST_StructuralTruss,
    BuiltInCategory.OST_StructuralStiffener,
    BuiltInCategory.OST_StructuralTendons,
    BuiltInCategory.OST_Columns
]
form2out = select_types_form(scategories2)
stypes2 = form2out[0]
categoryobjs2 = form2out[1]

#    ___
#   /   |
#  / /| |
# / /_| |
# \___  |
#     |_/
# Get filtered elements to check and elements to check against
# Get Active View (or don't)
active_view = doc.ActiveView
if sviewfilt == 'Yes':
    selected_view = active_view
else:
    selected_view = []

# Get filtered elements to check
elements = get_elements_of_categories(
    categories=categoryobjs,
    view=selected_view,
    systemtypes=stypes,
    readout=True
)

# Get filtered elements to check against
elements2=[]
if categoryobjs2:
    selels = get_elements_of_categories(
        categories=categoryobjs2,
        #view=selected_view,
        systemtypes=stypes2,
        readout=True
    )
    for el in selels:
        elements2.append(el)
# Get structural elements to check against
structuralelements = []
if 'Structural' in scategories2:
    selels = get_elements_of_categories(
        categories=structural_categories,
        #view=selected_view,
        linked=True,
        readout=True
    )
    for el in selels:
        elements2.append(el)

#  _____
# |  ___|
# |___ \
#     \ \
# /\__/ /
# \____/
# Get geometry for elements
solids1 = {}
bounding1 = {}
for el in elements:
    solids = get_solid_geometry(el)
    flat_solids = [s for s in solids if isinstance(s, Solid) and s.Volume > 0]
    bounding = el.get_BoundingBox(None)
    if flat_solids:
        solids1[el.Id] = flat_solids
        bounding1[el.Id] = bounding
element_ids1 = list(solids1.keys())

solids2 = {}
bounding2 = {}
for el in elements2:
    solids = get_solid_geometry(el)
    flat_solids = [s for s in solids if isinstance(s, Solid) and s.Volume > 0]
    bounding = el.get_BoundingBox(None)
    if flat_solids:
        solids2[el.Id] = flat_solids
        bounding2[el.Id] = bounding
element_ids2 = list(solids2.keys())

#   ____
#  / ___|
# / /___
# | ___ \
# | \_/ |
# \_____/
# Check for clashes and get clash geometry
intersection_results = {}
clash_geometry = []
for id1 in element_ids1:
    id1_intersections = []
    for id2 in element_ids2:
        bb1 = bounding1[id1]
        bb2 = bounding2[id2]
        bbintersect = bounding_boxes_intersect(bb1,bb2)                      # First, check if bounding boxes intersect
        if bbintersect is True:
            solid1 = solids1[id1]
            solid1 = flatten_solids(solid1)
            solid2 = solids2[id2]
            solid2 = flatten_solids(solid2)
            for s1 in solid1:
                for s2 in solid2:
                    if not isinstance(s1, Solid) or not isinstance(s2, Solid):
                        print 'Skipping non-solid: {}, {}'.format(type(solid1),type(solid2))
                        continue
                    try:                                                     # If they clash, check if solids intersect
                        result = BooleanOperationsUtils.ExecuteBooleanOperation(s1,s2,BooleanOperationsType.Intersect)
                        if result.Volume > 0:
                            id1_intersections.append(id2)
                            clash_geometry.append(result)
                            break
                    except:
                        continue
    if id1_intersections:
        intersection_results[id1] = id1_intersections
results_keys = list(intersection_results.keys())

print 'Elements Selected to Check: {}'.format(len(element_ids1))
print 'Elements to Check Against: {}'.format(len(element_ids2))
print 'Number of Checks: {}'.format(len(element_ids1)*len(element_ids2))
print 'Intersecting: {}'.format(len(results_keys))

#  ______
# |___  /
#    / /
#   / /
# ./ /
# \_/
# Get planar projection of clash geometry
plane = active_view.SketchPlane.GetPlane()
origin = plane.Origin
normal = plane.Normal
if sviewfilt == 'Yes':
    with Transaction(doc, "Create Detail Curve") as t:
        t.Start()
        for solid in clash_geometry:
            edges = []
            for face in solid.Faces:
                for loop in face.GetEdgesAsCurveLoops():
                    for curve in loop:
                        edges.append(curve)

            projected_curves = []
            for e in edges:
                start = e.GetEndPoint(0)
                end = e.GetEndPoint(1)

                projected_start = project_point_to_plane(start, origin, normal)
                projected_end = project_point_to_plane(end, origin, normal)

                try:
                    projected_line = Line.CreateBound(projected_start,projected_end)
                    projected_curves.append(projected_line)
                except:
                    continue
            for curve in projected_curves:
                try:
                    detail_line = doc.Create.NewDetailCurve(active_view,curve)
                    _id = detail_line.Id
                    ogs = OverrideGraphicSettings()
                    color = Color(red=255,green=0,blue=0)
                    ogs.SetProjectionLineColor(color)
                    doc.ActiveView.SetElementOverrides(_id,ogs)
                except:
                    continue
        t.Commit()