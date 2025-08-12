# -*- coding: utf-8 -*-
__title__   = "Calculate Exhaust"
__highlight__ = "new"
__doc__     = """Version = 1.0
Date    = 2025.08.12
_________________________________________________________________
Description:
Calculate exhaust rates for restrooms, janitor closets, and other spaces which require exhaust.
_________________________________________________________________
How-to:
-> Click button
-> Select between toilet / mop sink exhaust, air change exhaust, or both
-> Input exhaust / toilet and mop sink
-> Select volume or floor area option for air changes
-> Input air change exhaust rate
-> Done!
_________________________________________________________________
Last update:
- [2025.08.12] - 1.0 RELEASE
_________________________________________________________________
Author: Simeon Neese"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
import re
import clr
import math
from pyrevit import revit,forms
from pyrevit.forms import alert

from Autodesk.Revit.DB import *

from types import NoneType

from rpw.ui.forms import (FlexForm, Label, TextBox, Separator, Button)

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
from Snippets._selection import get_elements_of_categories,select_multiple
from Snippets._filledregions import get_rooms,get_faces,get_filled_region_corners
from Snippets._geometry import get_point_average,get_solid_geometry,flatten_solids


def draw_flattened_bounding_boxes(bboxes):
    """
    Draws flattened 2D red boxes in the active view from a list of BoundingBoxXYZ.

    Args:
        bboxes (list[BoundingBoxXYZ]): List of bounding boxes to flatten and draw.
    """
    if not bboxes:
        return

    red_color = Color(255, 0, 0)
    ogs = OverrideGraphicSettings().SetProjectionLineColor(red_color)

    t = Transaction(doc, "Draw Flattened Bounding Boxes")
    t.Start()

    for bbox in bboxes:
        if not bbox:
            continue

        min_pt = bbox.Min
        max_pt = bbox.Max

        # Flatten to a single Z plane (active view's origin Z or min_pt.Z)
        z = getattr(active_view, "Origin", XYZ(0, 0, min_pt.Z)).Z

        # Define 4 corners in 2D
        p1 = XYZ(min_pt.X, min_pt.Y, z)
        p2 = XYZ(max_pt.X, min_pt.Y, z)
        p3 = XYZ(max_pt.X, max_pt.Y, z)
        p4 = XYZ(min_pt.X, max_pt.Y, z)

        # Create lines for each side
        lines = [
            Line.CreateBound(p1, p2),
            Line.CreateBound(p2, p3),
            Line.CreateBound(p3, p4),
            Line.CreateBound(p4, p1)
        ]

        # Draw the lines and override their color
        for line in lines:
            detail_line = doc.Create.NewDetailCurve(active_view, line)
            active_view.SetElementOverrides(detail_line.Id, ogs)

    t.Commit()

def get_link_instance_from_link_doc(doc, link_doc):
    """Return the RevitLinkInstance for a given LinkDocument."""
    for link_instance in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        if link_instance.GetLinkDocument() and link_instance.GetLinkDocument().Equals(link_doc):
            return link_instance
    return None

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

debugmode = False

# 1
# What to update?
update = alert("Which would you like to calculate for?",options=['Restroom / Custodian Fixture Exhaust','Air Change Exhaust','Both'],warn_icon=False)
if not update:
    alert("Nothing chosen. Exiting script.",exitscript=True)

#  __
# /  |
# `| |
#  | |
# _| |_
# \___/
# Get Toilets and Mop Sinks
if update == 'Restroom / Custodian Fixture Exhaust' or update == 'Both':
    plumbingfix = get_elements_of_categories(
        categories=[BuiltInCategory.OST_PlumbingFixtures],
        linked=True,
        readout=debugmode)

    exfix = []
    for fix in plumbingfix:
        family = fix.LookupParameter("Family").AsValueString()
        if "Toilet" in family and not "Stall" in family and not "Tissue" in family:
            exfix.append(fix)
        elif "Urinal" in family:
            exfix.append(fix)
        elif "Mop Sink" in family:
            exfix.append(fix)

    components1 =   [Label('CFM / Fixture'),   TextBox('inpexhaust'),
                     Separator(),              Button('Continue')]
    form1 = FlexForm('Input Exhaust per Fixture',components1)
    form1.show()

    form1values = form1.values
    inpexhaust = float(form1values['inpexhaust'])


    #  _____
    # / __  \
    # `' / /'
    #   / /
    # ./ /___
    # \_____/
    # Group filled regions and fixtures by level
    active_view = doc.ActiveView

    roomregions = get_rooms()
    roomregions = [room for room in roomregions if room.LookupParameter("MEPCE Update Region").AsInteger() == 1]
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()

    levelnames = [level.Name for level in levels]
    levelregions = {}
    levelfixtures = {}
    levelbounding = {}

    for level in levelnames:
        if not level in levelregions:
            levelregions[level] = []
        if not level in levelfixtures:
            levelfixtures[level] = []
        if not level in levelbounding:
            levelbounding[level] = []
        for room in roomregions:
            if doc.GetElement(room.OwnerViewId).GenLevel.Name == level:
                levelregions[level].append(room)
        for fixture in exfix:
            if fixture.LookupParameter("Level").AsValueString() == level:
                levelfixtures[level].append(fixture)
                levelbounding[level].append(fixture.get_BoundingBox(None))


    #  _____
    # |____ |
    #     / /
    #     \ \
    # .___/ /
    # \____/
    # Get plumbing fixtures for each region and calculate fixture exhaust
    if debugmode is True:
        print "RESTROOM AND CUSTODIAN EXHAUST:\n\n"
    regionexhaust = {}
    for level in levelnames:
        regions = levelregions[level]
        if not regions:
            continue
        faces = get_faces([regions])[0]
        fixtures = levelfixtures[level]
        locations = [fixture.Location.Point for fixture in fixtures]
        if not locations:
            continue
        for r_idx,region in enumerate(regions):
            fixinregion = []
            for f_idx,fixture in enumerate(fixtures):
                bounding = fixture.get_BoundingBox(None)
                MIN = bounding.Min
                MAX = bounding.Max
                AVG = get_point_average(MIN,MAX)
                if faces[r_idx].Project(AVG):
                    fixinregion.append(fixture)
                    bounding = fixture.get_BoundingBox(None)
                    levelbounding[level].append(bounding)
            fixexhaust = len(fixinregion)*inpexhaust/60
            if len(fixinregion) > 0:
                regionexhaust[region] = fixexhaust
                print "\nRoom: {}".format(region.LookupParameter("MEPCE Room Number").AsString()+" "+region.LookupParameter("MEPCE Room Name").AsString())
                print "Fixtures: {}\nExhaust: {} CFM".format(len(fixinregion),int(fixexhaust*60))

    if debugmode is True:
        print "~"*100
        draw_flattened_bounding_boxes(levelbounding[active_view.GenLevel.Name])

#    ___
#   /   |
#  / /| |
# / /_| |
# \___  |
#     |_/
# Select exhaust calc rules and get rooms
ceilingheights = {}
if update == 'Air Change Exhaust' or update == 'Both':
    method = alert("Select exhaust calculation method",options=["ACH (Volume)","CFM/sf (Area)"],warn_icon=False)

    exhaustrate = 0
    if method == "ACH (Volume)":
        components =    [Label('Air Changes / Hour:'),     TextBox('exhaustrate'),
                         Separator(),       Button('Continue')]
        form = FlexForm('Input Exhaust Rate',components)
        form.show()
        user_inputs = form.values
        exhaustrate = user_inputs['exhaustrate']
    elif method == "CFM/sf (Area)":
        components =    [Label('CFM / Square Feet:'),     TextBox('exhaustrate'),
                         Separator(),       Button('Continue')]
        form = FlexForm('Input Exhaust Rate',components)
        form.show()
        user_inputs = form.values
        exhaustrate = user_inputs['exhaustrate']
    else:
        alert("No exhaust calculation method selected. Exiting script.",exitscript=True)

    regionsd = {}
    for region in roomregions:
        name = region.LookupParameter("MEPCE Room Number").AsString()+" "+region.LookupParameter("MEPCE Room Name").AsString()
        regionsd[name] = region
    r_names = regionsd.keys()

    selectiontype = alert("How would you like to select rooms?",options=["Manual Select","Keyword Search"],warn_icon=False)

    selectedregions = []
    if selectiontype == "Manual Select":
        alert("Manual Select",sub_msg="Select Filled Regions to calculate exhaust",warn_icon=False)
        selectedregions = select_multiple([BuiltInCategory.OST_DetailComponents],"Select Filled Regions")
    elif selectiontype == "Keyword Search":
        selectedregions = forms.SelectFromList.show(
            r_names,
            multiselect=True,
            title="Select Rooms",
            button_name='Select'
        )
        selectedregions = [regionsd[name] for name in selectedregions]
    else:
        alert("Did not pick a selection type. Exiting script.",exitscript=True)

    if debugmode is True:
        print "Exhaust Rate: {} {}".format(exhaustrate,method)
        print "{} Regions Selected".format(len(selectedregions))
        print "~"*100

    #  _____
    # |  ___|
    # |___ \
    #     \ \
    # /\__/ /
    # \____/
    # Get ceilings (organized by level)
    ceilings = get_elements_of_categories(
        categories=[BuiltInCategory.OST_Ceilings],
        linked=True,
        readout=debugmode)

    linkdoc = [ceiling.Document for ceiling in ceilings][0]
    linkinstance = get_link_instance_from_link_doc(doc,linkdoc)
    transform = linkinstance.GetTransform()

    levelceilings = {}
    for ceiling in ceilings:
        levelname = ceiling.LookupParameter("Level").AsValueString()
        ceilingsolid = get_solid_geometry(ceiling,transform)
        if not ceilingsolid:
            continue
        flatceilingsolid = flatten_solids(ceilingsolid)[0]
        faces = flatceilingsolid.Faces
        bottomface = None
        for face in faces:
            normal = face.ComputeNormal(UV(0.5,0.5))
            if normal.Z == -1:
                bottomface = face
        if levelname not in levelceilings:
            levelceilings[levelname] = [bottomface]
        else:
            levelceilings[levelname].append(bottomface)


    #   ____
    #  / ___|
    # / /___
    # | ___ \
    # | \_/ |
    # \_____/
    # Calculate exhaust CFM for selected rooms
    for region in selectedregions:
        roomname = region.LookupParameter("MEPCE Room Number").AsString()+" "+region.LookupParameter("MEPCE Room Name").AsString()
        area = region.LookupParameter("Area").AsDouble()
        level = doc.GetElement(region.OwnerViewId).GenLevel.Name
        if method == "CFM/sf (Area)":
            regionexhaust[region] = area*float(exhaustrate)/60
            print "\nRoom: {}\nArea: {} SF\nRate: {} CFM/SF\nExhaust: {} CFM".format(roomname,exhaustrate,area,regionexhaust[region]*60)
            continue
        elif method == "ACH (Volume)":
            ceilings = levelceilings[level]
            corners = get_filled_region_corners(region)
            ceilingheight = 10
            heights = []
            for ceiling in ceilings:
                if not ceiling:
                    continue
                for corner in corners:
                    if ceiling.Project(corner):
                        heights.append(ceiling.Project(corner).Distance)
                        break
            if heights:
                ceilingheight = sum(heights) / len(heights)
                ceilingheights[region] = ceilingheight
            volume = area*ceilingheight
            regionexhaust[region] = volume*float(exhaustrate)/3600
            print "\nRoom: {}\nArea: {} SF\nHeight: {} FT\nVolume: {} CF\nRate: {} ACH\nExhaust: {} CFM".format(roomname,area,ceilingheight,volume,exhaustrate,regionexhaust[region]*60)


#  ______
# |___  /
#    / /
#   / /
# ./ /
# \_/
# Update exhaust
t = Transaction(doc,"Update Exhaust Parameter")
t.Start()
for region in regionexhaust:
    region.LookupParameter("MEPCE Room Exhaust").Set(regionexhaust[region])
    if region in ceilingheights:
        region.LookupParameter("MEPCE Ceiling Height").Set(ceilingheights[region])
t.Commit()