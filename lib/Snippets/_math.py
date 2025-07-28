# -*- coding: utf-8 -*-

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
from Autodesk.Revit.DB import *

from pyrevit import DB, revit

from types import NoneType

from pyrevit.forms import alert

import math

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#==================================================
doc   = __revit__.ActiveUIDocument.Document #type: Document
uidoc = __revit__.ActiveUIDocument          #type: UIDocument
app   = __revit__.Application               #type: Application

#  _____                     _           _____         _                _____                   _
# /  __ \                   | |         /  __ \       | |              |  __ \                 | |
# | /  \/ _ __   ___   __ _ | |_   ___  | /  \/  ___  | |  ___   _ __  | |  \/ _ __   __ _   __| |
# | |    | '__| / _ \ / _` || __| / _ \ | |     / _ \ | | / _ \ | '__| | | __ | '__| / _` | / _` |
# | \__/\| |   |  __/| (_| || |_ |  __/ | \__/\| (_) || || (_) || |    | |_\ \| |   | (_| || (_| |
#  \____/|_|    \___| \__,_| \__| \___|  \____/ \___/ |_| \___/ |_|     \____/|_|    \__,_| \__,_|
# Create a gradient from a list of colors and a step size
def create_color_grad(colors=[],steps=2):
    """Create a list of colors which gradually transition between the input colors with the input number of steps"""
    if steps == 1:
        return [colors[0]]
    if type(steps) != int:
        alert("Input an integer for steps variable")
        return None

    red = []
    blue = []
    green = []
    for color in colors:
        red.append(color.Red)
        blue.append(color.Blue)
        green.append(color.Green)

    stepspercolor = int(math.ceil(steps*1.0 / (len(colors)-1)))
    lastcolor = colors[-1]

    gradient = []

    for i in range(len(colors)-1):
        color1 = colors[i]
        color2 = colors[i+1]
        color1rgb = [color1.Red,color1.Green,color1.Blue]
        color2rgb = [color2.Red,color2.Green,color2.Blue]
        rgbchange = [color2rgb[0]-color1rgb[0],color2rgb[1]-color1rgb[1],color2rgb[2]-color1rgb[2]]
        stepchange = [rgbchange[0]*1.0/(stepspercolor),rgbchange[1]*1.0/(stepspercolor),rgbchange[2]*1.0/(stepspercolor)]

        step = 1
        gradient.append(color1)
        old_color = color1rgb
        while step < stepspercolor:
            new_color = [old_color[0]+stepchange[0],old_color[1]+stepchange[1],old_color[2]+stepchange[2]]
            old_color = new_color
            gradient.append(Color(red=new_color[0],green=new_color[1],blue=new_color[2]))
            step += 1

    if len(gradient) < steps:
        gradient.append(lastcolor)

    return gradient

#  _____         _               ______  _        _
# /  __ \       | |              | ___ \(_)      | |
# | /  \/  ___  | |  ___   _ __  | |_/ / _   ___ | | __  ___  _ __
# | |     / _ \ | | / _ \ | '__| |  __/ | | / __|| |/ / / _ \| '__|
# | \__/\| (_) || || (_) || |    | |    | || (__ |   < |  __/| |
#  \____/ \___/ |_| \___/ |_|    \_|    |_| \___||_|\_\ \___||_|
def color_picker(index):
    """Pick a color set based on index:
    0 - Reds
    1 - Blues
    2 - Oranges
    3 - Greens
    4 - Purples
    5 - Cyans
    """
    colorsets = [
        [Color(104,14,14),Color(255,0,0)],      # Reds
        [Color(20,20,77),Color(0,0,255)],       # Blues
        [Color(69,43,12),Color(252,140,3)],     # Oranges
        [Color(34,69,23),Color(65,252,3)],      # Greens
        [Color(74,10,69),Color(242,5,222)],     # Purples
        [Color(22,79,77),Color(0,252,242)]      # Cyans
    ]

    return colorsets[index]

#  _   _             _         _           ______                      _____         _
# | | | |           | |       | |         |___  /                     /  __ \       | |
# | | | | _ __    __| |  __ _ | |_   ___     / /   ___   _ __    ___  | /  \/  ___  | |  ___   _ __  ___
# | | | || '_ \  / _` | / _` || __| / _ \   / /   / _ \ | '_ \  / _ \ | |     / _ \ | | / _ \ | '__|/ __|
# | |_| || |_) || (_| || (_| || |_ |  __/ ./ /___| (_) || | | ||  __/ | \__/\| (_) || || (_) || |   \__ \
#  \___/ | .__/  \__,_| \__,_| \__| \___| \_____/ \___/ |_| |_| \___|  \____/ \___/ |_| \___/ |_|   |___/
#        | |
#        |_|
def update_zone_colors(active_view,filledregions):
    """Updates filled region colors based on zones"""
    allexistingprefixes = []
    allexistingzones = {}
    for region in filledregions:
        regionzone = region.LookupParameter("MEPCE HVAC Zone").AsString()
        if regionzone:
            zoneparts = regionzone.split('-')
            del zoneparts[-1]
            deriveprefix = ""
            for part in zoneparts:
                deriveprefix = deriveprefix + part + "-"
            allexistingprefixes.append(deriveprefix)
            try:
                lookup = allexistingzones[regionzone]
            except:
                allexistingzones[regionzone] = []
            allexistingzones[regionzone].append(region)

    # Get unique prefixes
    alluniqueprefixes = list(set(allexistingprefixes))
    alluniqueprefixes.sort()

    # Get the Solid Fill pattern (used for solid fills in Revit)
    solid_fill = None
    collector = DB.FilteredElementCollector(revit.doc).OfClass(FillPatternElement)
    for fpe in collector:
        pattern = fpe.GetFillPattern()
        if pattern.IsSolidFill:
            solid_fill = fpe
            break

    # Assign colors per zone group
    for z_idx,zoneprefix in enumerate(alluniqueprefixes):
        color = color_picker(z_idx)
        keyswithpre = []
        for key in allexistingzones.keys():
            split = key.split("-")
            del split[-1]
            keyprefix = ""
            for part in split:
                keyprefix = keyprefix + part + "-"
            if keyprefix == zoneprefix:
                keyswithpre.append(key)

        zonecount = len(keyswithpre)
        gradient = create_color_grad(color,zonecount)

        for c_idx,zone in enumerate(keyswithpre):
            if len(gradient) != len(keyswithpre):
                alert("Error creating color gradients for zone group {}!".format(zoneprefix))
                continue

            for region in allexistingzones[zone]:
                ogs = OverrideGraphicSettings()
                icolor = gradient[c_idx]
                ogs.SetProjectionLineColor(Color(0,0,0))
                ogs.SetSurfaceForegroundPatternColor(icolor)
                ogs.SetSurfaceForegroundPatternId(solid_fill.Id)
                ogs.SetSurfaceBackgroundPatternId(DB.ElementId.InvalidElementId)
                active_view.SetElementOverrides(region.Id,ogs)
    return

#  _____        _     _                                  _      ___                 _  _         _      _
# |  __ \      | |   | |                                | |    / _ \               (_)| |       | |    | |
# | |  \/  ___ | |_  | |      ___  __      __  ___  ___ | |_  / /_\ \__   __  __ _  _ | |  __ _ | |__  | |  ___
# | | __  / _ \| __| | |     / _ \ \ \ /\ / / / _ \/ __|| __| |  _  |\ \ / / / _` || || | / _` || '_ \ | | / _ \
# | |_\ \|  __/| |_  | |____| (_) | \ V  V / |  __/\__ \| |_  | | | | \ V / | (_| || || || (_| || |_) || ||  __/
#  \____/ \___| \__| \_____/ \___/   \_/\_/   \___||___/ \__| \_| |_/  \_/   \__,_||_||_| \__,_||_.__/ |_| \___|
# Gets lowest unused (positive) integer from list
def get_lowest_available(numbers=[]):
    """Get the lowest unused integer from a list of integers"""
    uniquenums = list(set(numbers))
    sortednums = sorted(uniquenums)
    for num_idx,num in enumerate(sortednums):
        if num != num_idx + 1:
            return num_idx + 1
    return len(uniquenums)+1