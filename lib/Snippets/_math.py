# -*- coding: utf-8 -*-

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
from Autodesk.Revit.DB import *

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