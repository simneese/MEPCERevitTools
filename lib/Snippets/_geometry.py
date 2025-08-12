# -*- coding: utf-8 -*-

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
from Autodesk.Revit.DB import *

from types import NoneType

from pyrevit.forms import alert

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#==================================================
doc   = __revit__.ActiveUIDocument.Document #type: Document
uidoc = __revit__.ActiveUIDocument          #type: UIDocument
app   = __revit__.Application               #type: Application

# ______                 _              _     _____                             _            ______
# | ___ \               (_)            | |   /  __ \                           | |          |___  /
# | |_/ / _ __   ___     _   ___   ___ | |_  | /  \/ _   _  _ __ __   __  ___  | |_   ___      / /
# |  __/ | '__| / _ \   | | / _ \ / __|| __| | |    | | | || '__|\ \ / / / _ \ | __| / _ \    / /
# | |    | |   | (_) |  | ||  __/| (__ | |_  | \__/\| |_| || |    \ V / |  __/ | |_ | (_) | ./ /___
# \_|    |_|    \___/   | | \___| \___| \__|  \____/ \__,_||_|     \_/   \___|  \__| \___/  \_____/
#                      _/ |
#                     |__/
# Project curve to input Z elevation
def project_to_level(curve, elev):
    """
    Project input curve to input Z elevation
    """
    start = curve.GetEndPoint(0)
    end = curve.GetEndPoint(1)
    start = XYZ(start.X, start.Y, elev)
    end = XYZ(end.X, end.Y, elev)
    return Line.CreateBound(start, end)

#  _____        _     _____         _  _      _   _____                                _
# |  __ \      | |   /  ___|       | |(_)    | | |  __ \                              | |
# | |  \/  ___ | |_  \ `--.   ___  | | _   __| | | |  \/  ___   ___   _ __ ___    ___ | |_  _ __  _   _
# | | __  / _ \| __|  `--. \ / _ \ | || | / _` | | | __  / _ \ / _ \ | '_ ` _ \  / _ \| __|| '__|| | | |
# | |_\ \|  __/| |_  /\__/ /| (_) || || || (_| | | |_\ \|  __/| (_) || | | | | ||  __/| |_ | |   | |_| |
#  \____/ \___| \__| \____/  \___/ |_||_| \__,_|  \____/ \___| \___/ |_| |_| |_| \___| \__||_|    \__, |
#                                                                                                  __/ |
#                                                                                                 |___/
# Get solid geometry from an input element
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

# ______                           _  _                ______                           _____         _                                 _
# | ___ \                         | |(_)               | ___ \                         |_   _|       | |                               | |
# | |_/ /  ___   _   _  _ __    __| | _  _ __    __ _  | |_/ /  ___  __  __  ___  ___    | |   _ __  | |_   ___  _ __  ___   ___   ___ | |_
# | ___ \ / _ \ | | | || '_ \  / _` || || '_ \  / _` | | ___ \ / _ \ \ \/ / / _ \/ __|   | |  | '_ \ | __| / _ \| '__|/ __| / _ \ / __|| __|
# | |_/ /| (_) || |_| || | | || (_| || || | | || (_| | | |_/ /| (_) | >  < |  __/\__ \  _| |_ | | | || |_ |  __/| |   \__ \|  __/| (__ | |_
# \____/  \___/  \__,_||_| |_| \__,_||_||_| |_| \__, | \____/  \___/ /_/\_\ \___||___/  \___/ |_| |_| \__| \___||_|   |___/ \___| \___| \__|
#                                                __/ |
#                                               |___/
# Check if two bounding boxes intersect
def bounding_boxes_intersect(bb1, bb2):
    """Check if two bounding boxes intersect"""
    if not bb1 or not bb2:
        return False
    return (bb1.Max.X >= bb2.Min.X and bb1.Min.X <= bb2.Max.X and
            bb1.Max.Y >= bb2.Min.Y and bb1.Min.Y <= bb2.Max.Y and
            bb1.Max.Z >= bb2.Min.Z and bb1.Min.Z <= bb2.Max.Z)

# ______  _         _    _                  _____         _  _      _
# |  ___|| |       | |  | |                /  ___|       | |(_)    | |
# | |_   | |  __ _ | |_ | |_   ___  _ __   \ `--.   ___  | | _   __| | ___
# |  _|  | | / _` || __|| __| / _ \| '_ \   `--. \ / _ \ | || | / _` |/ __|
# | |    | || (_| || |_ | |_ |  __/| | | | /\__/ /| (_) || || || (_| |\__ \
# \_|    |_| \__,_| \__| \__| \___||_| |_| \____/  \___/ |_||_| \__,_||___/
# Flatten a list to only contain solids
def flatten_solids(items):
    """Flatten a list of solids to only contain solid objects, and no nested lists."""
    flat = []
    if isinstance(items, Solid):
        flat.append(items)
    elif isinstance(items, list) or isinstance(items, tuple):
        for i in items:
            flat.extend(flatten_solids(i))  # recursion
    return flat

# ______                 _              _    ______         _         _     _           ______  _
# | ___ \               (_)            | |   | ___ \       (_)       | |   | |          | ___ \| |
# | |_/ / _ __   ___     _   ___   ___ | |_  | |_/ /  ___   _  _ __  | |_  | |_   ___   | |_/ /| |  __ _  _ __    ___
# |  __/ | '__| / _ \   | | / _ \ / __|| __| |  __/  / _ \ | || '_ \ | __| | __| / _ \  |  __/ | | / _` || '_ \  / _ \
# | |    | |   | (_) |  | ||  __/| (__ | |_  | |    | (_) || || | | || |_  | |_ | (_) | | |    | || (_| || | | ||  __/
# \_|    |_|    \___/   | | \___| \___| \__| \_|     \___/ |_||_| |_| \__|  \__| \___/  \_|    |_| \__,_||_| |_| \___|
#                      _/ |
#                     |__/
# Project input point to input plane
def project_point_to_plane(point, plane_origin, plane_normal):
    # vector from origin to point
    vec = point - plane_origin
    distance = vec.DotProduct(plane_normal)
    return point - distance * plane_normal

#  _____        _      ___                                          ______         _         _
# |  __ \      | |    / _ \                                         | ___ \       (_)       | |
# | |  \/  ___ | |_  / /_\ \__   __  ___  _ __   __ _   __ _   ___  | |_/ /  ___   _  _ __  | |_
# | | __  / _ \| __| |  _  |\ \ / / / _ \| '__| / _` | / _` | / _ \ |  __/  / _ \ | || '_ \ | __|
# | |_\ \|  __/| |_  | | | | \ V / |  __/| |   | (_| || (_| ||  __/ | |    | (_) || || | | || |_
#  \____/ \___| \__| \_| |_/  \_/   \___||_|    \__,_| \__, | \___| \_|     \___/ |_||_| |_| \__|
#                                                       __/ |
#                                                      |___/
# Get average of two points
def get_point_average(MIN,MAX):
    """Input MIN and MAX points to get average point (can be for a bounding box or any two points)"""
    avgx = (MIN.X + MAX.X)/2
    avgy = (MIN.Y + MAX.Y)/2
    avgz = (MIN.Z + MAX.Z)/2
    return XYZ(avgx,avgy,avgz)