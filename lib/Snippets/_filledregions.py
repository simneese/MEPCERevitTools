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

#  _____        _    ______  _  _  _            _  ______               _
# |  __ \      | |   |  ___|(_)| || |          | | | ___ \             (_)
# | |  \/  ___ | |_  | |_    _ | || |  ___   __| | | |_/ /  ___   __ _  _   ___   _ __   ___
# | | __  / _ \| __| |  _|  | || || | / _ \ / _` | |    /  / _ \ / _` || | / _ \ | '_ \ / __|
# | |_\ \|  __/| |_  | |    | || || ||  __/| (_| | | |\ \ |  __/| (_| || || (_) || | | |\__ \
#  \____/ \___| \__| \_|    |_||_||_| \___| \__,_| \_| \_| \___| \__, ||_| \___/ |_| |_||___/
#                                                                 __/ |
#                                                                |___/
# This function will get a list of all filled regions, which can optionally be filtered by view discipline

def get_filled_regions(allviews, filter_discipline=""):
    # Filter views by input filter
    filtviews = []
    for view in allviews:
        if view.IsTemplate:
            continue
        disciplineparam = view.LookupParameter("DISCIPLINE")
        if type(disciplineparam) is NoneType:
            continue
        else:
            discipline = disciplineparam.AsString()
        if type(discipline) is NoneType:
            continue
        elif str(filter_discipline).lower() not in discipline.lower():
            continue
        else:
            try:
                filtviews.append(view)
            except:
                continue

    #print("Number of Views: "+str(len(filtviews)))
    #print("View filter: "+str(filter_discipline))

    # Alert if no views found
    if not filtviews:
        alert('Unable to find views with discipline "'+str(filter_discipline)+'" in project!')

    # Get filtered list of filled regions
    filledregions = []
    if filter_discipline == "":
        filledregions = list(FilteredElementCollector(doc).OfClass(FilledRegion))
    else:
        for view in filtviews:
            regions = list(FilteredElementCollector(doc).OwnedByView(view.Id).OfClass(FilledRegion))
            filledregions.append(regions)

    if not filledregions:
        alert('Unable to find filled regions in views with discipline "'+str(filter_discipline)+'." Please ensure you have created filled regions.')

    # Define dictionary with filtered views and filled regions
    d = {"filtviews" : filtviews, "filledregions" : filledregions}
    return d

#  _____        _    ______
# |  __ \      | |   |  ___|
# | |  \/  ___ | |_  | |_     __ _   ___   ___  ___
# | | __  / _ \| __| |  _|   / _` | / __| / _ \/ __|
# | |_\ \|  __/| |_  | |    | (_| || (__ |  __/\__ \
#  \____/ \___| \__| \_|     \__,_| \___| \___||___/
# This function will convert filled regions into faces

def get_faces(filledregions):
    normalfaces = []        # List of all faces, grouped by view they appear in
    for view in filledregions:
        normfaces = []      # List of normal faces in current view
        for region in view:
            boundaries = region.GetBoundaries()     # Get filled region boundary curves
            geometry = GeometryCreationUtilities.CreateExtrusionGeometry(boundaries,XYZ(0,0,1),10)     # Extrude boundaries to get solid
            faces = geometry.Faces      # Get faces of the solid
            for i in range(faces.Size):
                normal = faces[i].ComputeNormal(UV(0,0))    # Get surface normal of each face
                if normal.Z == 1:
                    normfaces.append(faces[i])              # Collect face with normal = (0,0,1)
                else:
                    continue
        normalfaces.append(normfaces)
    return normalfaces

#  _____        _    ______
# |  __ \      | |   | ___ \
# | |  \/  ___ | |_  | |_/ /  ___    ___   _ __ ___   ___
# | | __  / _ \| __| |    /  / _ \  / _ \ | '_ ` _ \ / __|
# | |_\ \|  __/| |_  | |\ \ | (_) || (_) || | | | | |\__ \
#  \____/ \___| \__| \_| \_| \___/  \___/ |_| |_| |_||___/
# This function will get filled regions with room number parameter

def get_rooms():
    allfilledregions = list(FilteredElementCollector(doc).OfClass(FilledRegion))
    filtregions = []
    for region in allfilledregions:
        paramexists = region.LookupParameter("MEPCE Room Number")
        if paramexists:
            roomnum = region.LookupParameter("MEPCE Room Number").AsString()
            if type(roomnum) is NoneType:
                continue
            else:
                filtregions.append(region)
    return filtregions

#  _____        _    ______                           ______         _
# |  __ \      | |   | ___ \                          |  _  \       | |
# | |  \/  ___ | |_  | |_/ /  ___    ___   _ __ ___   | | | |  __ _ | |_   __ _
# | | __  / _ \| __| |    /  / _ \  / _ \ | '_ ` _ \  | | | | / _` || __| / _` |
# | |_\ \|  __/| |_  | |\ \ | (_) || (_) || | | | | | | |/ / | (_| || |_ | (_| |
#  \____/ \___| \__| \_| \_| \___/  \___/ |_| |_| |_| |___/   \__,_| \__| \__,_|
# Get MEPCE room parameters from room

def get_room_data(rooms):
    roomdata = []
    paramslookup = ["MEPCE Room Number", "MEPCE Room Name", "MEPCE HVAC Zone", "Area", "MEPCE Room Airflow",
                    "MEPCE Room Ventilation"]
    paramsnames = ["Room Number", "Room Name", "HVAC Zone", "Area (ft^2)", "Supply Airflow (CFM)", "Ventilation (CFM)"]
    for room in rooms:
        roomparams = dict()
        for i in range(len(paramslookup)):
            param = paramslookup[i]
            exists = room.LookupParameter(param)
            if type(exists) is not NoneType:
                paramtype = str(exists.StorageType)
                if paramtype == "String":
                    roomparams.update({paramsnames[i]: room.LookupParameter(param).AsString()})
                if paramtype == "Double":
                    if i == 4 or i == 5:
                        roomparams.update({paramsnames[i]: room.LookupParameter(param).AsDouble() * 60})
                    else:
                        roomparams.update({paramsnames[i]: room.LookupParameter(param).AsDouble()})
            else:
                roomparams.update({paramsnames[i]: None})

        if roomparams['Area (ft^2)']:
            areaflow = roomparams['Supply Airflow (CFM)'] / roomparams['Area (ft^2)']
        else:
            areaflow = 0
        roomparams.update({'Supply / Area (CFM/ft^2)': areaflow})

        if roomparams['Room Number'] and roomparams['Room Name']:
            combname = roomparams['Room Number'] + " " + roomparams['Room Name']
        else:
            combname = None
        roomparams.update({'Room Name and Number': combname})

        roomdata.append(roomparams)

    paramsnames.append('Supply / Area (CFM/ft^2)')
    paramsnames.insert(0,'Room Name and Number')
    return [roomdata,paramsnames] # Returns list of dictionaries and list of dictionary keys

#  _____                     _          ______  _  _  _            _  ______               _
# /  __ \                   | |         |  ___|(_)| || |          | | | ___ \             (_)
# | /  \/ _ __   ___   __ _ | |_   ___  | |_    _ | || |  ___   __| | | |_/ /  ___   __ _  _   ___   _ __
# | |    | '__| / _ \ / _` || __| / _ \ |  _|  | || || | / _ \ / _` | |    /  / _ \ / _` || | / _ \ | '_ \
# | \__/\| |   |  __/| (_| || |_ |  __/ | |    | || || ||  __/| (_| | | |\ \ |  __/| (_| || || (_) || | | |
#  \____/|_|    \___| \__,_| \__| \___| \_|    |_||_||_| \___| \__,_| \_| \_| \___| \__, ||_| \___/ |_| |_|
#                                                                                    __/ |
#                                                                                   |___/
# Create filled region from room

def create_filled_region_from_room(room, view, region_type):
    """
    Creates a filled region in `view` using the room's boundaries.
    Projects boundaries to view's level plane.
    Skips invalid or zero-area loops.
    """
    options = SpatialElementBoundaryOptions()
    boundary_loops = room.GetBoundarySegments(options)
    if not boundary_loops:
        return False

    level_elev = view.GenLevel.Elevation
    boundaries = []

    filledregion = {}
    for loop_idx, segment_list in enumerate(boundary_loops):
        curves = CurveLoop()
        prev_end = None
        for seg in segment_list:
            c = seg.GetCurve()
            start = c.GetEndPoint(0)
            end = c.GetEndPoint(1)

            # flatten to level elevation
            flat_start = XYZ(start.X, start.Y, level_elev)
            flat_end = XYZ(end.X, end.Y, level_elev)

            # reconstruct line
            line = Line.CreateBound(flat_start, flat_end)
            curves.Append(line)

            prev_end = flat_end

        # optionally skip degenerate loops

        boundaries.append(curves)
        filledregion[loop_idx] = FilledRegion.Create(doc, region_type.Id, view.Id, [curves])


    if not boundaries:
        return False
    return filledregion

#  _____        _     _____
# |  __ \      | |   /  __ \
# | |  \/  ___ | |_  | /  \/  ___   _ __  _ __    ___  _ __  ___
# | | __  / _ \| __| | |     / _ \ | '__|| '_ \  / _ \| '__|/ __|
# | |_\ \|  __/| |_  | \__/\| (_) || |   | | | ||  __/| |   \__ \
#  \____/ \___| \__|  \____/ \___/ |_|   |_| |_| \___||_|   |___/
# Get filled region corners
def get_filled_region_corners(filled_region):
    """Return unique XYZ corner points from a filled region's boundaries."""
    corners = []

    boundaries = filled_region.GetBoundaries()  # IList<IList<Curve>>
    for loop in boundaries:
        for curve in loop:
            corners.append(curve.GetEndPoint(0))
            corners.append(curve.GetEndPoint(1))

    # Deduplicate points
    unique_corners = []
    seen = set()
    for p in corners:
        key = (round(p.X, 6), round(p.Y, 6), round(p.Z, 6))
        if key not in seen:
            seen.add(key)
            unique_corners.append(p)

    return unique_corners