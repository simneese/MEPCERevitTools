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