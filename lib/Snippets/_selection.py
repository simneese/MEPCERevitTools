# -*- coding: utf-8 -*-
from symbol import continue_stmt

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Mechanical import MechanicalSystemType, DuctSystemType
from Autodesk.Revit.DB.Plumbing import PipingSystemType
from types import NoneType

from pyrevit import revit, forms
from pyrevit.forms import alert

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#==================================================
doc   = __revit__.ActiveUIDocument.Document #type: Document
uidoc = __revit__.ActiveUIDocument          #type: UIDocument
app   = __revit__.Application               #type: Application

#  _____        _     _____        _              _              _   _____  _                                _
# |  __ \      | |   /  ___|      | |            | |            | | |  ___|| |                              | |
# | |  \/  ___ | |_  \ `--.   ___ | |  ___   ___ | |_   ___   __| | | |__  | |  ___  _ __ ___    ___  _ __  | |_  ___
# | | __  / _ \| __|  `--. \ / _ \| | / _ \ / __|| __| / _ \ / _` | |  __| | | / _ \| '_ ` _ \  / _ \| '_ \ | __|/ __|
# | |_\ \|  __/| |_  /\__/ /|  __/| ||  __/| (__ | |_ |  __/| (_| | | |___ | ||  __/| | | | | ||  __/| | | || |_ \__ \
#  \____/ \___| \__| \____/  \___||_| \___| \___| \__| \___| \__,_| \____/ |_| \___||_| |_| |_| \___||_| |_| \__||___/
# Get elements selected by user

def get_selected_elements(filter_types=None):
    """Get Selected Elements in Revit UI.
    You can provide a list of types for filter_types parameter (optional)

    e.g.
    sel_wall = get_selected_elements([Wall])"""
    selected_element_ids = uidoc.Selection.GetElementIds()
    selected_elements    = [doc.GetElement(e_id) for e_id in selected_element_ids]

    # Filter Selection (Optionally)
    if filter_types:
        return [el for el in selected_elements if type(el) in filter_types]
    return selected_elements

#  _____        _     _____  _                                _                  __   _____         _                              _
# |  __ \      | |   |  ___|| |                              | |                / _| /  __ \       | |                            (_)
# | |  \/  ___ | |_  | |__  | |  ___  _ __ ___    ___  _ __  | |_  ___    ___  | |_  | /  \/  __ _ | |_   ___   __ _   ___   _ __  _   ___  ___
# | | __  / _ \| __| |  __| | | / _ \| '_ ` _ \  / _ \| '_ \ | __|/ __|  / _ \ |  _| | |     / _` || __| / _ \ / _` | / _ \ | '__|| | / _ \/ __|
# | |_\ \|  __/| |_  | |___ | ||  __/| | | | | ||  __/| | | || |_ \__ \ | (_) || |   | \__/\| (_| || |_ |  __/| (_| || (_) || |   | ||  __/\__ \
#  \____/ \___| \__| \____/ |_| \___||_| |_| |_| \___||_| |_| \__||___/  \___/ |_|    \____/ \__,_| \__| \___| \__, | \___/ |_|   |_| \___||___/
#                                                                                                               __/ |
#                                                                                                              |___/
# Get elements of input categories, optionally filtered by SystemType and Level!

def get_elements_of_categories(categories=[],systemtypes=[],view=None,linked=False,readout=False):
    """Get elements of input categories, optionally filtered by SystemType and View!
    Currently only for Ducts and Pipes.
    Use separately for elements from links vs from the model!

    - Input "categories" should be list of desired builtin categories, eg: BuiltInCategory.OST_DuctCurves
    - Input "systemtypes" should be a list of type name strings, eg: 'Supply Air'. Leave empty to search all system types
    - Input "view" should be a view element. Leave empty to search all views
    - Input "linked" should be a boolean which will determine whether to look in linked docs. Default is False
    - Input "readout" should be a boolean which will determine if a readout is printed. Default is False

    eg:
    graded_pipes = get_elements_of_categories([BuiltInCategory.OST_PipeCurves],['PW-Grease Waste Pipe','PW-Waste Pipe','PW-Roof Drain'],view)"""
    if readout is True:
        print "~"*100+"\n-ELEMENT FILTERS-"
        print "\n\n"
        print "Linked Elements: {}".format(linked)
        print "\n\n"
        if view:
            viewexists = True
        else:
            viewexists = False
        print "Filtered by Active View: {}".format(viewexists)
        print "\n\n"
        print "Filtered Categories:\n"
        for category in categories:
            print category
        print "\n\n"

    # Get linked docs
    link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    link_docs = [link.GetLinkDocument() for link in link_instances]

    # Get view level
    if view:
        view_level_id = view.GenLevel.Id
    else:
        view_level_id = None

    collector = []
    if not categories:                                                      # Check if any categories have been added
        alert("Categories list is empty!",title = "Selection",exitscript = True)
    else:
        try:
            if linked is True:                                              # Get linked elements
                for link in link_docs:
                    if link:
                        if view:
                            for category in categories:
                                elements = FilteredElementCollector(link,view.Id).OfCategory(
                                    category).WhereElementIsNotElementType().ToElements()
                                for el in elements:
                                    try:
                                        viewfilter = VisibleInViewFilter(link,view.Id)
                                        if viewfilter.PassesFilter(el):
                                            collector.append(el)
                                    except:
                                        continue
                        else:
                            for category in categories:
                                elements = FilteredElementCollector(link).OfCategory(
                                category).WhereElementIsNotElementType().ToElements()
                                for el in elements:
                                    collector.append(el)
                    else:
                        continue
            else:
                if view:                                                        # Get all elements of categories in view
                    elements = [FilteredElementCollector(doc,view.Id).OfCategory(
                        category).WhereElementIsNotElementType().ToElements() for category in categories]
                    for category in elements:
                        for el in category:
                            try:
                                viewfilter = VisibleInViewFilter(doc, view.Id)
                                if viewfilter.PassesFilter(el):
                                    collector.append(el)
                            except:
                                continue
                else:                                                           # Get all elements of categories
                    elements = [FilteredElementCollector(doc).OfCategory(
                        category).WhereElementIsNotElementType().ToElements() for category in categories]
                    for category in elements:
                        for el in category:
                            collector.append(el)
        except Exception as e:
            alert(msg="Could not collect elements!", sub_msg="Error: {}".format(e),title="Selection", exitscript=True)

    #allelements = [el for category in collector for el in category]         # Flatten list of elements
    if view:
        allelements = [el for el in collector if hasattr(el,'LevelId') and el.LevelId == view_level_id]
    else:
        allelements = collector

    typfilter = []
    if systemtypes:                                                         # Check if system types have been added
        if readout is True:
            print "Filtered SystemTypes:"
            for type in systemtypes:
                print type
        for el in allelements:                                              # Filter by system type
            try:
                mep = el.MEPSystem.LookupParameter("Type").AsValueString()
                if mep in systemtypes:
                    typfilter.append(el)
                else:
                    continue
            except:
                continue
        filtered_elements = typfilter
    else:
        if readout is True:
            print 'No system filters added!'
        filtered_elements = allelements
    if readout is True:
        print "\n# Filtered Elements: {}".format(len(filtered_elements))+"\n"+"~"*100

    if not filtered_elements:
        alert("No elements passed the filter!",title="Selection",exitscript=True)
    return filtered_elements