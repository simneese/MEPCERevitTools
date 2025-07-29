# -*- coding: utf-8 -*-
__title__   = "Cleanup Room Tags"
__doc__     = """Version = 1.0
Date    = 2025.07.21
________________________________________________________________
Description:
Deletes blank tags and replaces them with correct tags.

Useful for when an architect moves or deletes a room, causing the
room tag to disconnect.
________________________________________________________________
How-To:
> Select your active view
> Click button
________________________________________________________________
TODO:
Filter Tagged Rooms
Create New Tags
Make apply to all views, not just active (may take longer to compute)
________________________________________________________________
Last Updates:
- [2025.07.21] 


________________________________________________________________
Author: Jacob Tarbet"""

# IMPORTS
#==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
from Autodesk.Revit.Creation import *
from pyrevit import forms,revit,script,DB
from pyrevit.forms import alert


#.NET Imports
import clr   #Iron Python Core module. Allows for revit to read Python.
clr.AddReference('System')
from System.Collections.Generic import List


# VARIABLES
#==================================================
app    = __revit__.Application
uidoc  = __revit__.ActiveUIDocument
doc    = __revit__.ActiveUIDocument.Document #type:Document


# FUNCTIONS
from Snippets._selection import get_elements_of_categories


# MAIN
#==================================================

# Get Linked Documents
linked_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
linked_docs = []
for linked_instance in linked_instances:
    linked_docs.append(linked_instance.GetLinkDocument())

# Get the active view's level (assuming a floor plan or similar)
active_view = doc.ActiveView
active_level = active_view.GenLevel

if not active_level:
    alert("Active view does not have an associated level!",exitscript=True)


# Get rooms on active level
allrooms = get_elements_of_categories(categories=[BuiltInCategory.OST_Rooms],linked=True,readout=True)
rooms_on_active_level = [room for room in allrooms if room.LevelId == active_level.Id]

# Get room tags in view
# Filter all SpatialElementTags (includes room tags, space tags, area tags)
room_tags = FilteredElementCollector(doc).OfClass(SpatialElementTag).WhereElementIsNotElementType().ToElements()


# Filter to include only Room tags (optional: check if tag is for a Room)
room_tags_in_view = [tag for tag in room_tags if isinstance(tag.TaggedRoomId, ElementId)]

# Get Room Names and Numbers
room_number_and_name = {}

for room in allrooms:
     room_number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
     room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
     room_number_and_name[room_number] = room_name
room_numkey = room_number_and_name.keys()

# ================Delete Disassociated Tags========================

alltags = get_elements_of_categories(categories=[BuiltInCategory.OST_RoomTags],readout=True)
tagtext = [tag.TagText for tag in alltags]
dtags = {}
for i in range(len(tagtext)):
    tagtx = tagtext[i]
    tagel = alltags[i]
    dtags[tagtx] = tagel
dtags_keys = dtags.keys()

# Deletes orphaned (???) tags and removes them from the list (works)
with revit.Transaction("Delete ??? tag",doc):
    for key in dtags_keys:
        if '?' in key:
            doc.Delete(dtags[key].Id)
            del dtags[key]


    # ================Create New Tags========================

    tag_types = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RoomTags).OfClass(FamilySymbol).ToElements()
    for tag in tag_types:
        tag_name = tag.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsValueString()
        if tag_name == "MEPCE - Room Tag sml: Room Name & Number":
            tag_type = tag
            break
        else:
            tag_type = None
    if tag_type is None:
        alert("RoomTag type not found!",exitscript=True)
    
    tag_type.Activate()
    
 # Compare room and tag names and create a new tag on the unmatched room
    for i in range(len(room_number_and_name)):
        roomname = room_number_and_name[i]
        room = allrooms[i]
        tagcount = 0
        for tag in tagtext:
            if roomname == tag:
                tagcount += 1
                break
        if tagcount == 0: 
            #reference = Reference(room).CreateLinkReference(linked_instance)
            reference = Reference.CreateLinkReference(linked_instance, room)
            room_tag = IndependentTag.Create(doc,tag_type.Id,active_view.Id,reference,True,TagOrientation.Horizontal,XYZ())
            room_tag.LeaderEndCondition = LeaderEndCondition.Free


         
            print room_tag
        
    

#==================================================