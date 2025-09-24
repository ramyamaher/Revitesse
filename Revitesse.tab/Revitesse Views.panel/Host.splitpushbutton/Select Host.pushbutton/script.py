from pyrevit import revit, forms
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import ElementId
from System.Collections.Generic import List

# Get the selected elements
selection = revit.uidoc.Selection.GetElementIds()

# Function to get element hosts
def getHostElement(elem):
    host = getattr(elem, 'Host', None)
    if host: return host
    if hasattr(elem, 'SuperComponent') and elem.SuperComponent: return elem.SuperComponent
    return None
    
# Function to check if elements have host elements, and zoom in on them
def zoomInOnHosts(selection):
    #Make an empty list for the host IDs    
    hostIds = []
    #Check if there is a host for every selected element
    for eid in selection:
        elem = revit.doc.GetElement(eid)
        host = getHostElement(elem)
        if host: hostIds.append(host.Id)
    if hostIds:
        revit.uidoc.ShowElements(List[ElementId](hostIds))
        revit.uidoc.Selection.SetElementIds(List[ElementId](hostIds))
    # If the selected elements have no hosts:
    else: forms.alert('There are no hosts for any of the selected elements.', exitscript=True)

# If elements have already been selected, zoom in on the host
if selection: zoomInOnHosts(selection)

# No elements selected. Prompt user to select elements with warning bar; get IDs of selection
else:
    try:
        with forms.WarningBar(title="Pick at least one element first:"):
            selectPrompt=revit.uidoc.Selection.PickObjects(ObjectType.Element)
            selection = List[ElementId]([ref.ElementId for ref in selectPrompt])
            zoomInOnHosts(selection)
    except: pass