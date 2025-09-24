from pyrevit import revit, forms
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import ElementId
from System.Collections.Generic import List

# Get the selected elements
selection = revit.uidoc.Selection.GetElementIds()

# If elements are selected, zoom in
if selection: revit.uidoc.ShowElements(selection)

# No elements selected. Prompt user to select elements with warning bar and get IDs of selection
else:
    try:
        with forms.WarningBar(title="Pick at least one element first:"):
            selectPrompt=revit.uidoc.Selection.PickObjects(ObjectType.Element)
            elementIds = List[ElementId]([ref.ElementId for ref in selectPrompt])

            # Show the elemenets in the view and keep them selected
            revit.uidoc.ShowElements(elementIds)
            revit.uidoc.Selection.SetElementIds(elementIds)

    except: pass