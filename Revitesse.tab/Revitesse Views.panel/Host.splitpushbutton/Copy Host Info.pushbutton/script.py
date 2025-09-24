from pyrevit import revit, DB, forms
import os
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import ElementId
from System.Collections.Generic import List

paramName1 = "Revitesse Host ID"
paramName2 = "Revitesse Host Info"
paramGroupName = "Revitesse"

doc = __revit__.ActiveUIDocument.Document
app = doc.Application
uidoc = revit.uidoc

t1 = DB.Transaction(doc, "Setting Up Parameters")
t2 = DB.Transaction(doc, "Copy Host Info")

bindings = doc.ParameterBindings

#Part 1: Checking if there is a shared parameter file, if the parameter exists, and if the parameter is bound
# Check if a parameter is already bound to model categories
def isParameterBound(parameterName):
    it = bindings.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        if it.Key.Name.strip().lower() == parameterName.strip().lower(): return True
    return False

# Setup shared parameter file
sharedParameterFilePath = app.SharedParametersFilename
sharedParameterFile = app.OpenSharedParameterFile()

if not sharedParameterFilePath or not sharedParameterFile:
    revitPath = doc.PathName
    folder = os.path.dirname(revitPath) if revitPath else os.environ.get("TEMP")
    sharedParameterFilePath = os.path.join(folder, "sharedParameters.txt")

    if not os.path.exists(sharedParameterFilePath):
        with open(sharedParameterFilePath, 'w'): pass

    app.SharedParametersFilename = sharedParameterFilePath
    sharedParameterFile = app.OpenSharedParameterFile()
    if not sharedParameterFile: forms.alert("Unable to create or open shared parameter file.", exitscript=True)

# Get or create group
group = next((g for g in sharedParameterFile.Groups if g.Name == paramGroupName), None)
if not group: group = sharedParameterFile.Groups.Create(paramGroupName)

# Get or create definitions
definition1 = next((d for d in group.Definitions if d.Name == paramName1), None)
definition2 = next((d for d in group.Definitions if d.Name == paramName2), None)

if not definition1:
    opt1 = DB.ExternalDefinitionCreationOptions(paramName1, DB.SpecTypeId.String.Text)
    definition1 = group.Definitions.Create(opt1)

if not definition2:
    opt2 = DB.ExternalDefinitionCreationOptions(paramName2, DB.SpecTypeId.String.Text)
    definition2 = group.Definitions.Create(opt2)

# Bind parameters to all categories
categories = DB.CategorySet()
for cat in doc.Settings.Categories:
    if cat.AllowsBoundParameters: categories.Insert(cat)

binding = DB.InstanceBinding(categories)

t1.Start()
for definition in [definition1, definition2]:
    if not isParameterBound(definition.Name):
        success = bindings.Insert(definition, binding, DB.GroupTypeId.Text)
        if not success: bindings.ReInsert(definition, binding, DB.GroupTypeId.Text)
    else: bindings.ReInsert(definition, binding, DB.GroupTypeId.Text)
t1.Commit()

# Get host of an element
def getHostElement(elem):
    try: return elem.Host
    except AttributeError:
        try: return elem.SuperComponent
        except: return None

# Copy host ID
def copyHostId(elem):
    host = getHostElement(elem)
    param = elem.LookupParameter(paramName1)
    if not param: return
    if host: param.Set(str(host.Id))
    else: param.Set("No host")

# Copy host info parameter
def copyHostInfo(elem, chosenParameterName):
    host = getHostElement(elem)
    param = elem.LookupParameter(paramName2)
    if not param or not host: return
    hostParameter = host.LookupParameter(chosenParameterName)
    if not hostParameter: param.Set("Unidentified")
    else:
        try: value = hostParameter.AsValueString() or hostParameter.AsString() or str(hostParameter.AsInteger())
        except: value = "Unidentified"
        param.Set(value)

# Ask for element selection
try:
    with forms.WarningBar(title="Select exactly one element"):
        pickedObject = uidoc.Selection.PickObject(ObjectType.Element)
        sourceElement = doc.GetElement(pickedObject.ElementId)
except: forms.alert("No element selected.", exitscript=True)

# Part 2: Copying the host parameter and the host ID to the hosted element
# 1. Ask user which host parameter to copy FIRST
host = getHostElement(sourceElement)
if not host: forms.alert("The selected element has no host.", exitscript=True)

hostParameters = [p for p in host.Parameters if p.Definition and p.HasValue]
parameterNames = sorted([p.Definition.Name for p in hostParameters])

chosenParameter = forms.SelectFromList.show(parameterNames, title="Select Host Parameter to Copy", multiselect=False)
if not chosenParameter: forms.alert("No host parameter selected.", exitscript=True)

# 2. Ask user for target scope
scopeChoice = forms.ask_for_one_item(
    ["Only this element", "All instances of the same category visible in view", "All instances of the same category in entire project"],
    default="Only this element", prompt="Which elements to apply host info to:")

if not scopeChoice: forms.alert("No option selected.", exitscript=True)

# 3. Get target elements
def getTargetElements(scope, sourceElement):
    categoryId = sourceElement.Category.Id
    if scope == "Only this element": return [sourceElement]
    elif scope == "All instances of the same category visible in view":
        return [
            e for e in DB.FilteredElementCollector(doc, doc.ActiveView.Id)
            .WhereElementIsNotElementType()
            if e.Category and e.Category.Id == categoryId
        ]
    elif scope == "All instances of the same category in entire project":
        return [
            e for e in DB.FilteredElementCollector(doc)
            .WhereElementIsNotElementType()
            if e.Category and e.Category.Id == categoryId
        ]
    return [sourceElement]

targetElements = getTargetElements(scopeChoice, sourceElement)

# 4. Copy host info
t2.Start()
for elem in targetElements:
    copyHostId(elem)
    copyHostInfo(elem, chosenParameter)
t2.Commit()

forms.alert("Host information copied to {} elements.".format(len(targetElements)))