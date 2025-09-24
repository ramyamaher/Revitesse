from pyrevit import revit, DB, forms
from Autodesk.Revit.UI.Selection import ObjectType
import clr, csv, sys
clr.AddReference('PresentationFramework')
from System.Windows import Window, Thickness, WindowStartupLocation
from System.Windows.Controls import StackPanel, ComboBox, Label, Button, Orientation
from System.Collections.Generic import List
from Autodesk.Revit.DB import ElementId

# Get parameter by name
def getParameterByName(elem, name):
    for p in elem.Parameters:
        if p.Definition and p.Definition.Name == name: return p
    return None

# Parameter value as string dropdown to choose from
def getParameterValue(p):
    if p is None: return ""
    try: return p.AsValueString() or p.AsString() or str(p.AsInteger())
    except: return ""

# User selects an element to define the category to filter on
try:
    with forms.WarningBar(title="Select one element to filter the category of"):
        picked = revit.uidoc.Selection.PickObject(ObjectType.Element, "Pick an element")
        sourceElement = revit.doc.GetElement(picked.ElementId)
except Exception: forms.alert("Selection cancelled.", exitscript=True)

if not sourceElement or not sourceElement.Category: forms.alert("Selected element has no category.", exitscript=True)

categoryId = sourceElement.Category.Id

# Collect all elements of the same category (not element types)
allElements = DB.FilteredElementCollector(revit.doc).WhereElementIsNotElementType().ToElements()
categoryElements = [e for e in allElements if e.Category and e.Category.Id == categoryId]

# Collect all parameter names and their distinct values from these elements
parameterNames = []
parameterValuesByName = {}

for e in categoryElements:
    for p in e.Parameters:
        if p.Definition:
            name = p.Definition.Name
            if name not in parameterNames:
                parameterNames.append(name)
                parameterValuesByName[name] = set()
            val = getParameterValue(p)
            if val: parameterValuesByName[name].add(val)

parameterNames.sort()

# Build the UI form
class FilterByParameterForm(Window):
    def __init__(self):
        self.Title, self.Width, self.Height = "Filter Tags by Parameter", 450, 250
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.result = None

        def label(text): return Label(Content=text)

        panel = StackPanel(Margin=Thickness(10))

        panel.Children.Add(label("Select parameter to filter with:"))
        self.parameterCombo = ComboBox(ItemsSource=parameterNames)
        self.parameterCombo.SelectionChanged += self.updateValueCombo
        panel.Children.Add(self.parameterCombo)

        panel.Children.Add(label("Select parameter value:"))
        self.valueCombo = ComboBox()
        panel.Children.Add(self.valueCombo)

        panel.Children.Add(label("Scope:"))
        self.scopeCombo = ComboBox(ItemsSource=["Active View", "Entire Project"], SelectedIndex=0)
        panel.Children.Add(self.scopeCombo)

        # Buttons
        buttonPanel = StackPanel(Orientation=Orientation.Horizontal)
        buttons = [("Select", self.selectClicked, 100), ("Select & Export CSV", self.exportClicked, 140), ("Cancel", self.cancelClicked, 100)]
        for text, handler, width in buttons:
            btn = Button(Content=text, Width=width, Margin=Thickness(10))
            btn.Click += handler
            buttonPanel.Children.Add(btn)

        panel.Children.Add(buttonPanel)
        self.Content = panel

    def updateValueCombo(self, sender, args):
        param = self.parameterCombo.SelectedItem
        if param and param in parameterValuesByName: self.valueCombo.ItemsSource = sorted(parameterValuesByName[param])
        else: self.valueCombo.ItemsSource = []

    def getResults(self):
        if self.parameterCombo.SelectedItem and self.valueCombo.SelectedItem:
            return (self.parameterCombo.SelectedItem, self.valueCombo.SelectedItem, self.scopeCombo.SelectedItem)
        else:
            forms.alert("Please select both a parameter and a value.")
            return None

    def selectClicked(self, sender, args):
        res = self.getResults()
        if res:
            self.result = ("select", res)
            self.Close()

    def exportClicked(self, sender, args):
        res = self.getResults()
        if res:
            self.result = ("export", res)
            self.Close()

    def cancelClicked(self, sender, args):
        self.result = None
        self.Close()

# Show the form
form = FilterByParameterForm()
form.ShowDialog()

if not form.result:
    forms.alert("Operation cancelled.")
    sys.exit()

action, (paramName, paramValue, scope) = form.result

# Choose scope of elements to filter
if scope == "Active View": elemsToFilter = DB.FilteredElementCollector(revit.doc, revit.doc.ActiveView.Id).WhereElementIsNotElementType()
else: elemsToFilter = DB.FilteredElementCollector(revit.doc).WhereElementIsNotElementType()

# Filter elements by category and parameter value
matchedElements = []
for e in elemsToFilter:
    if e.Category and e.Category.Id == categoryId:
        p = getParameterByName(e, paramName)
        if getParameterValue(p) == paramValue: matchedElements.append(e)

if not matchedElements:
    forms.alert("No matching elements found.")
    sys.exit()

# Select and zoom to matched elements
elementIds = List[ElementId]([e.Id for e in matchedElements])
revit.uidoc.ShowElements(elementIds)
revit.uidoc.Selection.SetElementIds(elementIds)

# Export CSV if requested
if action == "export":
    csvPath = forms.save_file(file_ext='csv', title="Save CSV file")
    if csvPath:
        allParams = sorted({p.Definition.Name for e in matchedElements for p in e.Parameters if p.Definition})
        with open(csvPath, 'w') as f:
            writer = csv.writer(f)
            writer.writerow(allParams)
            for e in matchedElements:
                row = []
                for pname in allParams:
                    p = getParameterByName(e, pname)
                    row.append(getParameterValue(p))
                writer.writerow(row)
        forms.alert("Exported to: " + csvPath)