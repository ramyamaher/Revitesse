from pyrevit import revit, DB, forms
from Autodesk.Revit.UI.Selection import ObjectType
import clr, sys, os

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
from System.Windows import Window, Thickness, HorizontalAlignment, WindowStartupLocation, TextWrapping
from System.Windows.Controls import StackPanel, ComboBox, Label, Button, Orientation, TextBox, TextBlock

doc = __revit__.ActiveUIDocument.Document
uidoc = revit.uidoc
app = doc.Application

paramGroupName = "Revitesse"
parameterName = "Revitesse Combined Parameters"
maximumRows = 5

# Check if a parameter is already bound to model categories
def isParameterBound(parameterName):
    bindings = doc.ParameterBindings
    it = bindings.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        if it.Key.Name.strip().lower() == parameterName.strip().lower(): return True
    return False

# Bind the shared parameter to all categories that allow binding
def bindSharedParameter():
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

    group = next((g for g in sharedParameterFile.Groups if g.Name == paramGroupName), None)
    if not group: group = sharedParameterFile.Groups.Create(paramGroupName)

    definition = next((d for d in group.Definitions if d.Name == parameterName), None)
    if not definition:
        options = DB.ExternalDefinitionCreationOptions(parameterName, DB.SpecTypeId.String.Text)
        definition = group.Definitions.Create(options)

    categories = DB.CategorySet()
    for cat in doc.Settings.Categories:
        try:
            if cat.AllowsBoundParameters: categories.Insert(cat)
        except: pass

    binding = DB.InstanceBinding(categories)

    t = DB.Transaction(doc, "Bind Revitesse Combined Parameters")
    t.Start()
    if not isParameterBound(parameterName): doc.ParameterBindings.Insert(definition, binding, DB.GroupTypeId.Text)
    t.Commit()
    return definition

# Get all parameters for a given category, sorted, including "<None>" option
def getCategoryParameters(category):
    params = set()
    collector = DB.FilteredElementCollector(doc).OfCategoryId(category.Id).WhereElementIsNotElementType()
    for elem in collector:
        for p in elem.Parameters:
            if p.Definition and p.Definition.Name: params.add(p.Definition.Name)
    parameterList = sorted(params)
    parameterList.insert(0, "<None>")
    return parameterList

# Window class for the combine parameters UI
class CombineParamsForm(Window):
    def __init__(self, category):
        self.Title = "Combine Parameters"
        self.Width, self.Height = 630, 330
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True

        self.category = category
        self.paramNames = getCategoryParameters(category)

        panel = StackPanel()
        panel.Margin = Thickness(10)

        self.paramCombos = []
        self.sepTextboxes = []

        for i in range(maximumRows):
            rowPanel = StackPanel(Orientation=Orientation.Horizontal, Margin=Thickness(0, 5, 0, 5))
            rowPanel.Height = 25
            labelParam = Label(Content="Select Parameter:")
            labelParam.Width = 120
            rowPanel.Children.Add(labelParam)

            combo = ComboBox(ItemsSource=self.paramNames, Width=220)
            combo.SelectedIndex = 0
            rowPanel.Children.Add(combo)
            self.paramCombos.append(combo)

            labelSep = Label(Content="Separator after value:")
            labelSep.Margin = Thickness(10,0,0,0)
            labelSep.Width = 138
            rowPanel.Children.Add(labelSep)

            txt = TextBox(Width=100)
            rowPanel.Children.Add(txt)
            self.sepTextboxes.append(txt)

            panel.Children.Add(rowPanel)

        infoText = TextBlock(Text="At least two parameters must be selected.\nIf separator is empty, a space will be used.", 
                             Margin=Thickness(0,10,0,10), TextWrapping=TextWrapping.Wrap)
        panel.Children.Add(infoText)

        buttonPanel = StackPanel(Orientation=Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Right, Margin=Thickness(0,10,0,0))
        self.cancelButton = Button(Content="Cancel", Width=80, Height=22, Margin=Thickness(5))
        self.cancelButton.Click += self.cancelClicked
        self.applyButton = Button(Content="Apply", Width=80, Height=22, Margin=Thickness(5))
        self.applyButton.Click += self.applyClicked
        buttonPanel.Children.Add(self.cancelButton)
        buttonPanel.Children.Add(self.applyButton)

        panel.Children.Add(buttonPanel)

        self.Content = panel
        self.result = None

    def applyClicked(self, sender, args):
        selectedParameters = [c.SelectedItem for c in self.paramCombos if c.SelectedItem != "<None>"]
        if len(selectedParameters) < 2:
            forms.alert("Please select at least two parameters.")
            return
        separators = []
        for txt in self.sepTextboxes:
            sep = txt.Text.strip()
            separators.append(sep if sep != "" else " ")
        separators = [separators[i] for i, c in enumerate(self.paramCombos) if c.SelectedItem != "<None>"]
        if len(separators) > len(selectedParameters) -1: separators = separators[:len(selectedParameters)-1]
        self.result = (selectedParameters, separators)
        self.DialogResult = True

    def cancelClicked(self, sender, args):
        self.result = None
        self.DialogResult = False

def getParameterValue(elem, param_name):
    param = None
    for p in elem.Parameters:
        if p.Definition and p.Definition.Name == param_name:
            param = p
            break
    if not param: return ""
    try:
        val = param.AsString()
        if val is None: val = param.AsValueString()
        if val is None: val = ""
        return val
    except:  return ""

def main():
    try:
        with forms.WarningBar(title="Select one element to use as reference for category"):
            picked = uidoc.Selection.PickObject(ObjectType.Element, "Pick an element")
            referenceElement = doc.GetElement(picked.ElementId)
    except:
        forms.alert("Selection cancelled.", exitscript=True)
        sys.exit()

    category = referenceElement.Category
    if not category:
        forms.alert("Selected element has no category.", exitscript=True)
        sys.exit()

    bindSharedParameter()

    form = CombineParamsForm(category)
    if not form.ShowDialog():
        forms.alert("Operation cancelled.", exitscript=True)
        sys.exit()

    selectedParameters, separators = form.result

    collector = DB.FilteredElementCollector(doc).OfCategoryId(category.Id).WhereElementIsNotElementType()

    combinedCount = 0
    t = DB.Transaction(doc, "Combine Parameters into '{}'".format(parameterName))
    t.Start()
    for elem in collector:
        combinedValues = []
        for i, param_name in enumerate(selectedParameters):
            val = getParameterValue(elem, param_name)
            combinedValues.append(val)
            if i < len(selectedParameters) - 1:  combinedValues.append(separators[i])
        combinedText = "".join(combinedValues).strip()

        combinedParameters = None
        for p in elem.Parameters:
            if p.Definition and p.Definition.Name == parameterName:
                combinedParameters = p
                break
        if combinedParameters and not combinedParameters.IsReadOnly:
            combinedParameters.Set(combinedText)
            combinedCount += 1
    t.Commit()
    forms.alert("{} elements updated with combined parameters.".format(combinedCount), title="Done")

if __name__ == "__main__": main()