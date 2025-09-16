from pyrevit import revit, DB, forms
from Autodesk.Revit.UI.Selection import ObjectType
import clr
import sys

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
from System.Windows import Window, Thickness, HorizontalAlignment, WindowStartupLocation, TextWrapping
from System.Windows.Controls import StackPanel, ComboBox, Label, Button, Orientation, TextBlock

doc = __revit__.ActiveUIDocument.Document
uidoc = revit.uidoc

# This script has to be in the foreground to run
class ForegroundAlert(Window):
    def __init__(self, message, title="Alert"):
        self.Title, self.Width, self.Height = title, 300, 120
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True

        panel = StackPanel(Margin=Thickness(10))
        panel.Children.Add(TextBlock(Text=message, TextWrapping=TextWrapping.Wrap, Margin=Thickness(0, 0, 0, 10)))

        okButton = Button(Content="OK", Width=80, HorizontalAlignment=HorizontalAlignment.Center)
        okButton.Click += self.okClicked
        panel.Children.Add(okButton)

        self.Content = panel

    def okClicked(self, sender, args):
        self.DialogResult = True
        self.Close()

def showForegroundAlert(message, title="Alert"):
    alert = ForegroundAlert(message, title)
    alert.ShowDialog()

# Select Source Element
try:
    with forms.WarningBar(title="Select one element to use as source"):
        picked = uidoc.Selection.PickObject(ObjectType.Element, "Pick an element")
        sourceElement = doc.GetElement(picked.ElementId)
except:
    showForegroundAlert("Selection cancelled.", "Cancelled")
    sys.exit()

# Collect Parameters
parameters = [p for p in sourceElement.Parameters if p.Definition]

sourceParameterNames = sorted(set(p.Definition.Name for p in parameters))
targetParameterNames = sorted(set(p.Definition.Name for p in parameters if p.StorageType == DB.StorageType.String and not p.IsReadOnly))

scopeOptions = ["Only this instance", "All instances in Active View", "All instances in Entire Project"]

# UI Form
class ParamTransferForm(Window):
    def __init__(self):
        self.Title = "Transfer Parameter"
        self.Width, self.Height = 360, 245
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True

        panel = StackPanel()
        panel.Margin = Thickness(10)

        panel.Children.Add(Label(Content="Select parameter value to transfer:"))
        self.sourceCombo = ComboBox(ItemsSource=sourceParameterNames)
        panel.Children.Add(self.sourceCombo)

        panel.Children.Add(Label(Content="Select parameter to transfer value to:"))
        self.targetCombo = ComboBox(ItemsSource=targetParameterNames)
        panel.Children.Add(self.targetCombo)

        panel.Children.Add(Label(Content="Which instances?"))
        self.scopeCombo = ComboBox(ItemsSource=scopeOptions)
        self.scopeCombo.SelectedIndex = 0
        panel.Children.Add(self.scopeCombo)

        buttonPanel = StackPanel(Orientation=Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Right, Margin=Thickness(0, 10, 0, 0))
        self.cancelButton = Button(Content="Cancel", Width=80, Margin=Thickness(5))
        self.cancelButton.Click += self.cancelClicked
        self.applyButton = Button(Content="Apply", Width=80, Margin=Thickness(5))
        self.applyButton.Click += self.applyClicked
        buttonPanel.Children.Add(self.cancelButton)
        buttonPanel.Children.Add(self.applyButton)

        panel.Children.Add(buttonPanel)
        self.Content = panel

        self.result = None

    def applyClicked(self, sender, args):
        if not self.sourceCombo.SelectedItem or not self.targetCombo.SelectedItem:
            showForegroundAlert("Please select source and target parameters.")
            return
        self.result = (self.sourceCombo.SelectedItem, self.targetCombo.SelectedItem, self.scopeCombo.SelectedItem)
        self.DialogResult = True

    def cancelClicked(self, sender, args):
        self.result = None
        self.DialogResult = False

form = ParamTransferForm()
if not form.ShowDialog():
    showForegroundAlert("Operation cancelled.", "Cancelled")
    sys.exit()

sourceParameter, targetParameter, scopeChoice = form.result

# Target Elements
def getTargetElements(scopeChoice, referenceElement):
    categoryId = referenceElement.Category.Id
    if scopeChoice == "Only this instance": return [referenceElement]
    elif scopeChoice == "All instances in Active View":
        return [e for e in DB.FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType() if e.Category and e.Category.Id == categoryId]
    elif scopeChoice == "All instances in Entire Project":
        return [e for e in DB.FilteredElementCollector(doc).WhereElementIsNotElementType() if e.Category and e.Category.Id == categoryId]
    return [referenceElement]

targetElements = getTargetElements(scopeChoice, sourceElement)

def getParameter(elem, name):
    for p in elem.Parameters:
        if p.Definition and p.Definition.Name == name: return p
    return None

# Transfer values
count = 0
t = DB.Transaction(doc, "Transfer Parameter Value")
t.Start()
for elem in targetElements:
    sourceParam = getParameter(elem, sourceParameter)
    targetParam = getParameter(elem, targetParameter)

    if sourceParam and targetParam and not targetParam.IsReadOnly:
        try: value = sourceParam.AsValueString() or sourceParam.AsString()
        except: value = None
        if value is None: value = ""
        targetParam.Set(value)
        count += 1
t.Commit()

showForegroundAlert("Parameter value copied to {} elements.".format(count), title="Done")