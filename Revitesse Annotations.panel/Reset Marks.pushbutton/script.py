import clr, os
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from pyrevit import forms, DB
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.ApplicationServices import Application
from System.Windows.Forms import *
from System.Drawing import *
from Autodesk.Revit.UI.Selection import ObjectType

paramName = "Revitesse Old Marks"
paramGroupName = "Revitesse"

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
uiapp = __revit__

bindings = doc.ParameterBindings

def isParameterBound(parameterName):
    # Check if shared parameter is already bound
    it = bindings.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        if it.Key.Name.strip().lower() == parameterName.strip().lower(): return True
    return False

def createAndBindParameter(paramName, paramGroupName):
    # Create a shared parameter and bind to all categories
    sharedParameterFilePath = app.SharedParametersFilename
    sharedParameterFile = app.OpenSharedParameterFile()
    # If there is no shared parameter file, create a new one in the same folder of the Revit file
    if not sharedParameterFilePath or not sharedParameterFile:
        revitPath = doc.PathName
        folder = os.path.dirname(revitPath) if revitPath else os.environ.get("TEMP")
        sharedParameterFilePath = os.path.join(folder, "sharedParameters.txt")
        if not os.path.exists(sharedParameterFilePath):
            with open(sharedParameterFilePath, 'w'): pass

        app.SharedParametersFilename = sharedParameterFilePath
        sharedParameterFile = app.OpenSharedParameterFile()
        if not sharedParameterFile:
            MessageBox.Show("Unable to create or open shared parameter file.", "Error")
            return None

    # Get or create group
    group = next((g for g in sharedParameterFile.Groups if g.Name == paramGroupName), None)
    if not group: group = sharedParameterFile.Groups.Create(paramGroupName)
    # Get or create definition
    definition = next((d for d in group.Definitions if d.Name == paramName), None)
    if not definition:
        opt = ExternalDefinitionCreationOptions(paramName, SpecTypeId.String.Text)
        definition = group.Definitions.Create(opt)
    # Bind to all categories that allow binding
    categories = CategorySet()
    for cat in doc.Settings.Categories:
        if cat.AllowsBoundParameters: categories.Insert(cat)

    binding = InstanceBinding(categories)

    t1 = Transaction(doc, "Bind Parameter")
    t1.Start()
    from Autodesk.Revit.DB import GroupTypeId
    if not isParameterBound(definition.Name):
        success = bindings.Insert(definition, binding, GroupTypeId.IdentityData)
        if not success: bindings.ReInsert(definition, binding, GroupTypeId.IdentityData)
    else: bindings.ReInsert(definition, binding, GroupTypeId.IdentityData)
    t1.Commit()
    return definition

def getTargetElements(scope, sourceElement):
    # Return elements to process based on scope and selected element.
    categoryId = sourceElement.Category.Id
    if scope == "Only this element": return [sourceElement]
    elif scope == "All instances of the same category visible in view":
        return [e for e in FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType() if e.Category and e.Category.Id == categoryId]
    elif scope == "All instances of the same category in entire project":
        return [e for e in FilteredElementCollector(doc).WhereElementIsNotElementType() if e.Category and e.Category.Id == categoryId]
    return [sourceElement]

# UI Form
class ResetMarksForm(Form):
    def __init__(self):
        self.Text = "Reset Marks"
        self.Size = Size(400, 190)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = self.MinimizeBox = False

        def addCtrl(ctrlType, text, x, y, w, h, **kwargs):
            ctrl = ctrlType()
            if text is not None: ctrl.Text = text
            ctrl.Location = Point(x, y)
            ctrl.Size = Size(w, h)
            for k, v in kwargs.items(): setattr(ctrl, k, v)
            self.Controls.Add(ctrl)
            return ctrl

        self.backupCheckbox = addCtrl(CheckBox, "Backup old marks", 20, 20, 200, 25)

        self.scopeLabel = addCtrl(Label, "Scope:", 20, 60, 50, 25)
        self.scopeCombo = addCtrl(ComboBox, None, 80, 60, 295, 25, DropDownStyle=ComboBoxStyle.DropDownList)
        self.scopeCombo.Items.Add("Only this element")
        self.scopeCombo.Items.Add("All instances of the same category visible in view")
        self.scopeCombo.Items.Add("All instances of the same category in entire project")
        self.scopeCombo.SelectedIndex = 0

        self.cancelButton = addCtrl(Button, "Cancel", 220, 110, 75, 30)
        self.cancelButton.Click += self.cancelClicked
        self.okButton = addCtrl(Button, "OK", 300, 110, 75, 30)
        self.okButton.Click += self.okClicked

    def okClicked(self, sender, args):
        self.DialogResult = DialogResult.OK
        self.Close()

    def cancelClicked(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

# Main reset function
def resetMarks(sourceElement, backupOldMarks, scope):
    elements = getTargetElements(scope, sourceElement)

    if not elements:
        MessageBox.Show("No elements found in selected scope.", "Info")
        return

    if backupOldMarks:
        definition = createAndBindParameter(paramName, paramGroupName)
        if not definition:
            MessageBox.Show("Failed to create or bind backup parameter.", "Error")
            return

    t2 = Transaction(doc, "Reset Marks")
    t2.Start()
    for e in elements:
        if backupOldMarks:
            oldMark = e.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            backupParam = e.LookupParameter(paramName)
            if oldMark and backupParam: backupParam.Set(oldMark.AsString() or "")

        markParam = e.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        if markParam and not markParam.IsReadOnly: markParam.Set("")
    t2.Commit()
    MessageBox.Show("Marks reset for {} elements.".format(len(elements)), "Success")

# Entry point
try:
    with forms.WarningBar(title="Select exactly one element"):
        pickedObject = uidoc.Selection.PickObject(ObjectType.Element)
        sourceElement = doc.GetElement(pickedObject.ElementId)
except: forms.alert("No element selected.", exitscript=True)

form = ResetMarksForm()
result = form.ShowDialog()
if result == DialogResult.OK:
    backup = form.backupCheckbox.Checked
    scopeChoice = form.scopeCombo.SelectedItem
    resetMarks(sourceElement, backup, scopeChoice)