# -*- coding: utf-8 -*-
import clr, os
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from System.Windows.Forms import *
from System.Drawing import *
from System import Array
from pyrevit import revit, DB
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

paramName = "Revitesse Clouds"
paramGroupName = "Project Parameters"

# Step 1: Revision Selection Form
class RevisionSelectionForm(Form):
    def __init__(self, revisions):
        Form.__init__(self)
        self.Text = "Select Revisions"
        self.Width = 400
        self.Height = 500
        self.StartPosition = FormStartPosition.CenterScreen

        self.checkedListBox = CheckedListBox(Dock=DockStyle.Top, Height=400, CheckOnClick=True)
        items = []
        for r in revisions:
            dateString = r.RevisionDate.ToShortDateString() if hasattr(r.RevisionDate, "ToShortDateString") else str(r.RevisionDate)
            items.append(r.Description + " - " + dateString)
        self.checkedListBox.Items.AddRange(Array[object](items))
        for i in range(self.checkedListBox.Items.Count): self.checkedListBox.SetItemChecked(i, True)

        buttonPanel = FlowLayoutPanel()
        buttonPanel.Height = 50
        buttonPanel.Width = 360
        buttonPanel.FlowDirection = FlowDirection.LeftToRight
        buttonPanel.Left = (self.ClientSize.Width - buttonPanel.Width) // 2
        buttonPanel.Top = self.ClientSize.Height - buttonPanel.Height

        self.buttonSelectAll = Button(Text="Select All", Width=85)
        self.buttonSelectNone = Button(Text="Select None", Width=110)
        self.buttonCancel = Button(Text="Cancel", Width=70)
        self.buttonApply = Button(Text="Apply", Width=70)

        self.buttonSelectAll.Click += self.selectAll
        self.buttonSelectNone.Click += self.selectNone
        self.buttonCancel.Click += self.cancel
        self.buttonApply.Click += self.apply

        buttonPanel.Controls.AddRange(Array[Control]([self.buttonSelectAll, self.buttonSelectNone, self.buttonCancel, self.buttonApply]))
        self.Controls.AddRange(Array[Control]([self.checkedListBox, buttonPanel]))
        self.revisions = revisions
        self.selectedRevisions = []

    def selectAll(self, s, e):
        for i in range(self.checkedListBox.Items.Count): self.checkedListBox.SetItemChecked(i, True)

    def selectNone(self, s, e):
        for i in range(self.checkedListBox.Items.Count): self.checkedListBox.SetItemChecked(i, False)

    def apply(self, s, e):
        self.selectedRevisions = [
            self.revisions[i] for i in range(self.checkedListBox.Items.Count)
            if self.checkedListBox.GetItemChecked(i)
        ]
        self.DialogResult = DialogResult.OK
        self.Close()

    def cancel(self, s, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()

# Step 2: Cloud Filter Form (with separator input)
class CloudExportForm(Form):
    def __init__(self, doc, selectedRevisions):
        Form.__init__(self)
        self.doc = doc
        self.selectedRevisions = selectedRevisions
        self.result = None
        self.separator = " "   # default

        self.Text = "Filter Revision Clouds"
        self.Width = 450
        self.Height = 200
        self.StartPosition = FormStartPosition.CenterScreen

        lblCloudFilter = Label(Text='Cloud Placement:', Left=10, Top=20, Width=150, TextAlign=ContentAlignment.MiddleRight)
        self.comboCloudScope = ComboBox(Left=170, Top=20, Width=250, DropDownStyle=ComboBoxStyle.DropDownList)
        self.comboCloudScope.Items.AddRange(Array[object](["All revision clouds", "Revision clouds placed on sheets", "Revision clouds placed on views"]))
        self.comboCloudScope.SelectedIndex = 0

        # Separator input
        lblSeparator = Label(Text='Mark/Comment Separator:', Left=10, Top=60, Width=150, TextAlign=ContentAlignment.MiddleRight)
        self.txtSeparator = TextBox(Left=170, Top=60, Width=250, Text=self.separator)

        panelButtons = FlowLayoutPanel(Dock=DockStyle.Bottom, Height=52, FlowDirection=FlowDirection.RightToLeft)
        panelButtons.Padding = Padding(0, 10, 10, 0)

        btnCancel = Button(Text="Cancel", Width=85)
        btnApply = Button(Text="Apply", Width=100)

        btnCancel.Click += self.onCancel
        btnApply.Click += self.onApply

        panelButtons.Controls.AddRange(Array[Control]([btnApply, btnCancel]))
        self.Controls.AddRange(Array[Control]([lblCloudFilter, self.comboCloudScope, lblSeparator, self.txtSeparator, panelButtons]))

    def onCancel(self, s, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def onApply(self, s, e):
        self.result = self.comboCloudScope.SelectedItem
        separatorValue = self.txtSeparator.Text
        if not separatorValue.strip(): separatorValue = " "
        self.separator = separatorValue
        self.DialogResult = DialogResult.OK
        self.Close()

    def filterClouds(self):
        cloudFilter = self.result
        def shouldInclude(view, isSheet):
            if cloudFilter == "All revision clouds": return True
            elif cloudFilter == "Revision clouds placed on sheets" and isSheet: return True
            elif cloudFilter == "Revision clouds placed on views" and not isSheet: return True
            return False

        clouds = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_RevisionClouds).WhereElementIsNotElementType().ToElements()
        selectedRevisionIds = [r.Id for r in self.selectedRevisions]
        filtered = []
        for cloud in clouds:
            if not isinstance(cloud, RevisionCloud): continue
            revId = cloud.get_Parameter(BuiltInParameter.REVISION_CLOUD_REVISION).AsElementId()
            if revId not in selectedRevisionIds: continue
            view = self.doc.GetElement(cloud.OwnerViewId)
            isSheet = isinstance(view, ViewSheet)
            if not shouldInclude(view, isSheet): continue
            filtered.append(cloud)
        return filtered

# Step 3: Ensure parameter exists
def ensureRevitesseParameter():
    bindings = doc.ParameterBindings

    def isBound(name):
        it = bindings.ForwardIterator()
        it.Reset()
        while it.MoveNext():
            if it.Key.Name.strip().lower() == name.strip().lower(): return True
        return False

    sharedParameterFilePath = app.SharedParametersFilename
    sharedParameterFile = app.OpenSharedParameterFile()

    if not sharedParameterFilePath or not sharedParameterFile:
        folder = os.path.dirname(doc.PathName) if doc.PathName else os.environ.get("TEMP")
        sharedParameterFilePath = os.path.join(folder, "sharedParameters.txt")
        if not os.path.exists(sharedParameterFilePath):
            with open(sharedParameterFilePath, 'w'): pass
        app.SharedParametersFilename = sharedParameterFilePath
        sharedParameterFile = app.OpenSharedParameterFile()

    group = next((g for g in sharedParameterFile.Groups if g.Name == paramGroupName), None)
    if not group: group = sharedParameterFile.Groups.Create(paramGroupName)

    definition = next((d for d in group.Definitions if d.Name == paramName), None)
    if not definition:
        opt = ExternalDefinitionCreationOptions(paramName, SpecTypeId.String.Text)
        definition = group.Definitions.Create(opt)

    cats = CategorySet()
    for cat in [doc.Settings.Categories.get_Item(BuiltInCategory.OST_Views), doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sheets)]:
        if cat.AllowsBoundParameters: cats.Insert(cat)

    binding = InstanceBinding(cats)
    t = Transaction(doc, "Bind Revitesse Clouds parameter")
    t.Start()
    if not isBound(paramName): bindings.Insert(definition, binding, GroupTypeId.Text)
    else: bindings.ReInsert(definition, binding, GroupTypeId.Text)
    t.Commit()

# Step 4: Main Logic
# 1. Select revisions
all_revisions = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Revisions).WhereElementIsNotElementType())
revForm = RevisionSelectionForm(all_revisions)
if revForm.ShowDialog() != DialogResult.OK: script.exit()
selectedRevisions = revForm.selectedRevisions

# 2. Cloud filter & separator input
cloudForm = CloudExportForm(doc, selectedRevisions)
if cloudForm.ShowDialog() != DialogResult.OK: script.exit()

filteredClouds = cloudForm.filterClouds()
if not filteredClouds:
    MessageBox.Show("No matching revision clouds found.", "Info")
    script.exit()

# 3. Ensure parameter exists
ensureRevitesseParameter()

# 4. Group clouds by parent element and build string
viewCloudMap = {}
for cloud in filteredClouds:
    parent = doc.GetElement(cloud.OwnerViewId)
    if not parent: continue
    mark = cloud.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString() or ""
    comment = cloud.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() or ""
    if mark: entry = "{}{}{}".format(mark, cloudForm.separator, comment).strip()
    else: entry = comment.strip()
    if parent.Id not in viewCloudMap: viewCloudMap[parent.Id] = []
    viewCloudMap[parent.Id].append(entry)

# Sort inside each view/sheet
for pid in viewCloudMap: viewCloudMap[pid] = sorted(viewCloudMap[pid])

# 5. Write parameter values
t = Transaction(doc, "Populate Revitesse Clouds parameter")
t.Start()
for pid, entries in viewCloudMap.items():
    element = doc.GetElement(pid)
    value = "\n".join(entries)
    param = element.LookupParameter(paramName)
    if param and not param.IsReadOnly: param.Set(value)
t.Commit()

MessageBox.Show("Updated '{}' parameter for {} views/sheets.".format(paramName, len(viewCloudMap)), "Success")