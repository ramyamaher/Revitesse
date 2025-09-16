# -*- coding: utf-8 -*-
import clr
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import System
from System import Array
from System.Windows.Forms import (Application, Form, CheckedListBox, DockStyle, Button, FlowLayoutPanel, FlowDirection, FormStartPosition, ComboBox, Label, Padding, DialogResult, Control)
from Autodesk.Revit.DB import (FilteredElementCollector, BuiltInCategory, RevisionCloud, IndependentTag, ViewSheet, OverrideGraphicSettings, BuiltInParameter, Transaction, ElementId)
from Autodesk.Revit.UI import TaskDialog

# Fall back to SystemExit when not available
try: from pyrevit import script as pyScript
except Exception: pyScript = None

def exitScript():
    if pyScript: pyScript.exit()
    else: raise SystemExit

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Form 1: Revision Selection
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
            try: dateString = r.RevisionDate if isinstance(r.RevisionDate, str) else r.RevisionDate.ToShortDateString()
            except Exception: dateString = ""
            try: desc = r.Description or ""
            except Exception: desc = ""
            items.append("{}{}".format(desc, (" - " + dateString) if dateString else ""))

        self.checkedListBox.Items.AddRange(Array[object](items))

        for i in range(self.checkedListBox.Items.Count): self.checkedListBox.SetItemChecked(i, True)

        # Buttons row
        buttonPanel = FlowLayoutPanel()
        buttonPanel.Height = 50
        buttonPanel.Width = 360
        buttonPanel.FlowDirection = FlowDirection.LeftToRight
        buttonPanel.WrapContents = False
        buttonPanel.AutoSize = False
        buttonPanel.Left = (self.ClientSize.Width - buttonPanel.Width) // 2
        buttonPanel.Top = self.ClientSize.Height - buttonPanel.Height

        self.buttonSelectAll = Button(Text="Select All", Width=85)
        self.buttonSelectNone = Button(Text="Select None", Width=110)
        self.buttonCancel = Button(Text="Cancel", Width=70)
        self.buttonApply = Button(Text="OK", Width=70)

        self.buttonSelectAll.Click += self.selectAll
        self.buttonSelectNone.Click += self.selectNone
        self.buttonCancel.Click += self.cancel
        self.buttonApply.Click += self.apply

        buttonPanel.Controls.AddRange(Array[Control]([self.buttonSelectAll, self.buttonSelectNone, self.buttonCancel, self.buttonApply]))

        self.Controls.AddRange(Array[Control]([self.checkedListBox, buttonPanel]))

        self.revisions = revisions
        self.selectedRevisions = []  # will be list of Revision elements

    def selectAll(self, sender, args):
        for i in range(self.checkedListBox.Items.Count): self.checkedListBox.SetItemChecked(i, True)

    def selectNone(self, sender, args):
        for i in range(self.checkedListBox.Items.Count): self.checkedListBox.SetItemChecked(i, False)

    def cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def apply(self, sender, args):
        self.selectedRevisions = [
            self.revisions[i]
            for i in range(self.checkedListBox.Items.Count)
            if self.checkedListBox.GetItemChecked(i)
        ]
        self.DialogResult = DialogResult.OK
        self.Close()

# Form 2: Cloud Override Reset
class CloudOverrideResetForm(Form):
    def __init__(self):
        Form.__init__(self)  # explicit base ctor
        self.Text = "Cloud Override Reset"
        self.Width = 400
        self.Height = 150
        self.StartPosition = FormStartPosition.CenterScreen

        self.label = Label(Text="Revision clouds to reset:", Left=10, Top=20, Width=150)

        self.cmbCloudOverride = ComboBox()
        self.cmbCloudOverride.Left = 160
        self.cmbCloudOverride.Top = 20
        self.cmbCloudOverride.Width = 210

        # Use user's provided snippet for items & default
        self.cmbCloudOverride.Items.AddRange(Array[object](["All revision clouds", "Revision clouds placed on sheets", "Revision clouds placed on views"]))
        self.cmbCloudOverride.SelectedIndex = 0

        panelButtons = FlowLayoutPanel(Dock=DockStyle.Bottom, Height=52, FlowDirection=FlowDirection.RightToLeft)
        panelButtons.Padding = Padding(0, 0, 10, 0)

        self.buttonOK = Button(Text="OK", Width=80)
        self.buttonCancel = Button(Text="Cancel", Width=80)
        self.buttonApply = Button(Text="Apply", Width=80)

        self.buttonOK.Click += self.onOk
        self.buttonCancel.Click += self.onCancel
        self.buttonApply.Click += self.onApply

        panelButtons.Controls.AddRange(Array[Control]([self.buttonApply, self.buttonCancel, self.buttonOK]))
        self.Controls.AddRange(Array[Control]([self.label, self.cmbCloudOverride, panelButtons]))

        # Flags read by main loop
        self.applyClicked = False
        self.okClicked = False

    def onOk(self, sender, args):
        self.applyClicked = True
        self.okClicked = True
        self.DialogResult = DialogResult.OK
        self.Close()

    def onCancel(self, sender, args):
        self.applyClicked = False
        self.okClicked = False
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def onApply(self, sender, args):
        self.applyClicked = True
        self.okClicked = False
        self.DialogResult = DialogResult.Retry  # arbitrary non-cancel result
        self.Close()

# Override reset
def elementRevisionIdForCloud(cloud):
    # Try property first
    try:
        rid = getattr(cloud, "RevisionId", None)
        if isinstance(rid, ElementId) and rid.IntegerValue != -1: return rid
    except Exception: pass
    # Fallback to parameter
    try:
        p = cloud.get_Parameter(BuiltInParameter.REVISION_CLOUD_REVISION)
        if p: return p.AsElementId()
    except Exception: pass
    return ElementId.InvalidElementId

def resetOverrides(selectedRevisions, cloudFilter):
    # Collect targets
    clouds = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RevisionClouds).WhereElementIsNotElementType().ToElements()
    cloudTags = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RevisionCloudTags).WhereElementIsNotElementType().ToElements()

    selectedRevisionIds = set([r.Id for r in selectedRevisions])

    def shouldOverride(view):
        isSheet = isinstance(view, ViewSheet)
        if cloudFilter == "All revision clouds": return True
        elif cloudFilter == "Revision clouds placed on sheets" and isSheet: return True
        elif cloudFilter == "Revision clouds placed on views" and not isSheet: return True
        return False

    t = Transaction(doc, "Reset Revision Cloud Overrides")
    t.Start()

    resetCountClouds = 0
    resetCountTags = 0
    ogsDefault = OverrideGraphicSettings()

    # Clouds
    for cloud in clouds:
        try:
            if not isinstance(cloud, RevisionCloud): continue
            revId = elementRevisionIdForCloud(cloud)
            if revId not in selectedRevisionIds: continue
            view = doc.GetElement(cloud.OwnerViewId)
            if view is None or not shouldOverride(view): continue
            view.SetElementOverrides(cloud.Id, ogsDefault)  # reset to default
            resetCountClouds += 1
        except Exception: pass

    # Tags
    for tag in cloudTags:
        try:
            if not isinstance(tag, IndependentTag): continue
            # Get the first tagged local element (revision cloud)
            try: tagged_ids = list(tag.GetTaggedLocalElementIds())
            except Exception:
                try: tagged_ids = list(tag.GetTaggedElementIds())
                except Exception: tagged_ids = []
            if not tagged_ids: continue
            tagged_el = doc.GetElement(tagged_ids[0])
            if not isinstance(tagged_el, RevisionCloud): continue
            revId = elementRevisionIdForCloud(tagged_el)
            if revId not in selectedRevisionIds: continue
            view = doc.GetElement(tagged_el.OwnerViewId)
            if view is None or not shouldOverride(view): continue
            view.SetElementOverrides(tag.Id, ogsDefault)  # reset to default
            resetCountTags += 1
        except Exception: pass

    t.Commit()
    return resetCountClouds, resetCountTags

# Get all revisions in document
revisions = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Revisions).WhereElementIsNotElementType().ToElements())

if not revisions:
    TaskDialog.Show("Cloud Override Reset", "No Revisions found in this document.")
    exitScript()

# Step 1: let user select revisions
revForm = RevisionSelectionForm(revisions)
if revForm.ShowDialog() != DialogResult.OK or not revForm.selectedRevisions: exitScript()

selectedRevisions = revForm.selectedRevisions

# Step 2: choose where to reset & allow Apply/OK behavior
while True:
    resetForm = CloudOverrideResetForm() 
    resetForm.ShowDialog() # Use ShowDialog() instead of Application.Run() for modal dialog within Revit
    if not resetForm.applyClicked: break # Cancel or closed without Apply/OK

    cloudFilter = resetForm.cmbCloudOverride.SelectedItem
    cloudsReset, tagsReset = resetOverrides(selectedRevisions, cloudFilter)

    # Simple feedback
    TaskDialog.Show("Cloud Override Reset", "Reset overview:\n\n• Revision Clouds: {}\n• Cloud Tags: {}".format(cloudsReset, tagsReset))

    if resetForm.okClicked: break # If OK was used, stop; if Apply, loop to allow another apply or OK