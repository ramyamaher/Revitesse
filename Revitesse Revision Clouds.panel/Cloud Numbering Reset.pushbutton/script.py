# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')

from Autodesk.Revit.DB import *
from System.Windows.Forms import *
from System import Array

doc = __revit__.ActiveUIDocument.Document

# Revision Selection Form
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

        buttonPanel = FlowLayoutPanel(Height=50, Width=360, FlowDirection=FlowDirection.LeftToRight, WrapContents=False)
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
        self.selectedRevisions = [self.revisions[i] for i in range(self.checkedListBox.Items.Count) if self.checkedListBox.GetItemChecked(i)]
        self.DialogResult = DialogResult.OK
        self.Close()

    def cancel(self, s, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()

# Reset Marks Form
class ResetCloudMarksForm(Form):
    def __init__(self, doc, selectedRevisions):
        Form.__init__(self)
        self.doc = doc
        self.selectedRevisions = selectedRevisions

        self.Text = "Reset Revision Cloud Marks"
        self.Width = 400
        self.Height = 140
        self.StartPosition = FormStartPosition.CenterScreen

        lblPlacement = Label(Text="Clouds to reset:", Left=10, Top=20, Width=150)
        self.cmbPlacement = ComboBox(Left=170, Top=20, Width=200, DropDownStyle=ComboBoxStyle.DropDownList)
        self.cmbPlacement.Items.AddRange(Array[object](["On Sheets", "On Views"]))
        self.cmbPlacement.SelectedIndex = 0

        self.btnCancel = Button(Text="Cancel", Left=200, Top=60, Width=80)
        self.btnApply = Button(Text="Apply", Left=290, Top=60, Width=80)
        self.btnApply.Click += self.onApply
        self.btnCancel.Click += self.onCancel

        self.Controls.AddRange(Array[Control]([lblPlacement, self.cmbPlacement, self.btnApply, self.btnCancel]))

    def onApply(self, s, e):
        from System.Windows.Forms import MessageBox, DialogResult
        from Autodesk.Revit.DB import (FilteredElementCollector, RevisionCloud, BuiltInParameter, Transaction, View, ViewSheet, ElementId)

        # Scope from UI
        cloudsOnSheets = self.cmbPlacement.SelectedItem == "On Sheets"
        # Selected revisions as a set of ElementId for fast membership checks
        selectedRevisionIds = set(r.Id for r in self.selectedRevisions)
        # Robust way to get a cloud's RevisionId across Revit versions
        def getRevisionId(cloud):
            try:
                rid = cloud.RevisionId
                if rid and rid.IntegerValue != -1: return rid
            except: pass
            p = cloud.get_Parameter(BuiltInParameter.REVISION_CLOUD_REVISION)
            return p.AsElementId() if p else ElementId.InvalidElementId

        # Scope predicate
        def inScope(view):
            if cloudsOnSheets: return isinstance(view, ViewSheet)
            else: return isinstance(view, View) and not isinstance(view, ViewSheet)

        # Collect revision clouds by class and reset marks
        clouds = list(FilteredElementCollector(self.doc).OfClass(RevisionCloud).WhereElementIsNotElementType())

        t = Transaction(self.doc, "Reset Revision Cloud Marks")
        t.Start()
        changed = 0
        for cloud in clouds:
            revId = getRevisionId(cloud)
            if revId not in selectedRevisionIds: continue

            parentView = self.doc.GetElement(cloud.OwnerViewId)
            if not inScope(parentView): continue

            markParam = cloud.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            if markParam and not markParam.IsReadOnly:
                # Only write if there's actually something to clear
                existing = markParam.AsString() or markParam.AsValueString()
                if existing:
                    markParam.Set("")
                    changed += 1
        t.Commit()

        MessageBox.Show("Cleared Mark on {} revision cloud(s).".format(changed), "Reset complete")
        self.DialogResult = DialogResult.OK
        self.Close()

    def onCancel(self, s, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()

def main():
    revisions = FilteredElementCollector(doc).OfClass(Revision).ToElements()
    if not revisions:
        MessageBox.Show("No Revisions found in this project.", "Info")
        return

    revForm = RevisionSelectionForm(revisions)
    if revForm.ShowDialog() != DialogResult.OK: return

    resetForm = ResetCloudMarksForm(doc, revForm.selectedRevisions)
    resetForm.ShowDialog()

if __name__ == "__main__": main()