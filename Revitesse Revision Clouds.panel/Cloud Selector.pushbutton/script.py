# -*- coding: utf-8 -*-
import clr, os, csv
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from System.Windows.Forms import *
from System.Drawing import *
from System import Array
from pyrevit import revit, DB
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def getParameterString(elem, parameterName):
    p = elem.LookupParameter(parameterName)
    return p.AsString() if p else ""

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
            dateString = r.RevisionDate if isinstance(r.RevisionDate, str) else r.RevisionDate.ToShortDateString()
            items.append(r.Description + " - " + dateString)
        self.checkedListBox.Items.AddRange(Array[object](items))

        for i in range(self.checkedListBox.Items.Count): self.checkedListBox.SetItemChecked(i, True)

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

# Cloud Selection & Export Form
class CloudSelectionForm(Form):
    def __init__(self, doc, selectedRevisions):
        Form.__init__(self)
        self.doc = doc
        self.selectedRevisions = selectedRevisions

        self.Text = "Select Revision Clouds"
        self.Width = 450
        self.Height = 180
        self.StartPosition = FormStartPosition.CenterScreen

        # Comment Filter
        lblDefining = Label(Text='Filter text in "Comments":', Left=10, Top=20, Width=150, TextAlign=ContentAlignment.MiddleRight)
        self.txtDefiningText = TextBox(Left=170, Top=20, Width=250)
        # Cloud placement dropdown
        lblCloudFilter = Label(Text='Cloud Placement:', Left=10, Top=60, Width=150, TextAlign=ContentAlignment.MiddleRight)
        self.cmbCloudOverride = ComboBox(Left=170, Top=60, Width=250, DropDownStyle=ComboBoxStyle.DropDownList)
        self.cmbCloudOverride.Items.AddRange(Array[object](["All revision clouds", "Revision clouds placed on sheets", "Revision clouds placed on views"]))
        self.cmbCloudOverride.SelectedIndex = 0
        # Buttons
        panelButtons = FlowLayoutPanel(Dock=DockStyle.Bottom, Height=52, FlowDirection=FlowDirection.RightToLeft)
        panelButtons.Padding = Padding(0, 10, 10, 0)

        self.buttonExport = Button(Text="Select & Export", Width=120)
        self.buttonSelect = Button(Text="Select", Width=85)
        self.buttonCancel = Button(Text="Cancel", Width=85)

        self.buttonExport.Click += self.onExport
        self.buttonSelect.Click += self.onSelect
        self.buttonCancel.Click += self.onCancel

        panelButtons.Controls.AddRange(Array[Control]([self.buttonSelect, self.buttonCancel, self.buttonExport]))
        self.Controls.AddRange(Array[Control]([lblDefining, self.txtDefiningText, lblCloudFilter, self.cmbCloudOverride, panelButtons]))

    def filterClouds(self):
        definingText = self.txtDefiningText.Text.strip()
        cloudFilter = self.cmbCloudOverride.SelectedItem

        def shouldInclude(view, isSheet):
            if cloudFilter == "All revision clouds": return True
            elif cloudFilter == "Revision clouds placed on sheets" and isSheet: return True
            elif cloudFilter == "Revision clouds placed on views" and not isSheet: return True
            return False

        clouds = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_RevisionClouds).WhereElementIsNotElementType().ToElements()

        selectedRevisionIds = [r.Id for r in self.selectedRevisions]
        filteredClouds = []

        for cloud in clouds:
            if not isinstance(cloud, RevisionCloud): continue
            revId = cloud.get_Parameter(BuiltInParameter.REVISION_CLOUD_REVISION).AsElementId()
            if revId not in selectedRevisionIds: continue
            comments = cloud.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() or ""
            if definingText and definingText.lower() not in comments.lower(): continue
            view = self.doc.GetElement(cloud.OwnerViewId)
            isSheet = isinstance(view, ViewSheet)
            if not shouldInclude(view, isSheet): continue
            filteredClouds.append(cloud)

        return filteredClouds

    def onSelect(self, s, e):
        filteredClouds = self.filterClouds()
        elementIds = List[ElementId]([c.Id for c in filteredClouds])
        revit.uidoc.ShowElements(elementIds)
        revit.uidoc.Selection.SetElementIds(elementIds)
        self.Close()
        
    def onExport(self, s, e):
        filteredClouds = self.filterClouds()
        elementIds = List[ElementId]([c.Id for c in filteredClouds])
        revit.uidoc.ShowElements(elementIds)
        revit.uidoc.Selection.SetElementIds(elementIds)

        if not filteredClouds:
            MessageBox.Show("No revision clouds found for export.", "Info")
            return

        saveDialog = SaveFileDialog()
        saveDialog.Filter = "CSV files (*.csv)|*.csv"
        if saveDialog.ShowDialog() != DialogResult.OK: return
        filepath = saveDialog.FileName

        # Header
        header = ['Parent Name', 'Revision Description', 'Mark', 'Comment', 'ElementId']
        if filteredClouds: header += [p.Definition.Name for p in filteredClouds[0].Parameters]

        with open(filepath, 'w') as csvfile:
            writer = csv.writer(csvfile, delimiter='\t')
            writer.writerow(header)
            rows = []

            for cloud in filteredClouds:
                # Revision info
                revisionId = cloud.get_Parameter(BuiltInParameter.REVISION_CLOUD_REVISION).AsElementId()
                revision = doc.GetElement(revisionId) if revisionId != -1 else None
                revisionDescription = revision.Description if revision else ""

                parent = doc.GetElement(cloud.OwnerViewId)
                parentName = parent.Name if parent else ""
                # Mark and comment
                markParameter = cloud.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString()
                commentParameter = cloud.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString()
                cloudId = cloud.Id
                # Parameter values
                parameterValues = []
                for p in cloud.Parameters:
                    val = p.AsString() if p.AsString() is not None else p.AsValueString()
                    parameterValues.append(val)

                rows.append([parentName, revisionDescription, markParameter, commentParameter, cloudId] + parameterValues)

            # Sort rows by parent name
            rows.sort(key=lambda r: r[0])
            for row in rows: writer.writerow(row)

        MessageBox.Show("Export Complete.\nNumber of selected revision clouds: {}".format(len(filteredClouds)), "Info")

    def onCancel(self, s, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()

def main():
    revisions = FilteredElementCollector(doc).OfClass(Revision).ToElements()
    if not revisions:
        MessageBox.Show("No Revisions found in this project.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        return

    revisionForm = RevisionSelectionForm(revisions)
    if revisionForm.ShowDialog() != DialogResult.OK: return

    cloudForm = CloudSelectionForm(doc, revisionForm.selectedRevisions)
    cloudForm.ShowDialog()

if __name__ == '__main__': main()