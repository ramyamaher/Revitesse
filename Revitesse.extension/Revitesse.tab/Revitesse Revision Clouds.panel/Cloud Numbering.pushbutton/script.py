# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import Autodesk
from Autodesk.Revit.DB import *
from System.Windows.Forms import *
from System.Drawing import *
from System import Array
from pyrevit import revit, DB
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

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
        self.selectedRevisions = [self.revisions[i] for i in range(self.checkedListBox.Items.Count) if self.checkedListBox.GetItemChecked(i)]
        self.DialogResult = DialogResult.OK
        self.Close()

    def cancel(self, s, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()

# Cloud Numbering Form
class CloudNumberingForm(Form):
    def __init__(self, doc, selectedRevisions):
        Form.__init__(self)
        self.doc = doc
        self.selectedRevisions = selectedRevisions

        self.Text = "Number Revision Clouds"
        self.Width = 500
        self.Height = 300
        self.StartPosition = FormStartPosition.CenterScreen

        # Clouds to number
        lblPlacement = Label(Text="Clouds to number:", Left=10, Top=20, Width=200, TextAlign=ContentAlignment.MiddleRight)
        self.cmbPlacement = ComboBox(Left=220, Top=20, Width=250, DropDownStyle=ComboBoxStyle.DropDownList)
        self.cmbPlacement.Items.AddRange(Array[object](["On Sheets", "On Views"]))
        self.cmbPlacement.SelectedIndex = 0
        self.cmbPlacement.SelectedIndexChanged += self.updateParentParameters
        # Parameter to use
        lblParam = Label(Text="Parameter to use for numbering:", Left=10, Top=60, Width=200, TextAlign=ContentAlignment.MiddleRight)
        self.cmbParameter = ComboBox(Left=220, Top=60, Width=250, DropDownStyle=ComboBoxStyle.DropDownList)
        # From / To characters
        lblFrom = Label(Text="From character:", Left=10, Top=100, Width=200, TextAlign=ContentAlignment.MiddleRight)
        self.txtFrom = TextBox(Left=220, Top=100, Width=50)
        lblTo = Label(Text="To character:", Left=340, Top=100, Width=80, TextAlign=ContentAlignment.MiddleRight)
        self.txtTo = TextBox(Left=420, Top=100, Width=50)
        # Separator
        lblSep = Label(Text="Separator:", Left=10, Top=140, Width=200, TextAlign=ContentAlignment.MiddleRight)
        self.txtSep = TextBox(Left=220, Top=140, Width=50)
        self.txtSep.Text = "."
        # Add revision index
        self.chkRevIndex = CheckBox(Text="Add revision index after number", Left=220, Top=180, Width=250)
        # Buttons
        self.btnCancel = Button(Text="Cancel", Left=300, Top=220, Width=80)
        self.btnApply = Button(Text="Apply", Left=390, Top=220, Width=80)
        self.btnApply.Click += self.onApply
        self.btnCancel.Click += self.onCancel

        self.Controls.AddRange(Array[Control]([
            lblPlacement, self.cmbPlacement, lblParam, self.cmbParameter,
            lblFrom, self.txtFrom, lblTo, self.txtTo,
            lblSep, self.txtSep, self.chkRevIndex,
            self.btnApply, self.btnCancel
        ]))

        self.updateParentParameters(None, None)

    def updateParentParameters(self, sender, event):
        self.cmbParameter.Items.Clear()
        cloudsOnSheets = self.cmbPlacement.SelectedItem == "On Sheets"
        sampleElement = None
        if cloudsOnSheets:
            collector = FilteredElementCollector(self.doc).OfClass(Autodesk.Revit.DB.ViewSheet) 
            sampleElement = next((e for e in collector), None)
        else:
            collector = FilteredElementCollector(self.doc).OfClass(Autodesk.Revit.DB.View).WhereElementIsNotElementType() #because windows has something else called view
            sampleElement = next((e for e in collector if not e.IsTemplate), None)

        if sampleElement:
            paramNames = list({p.Definition.Name for p in sampleElement.Parameters})
            paramNames.sort(key=lambda x: x.lower())
            self.cmbParameter.Items.AddRange(Array[object](paramNames))
            self.cmbParameter.SelectedIndex = 0

    def onApply(self, s, e):
        from System.Text import StringBuilder
        from System.Windows.Forms import MessageBox, DialogResult
        from Autodesk.Revit.DB import (FilteredElementCollector, BuiltInCategory, RevisionCloud, BuiltInParameter, Transaction, View, ViewSheet)

        # Parse user inputs
        fromChar = int(self.txtFrom.Text) - 1 if self.txtFrom.Text.strip() else None
        toChar = int(self.txtTo.Text) if self.txtTo.Text.strip() else None
        separator = self.txtSep.Text
        addRevIndex = self.chkRevIndex.Checked
        paramName = self.cmbParameter.SelectedItem
        cloudsOnSheets = self.cmbPlacement.SelectedItem == "On Sheets"

        # Collect all revision clouds
        allClouds = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_RevisionClouds).WhereElementIsNotElementType().ToElements()
        # Selected revision IDs
        selectedRevisionIds = [r.Id for r in self.selectedRevisions]

        # Group clouds by parent
        parentCloudDict = {}
        for cloud in allClouds:
            if not isinstance(cloud, RevisionCloud): continue
            revId = cloud.get_Parameter(BuiltInParameter.REVISION_CLOUD_REVISION).AsElementId()
            if revId not in selectedRevisionIds: continue

            parent = self.doc.GetElement(cloud.OwnerViewId)
            if cloudsOnSheets and isinstance(parent, ViewSheet): parentCloudDict.setdefault(parent.Id, []).append(cloud)
            elif not cloudsOnSheets and isinstance(parent, View) and not isinstance(parent, ViewSheet): parentCloudDict.setdefault(parent.Id, []).append(cloud)

        # Start transaction
        t = Transaction(self.doc, "Number Revision Clouds")
        t.Start()
        for parentId, cloudList in parentCloudDict.items(): 
            parent = self.doc.GetElement(parentId)
            # Get parent parameter value
            paramValue = ""
            param = parent.LookupParameter(paramName)
            if param:
                paramValue = param.AsString() or ""
                if fromChar is not None or toChar is not None:
                    start = fromChar if fromChar is not None else 0
                    end = toChar if toChar is not None else len(paramValue)
                    paramValue = paramValue[start:end]
            # Sort clouds for consistent numbering
            cloudList.sort(key=lambda c: c.Id)
            
            for i, cloud in enumerate(cloudList, 1):
                numberParts = [paramValue, "{0:03d}".format(i)]
                if addRevIndex:
                    revId = cloud.get_Parameter(BuiltInParameter.REVISION_CLOUD_REVISION).AsElementId()
                    revIndex = selectedRevisionIds.index(revId) + 1
                    numberParts.append("{0:02d}".format(revIndex))
                fullNumber = separator.join(numberParts)

                # Set built-in Mark parameter on cloud
                markParam = cloud.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
                markParam.Set(fullNumber)
        t.Commit()

        # Show debug output
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

    cloudForm = CloudNumberingForm(doc, revForm.selectedRevisions)
    cloudForm.ShowDialog()

if __name__ == "__main__": main()