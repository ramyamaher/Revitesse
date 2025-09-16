# -*- coding: utf-8 -*-
import clr, os, csv
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import Autodesk
from Autodesk.Revit.DB import *
from System.Windows.Forms import *
from System.Drawing import *
from System import Array
from pyrevit import DB
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, ElementId
from Autodesk.Revit.UI import TaskDialog
from System.Windows.Forms import SaveFileDialog, DialogResult

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def getParameterString(elem, parameterName):
    p = elem.LookupParameter(parameterName)
    return p.AsString() if p else ""

def getParameterInteger(elem, parameterName):
    p = elem.LookupParameter(parameterName)
    return p.AsInteger() if p else None

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

        buttonPanel.Controls.AddRange(Array[Control]([ self.buttonSelectAll, self.buttonSelectNone, self.buttonCancel, self.buttonApply ]))
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

# Override graphics form
class OverrideGraphicsForm(Form):
    def __init__(self, doc, selectedRevisions):
        Form.__init__(self)
        self.doc = doc
        self.selectedRevisions = selectedRevisions
        self.Text = "Cloud Override Graphics"
        self.Width = 460
        self.Height = 360
        self.StartPosition = FormStartPosition.CenterScreen

        self.chkHalftone = CheckBox(Text="Halftone", Left=310, Top=10)
        self.txtDefiningText = TextBox(Left=190, Top=45, Width=240)
        self.cmbLinePattern = ComboBox(Left=190, Top=110, Width=240, DropDownStyle=ComboBoxStyle.DropDownList)
        self.cmbLineweight = ComboBox(Left=190, Top=180, Width=240, DropDownStyle=ComboBoxStyle.DropDownList)
        self.cmbCloudOverride = ComboBox(Left=190, Top=215, Width=240, DropDownStyle=ComboBoxStyle.DropDownList)
        self.cmbCloudOverride.Items.AddRange(Array[object](["All revision clouds", "Revision clouds placed on sheets", "Revision clouds placed on views"]))
        self.cmbCloudOverride.SelectedIndex = 0

        self.buttonColor = Button(Text="<By Object Style>", Left=210, Top=145, Width=220)
        self.buttonColor.TextAlign = ContentAlignment.MiddleLeft
        self.colorBox = Panel(Left=190, Top=149, Width=16, Height=16, BorderStyle=BorderStyle.FixedSingle, BackColor=Color.Transparent)
        self.buttonColor.Click += self.onColorPick
        self.colorBox.Click += self.onColorPick

        labels = {
            "defining": (10, 45, 'Defining text in "Comments":', ContentAlignment.MiddleRight, False, 170),
            "Projection Lines": (10, 85, ' Projection Lines', ContentAlignment.MiddleLeft, True, 170),
            "linepattern": (10, 110, "Pattern:", ContentAlignment.MiddleRight, False, 170),
            "color": (10, 145, "Color:", ContentAlignment.MiddleRight, False, 170),
            "lineweight": (10, 180, "Weight:", ContentAlignment.MiddleRight, False, 170),
            "cloudoverride": (10, 215, "Revision clouds to override:", ContentAlignment.MiddleRight, False, 170)
        }
        
        for _, (x, y, text, align, bold, w) in labels.items():
            lbl = Label(Text=text, Left=x, Top=y, Width=w)
            lbl.TextAlign = align
            if bold: lbl.Font = Font(lbl.Font.FontFamily, lbl.Font.Size, FontStyle.Bold)
            self.Controls.Add(lbl)

        panelButtons = FlowLayoutPanel(Dock=DockStyle.Bottom, Height=52, FlowDirection=FlowDirection.RightToLeft)
        panelButtons.Padding = Padding(0, 0, 10, 0)
        self.buttonApplyExport = Button(Text="Apply and Export CSV", Width=150)
        self.buttonOK = Button(Text="OK", Width=80)
        self.buttonCancel = Button(Text="Cancel", Width=80)
        self.buttonApply = Button(Text="Apply", Width=80)

        self.buttonApplyExport.Click += self.onApplyExport
        self.buttonOK.Click += self.onOk
        self.buttonCancel.Click += self.onCancel
        self.buttonApply.Click += self.onApply

        panelButtons.Controls.AddRange(Array[Control]([self.buttonApply, self.buttonCancel, self.buttonOK, self.buttonApplyExport]))
        self.Controls.AddRange(Array[Control]([ self.txtDefiningText, self.chkHalftone, self.cmbLinePattern, self.cmbLineweight, self.cmbCloudOverride, self.buttonColor, self.colorBox, panelButtons ]))

        self.selectedColor = None
        self.populateLinePatterns()
        self.populateLineWeights()

    def populateLinePatterns(self):
        self.cmbLinePattern.Items.Add("<By Object Style>")
        patterns = [lp.Name for lp in FilteredElementCollector(self.doc).OfClass(LinePatternElement)]
        self.cmbLinePattern.Items.AddRange(Array[object](patterns))
        self.cmbLinePattern.SelectedIndex = 0

    def populateLineWeights(self):
        self.cmbLineweight.Items.Add("<By Object Style>")
        self.cmbLineweight.Items.AddRange(Array[object]([str(i) for i in range(1, 17)]))
        self.cmbLineweight.SelectedIndex = 0

    def onColorPick(self, s, e):
        cd = ColorDialog()
        cd.FullOpen = True
        if self.selectedColor: cd.Color = self.selectedColor
        result = cd.ShowDialog(self)
        if result == DialogResult.OK:
            self.selectedColor = cd.Color
            self.colorBox.BackColor = self.selectedColor
            text_rgb = "RGB, " + str(self.selectedColor.R) + ", " + str(self.selectedColor.G) + ", " + str(self.selectedColor.B)
            self.buttonColor.Text = text_rgb
        else:
            self.selectedColor = None
            self.colorBox.BackColor = Color.Transparent
            self.buttonColor.Text = "<By Object Style>"
            self.buttonColor.TextAlign = ContentAlignment.MiddleLeft

    def applyOverride(self):
        definingText = self.txtDefiningText.Text.strip()
        halftone = self.chkHalftone.Checked
        linePatternName = self.cmbLinePattern.SelectedItem
        lineweightText = self.cmbLineweight.SelectedItem
        cloudFilter = self.cmbCloudOverride.SelectedItem

        linePatternId = None
        if linePatternName != "<By Object Style>":
            linePatternId = next((lp.Id for lp in FilteredElementCollector(self.doc).OfClass(LinePatternElement) if lp.Name == linePatternName), None)

        lineweight = 0 if lineweightText == "<By Object Style>" else int(lineweightText)

        color = None
        if self.selectedColor: color = Autodesk.Revit.DB.Color(self.selectedColor.R, self.selectedColor.G, self.selectedColor.B)

        clouds = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_RevisionClouds).WhereElementIsNotElementType().ToElements()
        cloudTags = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_RevisionCloudTags).WhereElementIsNotElementType().ToElements()

        selectedRevisionIds = [r.Id for r in self.selectedRevisions]

        t = Transaction(self.doc, "Override Revision Clouds and Tags")
        t.Start()

        def shouldOverride(view, isSheet):
            if cloudFilter == "All revision clouds": return True
            elif cloudFilter == "Revision clouds placed on sheets" and isSheet: return True
            elif cloudFilter == "Revision clouds placed on views" and not isSheet: return True
            return False

        for cloud in clouds:
            if isinstance(cloud, RevisionCloud):
                revId = getattr(cloud, "RevisionId", None) or cloud.get_Parameter(BuiltInParameter.REVISION_CLOUD_REVISION).AsElementId()
                if revId not in selectedRevisionIds: continue

                commentParameter = cloud.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                comments = commentParameter.AsString() or ""
                if definingText and definingText.lower() not in comments.lower(): continue

                view = self.doc.GetElement(cloud.OwnerViewId)
                isSheet = isinstance(view, ViewSheet)
                if not shouldOverride(view, isSheet): continue

                ogs = OverrideGraphicSettings(view.GetElementOverrides(cloud.Id))
                if color: ogs.SetProjectionLineColor(color)
                ogs.SetHalftone(halftone)
                if lineweight: ogs.SetProjectionLineWeight(lineweight)
                if linePatternId: ogs.SetProjectionLinePatternId(linePatternId)
                view.SetElementOverrides(cloud.Id, ogs)

        for tag in cloudTags:
            if isinstance(tag, IndependentTag):
                taggedIds = list(tag.GetTaggedLocalElementIds())
                if not taggedIds: continue
                taggedElement = self.doc.GetElement(taggedIds[0])
                if not isinstance(taggedElement, RevisionCloud): continue

                revId = getattr(taggedElement, "RevisionId", None) or taggedElement.get_Parameter(BuiltInParameter.REVISION_CLOUD_REVISION).AsElementId()
                if revId not in selectedRevisionIds: continue

                commentParameter = taggedElement.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                comments = commentParameter.AsString() or ""
                if definingText and definingText.lower() not in comments.lower(): continue

                view = self.doc.GetElement(taggedElement.OwnerViewId)
                isSheet = isinstance(view, ViewSheet)
                if not shouldOverride(view, isSheet): continue

                ogs = OverrideGraphicSettings(view.GetElementOverrides(tag.Id))
                if color: ogs.SetProjectionLineColor(color)
                ogs.SetHalftone(halftone)
                if lineweight: ogs.SetProjectionLineWeight(lineweight)
                if linePatternId: ogs.SetProjectionLinePatternId(linePatternId)
                view.SetElementOverrides(tag.Id, ogs)
        t.Commit()

    def exportCsv(self):
        # Apply overrides first
        self.applyOverride()
        definingText = self.txtDefiningText.Text.strip()
        cloudFilter = self.cmbCloudOverride.SelectedItem

        # Filtering clouds/views
        def shouldOverride(view, isSheet):
            if cloudFilter == "All revision clouds": return True
            elif cloudFilter == "Revision clouds placed on sheets" and isSheet: return True
            elif cloudFilter == "Revision clouds placed on views" and not isSheet: return True
            return False

        # Collect revision clouds
        clouds = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_RevisionClouds).WhereElementIsNotElementType().ToElements()
        # Get selected revision IDs
        selectedRevisionIds = [r.Id for r in self.selectedRevisions]

        # Filter clouds
        filteredClouds = []
        for cloud in clouds:
            if not isinstance(cloud, RevisionCloud): continue

            revisionParameter = cloud.get_Parameter(BuiltInParameter.REVISION_CLOUD_REVISION)
            revisionId = getattr(cloud, "RevisionId", None) or revisionParameter.AsElementId()
            if revisionId not in selectedRevisionIds: continue

            commentParameter = cloud.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            comments = commentParameter.AsString() or ""
            if definingText and definingText.lower() not in comments.lower(): continue

            view = self.doc.GetElement(cloud.OwnerViewId)
            isSheet = isinstance(view, ViewSheet)
            if not shouldOverride(view, isSheet): continue

            filteredClouds.append(cloud)

        if not filteredClouds:
            MessageBox.Show("No revision clouds found for export.", "Info")
            return

        # Prepare CSV header: Parent Name, Revision Description, Mark, Comment, ElementId, plus parameters
        header = ['Parent Name', 'Revision Description', 'Mark', 'Comment', 'ElementId']
        if filteredClouds: header += [p.Definition.Name for p in filteredClouds[0].Parameters]

        # Ask user for file path
        saveDialog = SaveFileDialog()
        saveDialog.Filter = "CSV files (*.csv)|*.csv"
        if saveDialog.ShowDialog() != DialogResult.OK: return
        filepath = saveDialog.FileName

        # Collect row data
        rows = []
        for cloud in filteredClouds:
            # Revision description
            revisionId = cloud.get_Parameter(BuiltInParameter.REVISION_CLOUD_REVISION).AsElementId()
            revision = self.doc.GetElement(revisionId) if revisionId != -1 else None
            revisionDescription = revision.Description if revision else ""
            # Parent view/sheet name
            parent = self.doc.GetElement(cloud.OwnerViewId)
            parentName = parent.Name if parent else ""
            # Mark and Comment
            markParameter = cloud.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString()
            commentParameter = cloud.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString()
            cloudId = cloud.Id
            # All parameter values
            parameterValues = []
            for p in cloud.Parameters:
                val = p.AsString() if p.AsString() is not None else p.AsValueString()
                parameterValues.append(val)

            rows.append([parentName, revisionDescription, markParameter, commentParameter, cloudId] + parameterValues)

        # Sort rows by Parent Name (you can add more sort keys if needed)
        rows.sort(key=lambda r: r[0])
        # Write CSV
        with open(filepath, 'w') as csvfile:
            writer = csv.writer(csvfile, delimiter='\t')
            writer.writerow(header)
            for row in rows: writer.writerow(row)

        MessageBox.Show("Export Complete.\nNumber of selected revision clouds: {}".format(len(filteredClouds)),"Info")

    def onApplyExport(self, s, e):
        self.applyOverride()
        self.exportCsv()

    def onApply(self, s, e):
        self.applyOverride()

    def onOk(self, s, e):
        self.applyOverride()
        self.DialogResult = DialogResult.OK
        self.Close()

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
    overrideForm = OverrideGraphicsForm(doc, revisionForm.selectedRevisions)
    overrideForm.ShowDialog()

if __name__ == '__main__': main()