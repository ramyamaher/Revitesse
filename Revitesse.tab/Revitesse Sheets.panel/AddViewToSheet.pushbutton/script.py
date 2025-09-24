from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import (FilteredElementCollector, ViewSheet, Transaction, BuiltInCategory, ElementId, View, Viewport, ViewType, XYZ)
from System.Windows.Forms import (Form, Label, TextBox, Button, ComboBox, FormStartPosition, FormBorderStyle, ComboBoxStyle)
from System.Drawing import Point, Size, Font, FontStyle
import System

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

activeView = doc.ActiveView
if not isinstance(activeView, ViewSheet): forms.alert("A sheet has to be open for running this script. Please open a sheet first.", exitscript=True)

existingSheetNumbers = [s.SheetNumber for s in FilteredElementCollector(doc).OfClass(ViewSheet)]

def getTitleBlock(sheet):
    titleblocks = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
    return titleblocks[0] if titleblocks else None

def getPlaceableViews():
    allViews = FilteredElementCollector(doc).OfClass(View).ToElements()
    sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet).ToElements())
    allowedViewTypes = set([ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.DraftingView, ViewType.Elevation, ViewType.ThreeD])
    placedViewIds = set()
    for sheet in sheets:
        vports = FilteredElementCollector(doc, sheet.Id).OfClass(Viewport).ToElements()
        for vp in vports: placedViewIds.add(vp.ViewId)
    placeableViews = []
    for v in allViews:
        if v.IsTemplate: continue
        if v.Id in placedViewIds: continue
        if v.ViewType in allowedViewTypes: placeableViews.append(v)
    return sorted(placeableViews, key=lambda v: v.Name)

placableViews = getPlaceableViews()
viewNameMap = {v.Name: v for v in placableViews}
viewNames = list(viewNameMap.keys())
viewNames.sort()

def copyParameters(source, target):
    try:
        if source.StorageType == DB.StorageType.String: target.Set(source.AsString())
        elif source.StorageType == DB.StorageType.Integer: target.Set(source.AsInteger())
        elif source.StorageType == DB.StorageType.Double: target.Set(source.AsDouble())
        elif source.StorageType == DB.StorageType.ElementId: target.Set(source.AsElementId())
    except: pass

class SheetDuplicationForm(Form):
    def __init__(self):
        self.Text = "Duplicate Sheet with Views"
        self.Size = Size(580, 370)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = self.MinimizeBox = False

        self._labelFont = Font("Segoe UI", 9, FontStyle.Bold)

        # headers
        for text, x in [("New sheet number(s)", 20), ("New sheet name(s)", 200), ("View name(s)", 380)]: self.Controls.Add(self._label(text, x, 10))
        # storage
        self.sheetNumberBoxes, self.sheetNameBoxes, self.viewDropdowns = [], [], []
        self.previouslySelectedViews = [None] * 10
        self.selectedViewsSet = set()
        self.currentViewNames = list(viewNames)
        # 10 rows
        for i in range(10):
            y = 40 + i * 25
            nb = self._textbox(20, y)
            nm = self._textbox(200, y)
            cb = self._dropdown(380, y, self.currentViewNames, self.makeDropdownHandler(i))

            for c in (nb, nm, cb): self.Controls.Add(c)

            self.sheetNumberBoxes.append(nb)
            self.sheetNameBoxes.append(nm)
            self.viewDropdowns.append(cb)
        # buttons
        self.Controls.Add(self._button("Close", 210, 300, self.onClose))
        self.Controls.Add(self._button("Apply", 290, 300, self.onApply))

    def _label(self, text, x, y):
        lbl = Label()
        lbl.Text, lbl.Location = text, Point(x, y)
        lbl.AutoSize, lbl.Font = True, self._labelFont
        return lbl

    def _textbox(self, x, y, w=170, h=20):
        tb = TextBox()
        tb.Size, tb.Location = Size(w, h), Point(x, y)
        return tb

    def _dropdown(self, x, y, items, handler=None, w=170, h=20):
        cb = ComboBox()
        cb.Size, cb.Location = Size(w, h), Point(x, y)
        cb.DropDownStyle = ComboBoxStyle.DropDownList
        cb.Items.AddRange(System.Array[object](["None"] + items))
        cb.SelectedIndex = 0  # default to None
        if handler: cb.SelectionChangeCommitted += handler
        return cb

    def _button(self, text, x, y, handler, w=70, h=22):
        btn = Button()
        btn.Text, btn.Size, btn.Location = text, Size(w, h), Point(x, y)
        btn.Click += handler
        return btn

    # Dropdown logic
    def makeDropdownHandler(self, idx):
        def handler(sender, args):
            newView = sender.Text
            if newView == "None": newView = None

            oldView = self.previouslySelectedViews[idx]
            if oldView == newView: return

            if oldView:
                self.addViewToAll(oldView)
                self.selectedViewsSet.discard(oldView)

            if newView:
                if newView in self.selectedViewsSet:
                    forms.alert("This view has already been selected. Please choose another view.", title="Duplicate View")
                    sender.SelectedIndex = 0
                    return
                self.removeViewFromAll(newView)
                self.selectedViewsSet.add(newView)

            self.previouslySelectedViews[idx] = newView

            for i, dropdown in enumerate(self.viewDropdowns):
                if i != idx:
                    currentSelection = dropdown.Text
                    dropdown.Items.Clear()
                    dropdown.Items.AddRange(System.Array[object](["None"] + self.getAvailableViewsForDropdown(i)))
                    if currentSelection in dropdown.Items: dropdown.SelectedItem = currentSelection
                    else: dropdown.SelectedIndex = 0
        return handler

    def getAvailableViewsForDropdown(self, dropdownIdx):
        selectedExceptCurrent = set(self.previouslySelectedViews)
        selectedExceptCurrent.discard(self.previouslySelectedViews[dropdownIdx])
        return sorted([v for v in viewNames if v not in selectedExceptCurrent])

    def removeViewFromAll(self, viewName):
        if viewName and viewName not in self.currentViewNames: self.currentViewNames.append(viewName)

    def addViewToAll(self, viewName):
        if viewName and viewName not in self.currentViewNames:
            self.currentViewNames.append(viewName)
            self.currentViewNames.sort()

    # Apply / duplicate sheets
    def onApply(self, sender, args):
        inputs = []
        errorRows = []

        # Validate inputs first
        for idx in range(10):
            number = self.sheetNumberBoxes[idx].Text.strip()
            name = self.sheetNameBoxes[idx].Text.strip()
            viewName = self.viewDropdowns[idx].Text.strip()
            if viewName == "None": viewName = None

            # Skip completely empty rows
            if not number and not name and not viewName: continue
            # If any of the required fields is missing, mark as error
            if (not number or not name or not viewName):
                errorRows.append(idx + 1)
                continue
            # Check for duplicate sheet number
            if number in existingSheetNumbers:
                forms.alert("Sheet number '" + number + "' already exists. Please choose another.", title="Duplicate Error")
                return
            # Check if view exists
            if viewName not in viewNameMap:
                forms.alert("Invalid view selected for row " + str(idx + 1) + ": " + viewName)
                return

            inputs.append((number, name, viewNameMap[viewName]))

        if errorRows:
            forms.alert("Please fill valid sheet number, sheet name, and view for row(s):\n" + ", ".join([str(r) for r in errorRows]), title="Incomplete Entries")
            return

        if not inputs:
            forms.alert("No valid inputs provided.")
            return

        # Start transaction only after validation
        titleBlock = getTitleBlock(activeView)
        failedViews = []

        t = Transaction(doc, "Duplicate Sheets with Views")
        t.Start()
        for number, name, view in inputs:
            newSheet = ViewSheet.Create(doc, titleBlock.GetTypeId() if titleBlock else ElementId.InvalidElementId)
            newSheet.SheetNumber = number
            newSheet.Name = name

            for param in activeView.Parameters:
                if not param.IsReadOnly:
                    pname = param.Definition.Name
                    if pname not in ["Sheet Number", "Sheet Name"]:
                        targetParam = newSheet.LookupParameter(pname)
                        if targetParam and not targetParam.IsReadOnly: copyParameters(param, targetParam)
            try:
                outline = newSheet.Outline
                center = XYZ((outline.Max.U + outline.Min.U) / 2, (outline.Max.V + outline.Min.V) / 2, 0)
                Viewport.Create(doc, newSheet.Id, view.Id, center)
            except Exception: failedViews.append(view.Name)
        t.Commit()

        msg = "Sheet(s) duplicated successfully."
        if failedViews: msg += "\nFailed to place views: " + ", ".join(failedViews)

        forms.alert(msg)
        if not failedViews: self.Close()

    def onClose(self, sender, args):
        self.Close()

# launch the form
form = SheetDuplicationForm()
form.ShowDialog()