from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import FilteredElementCollector, ViewSheet, Transaction, BuiltInCategory, ElementId
from System.Windows.Forms import Application, Form, Label, TextBox, Button, FormStartPosition, FormBorderStyle
from System.Drawing import Point, Size, Font, FontStyle

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# 1. The script only works on an open sheet that it duplicates
activeView = doc.ActiveView
if not isinstance(activeView, ViewSheet): forms.alert("A sheet has to be open for running this script. Please open a sheet first.", exitscript=True)

# 2. Collect existing sheet numbers to prevent duplicates
existingSheetNumbers = [s.SheetNumber for s in FilteredElementCollector(doc).OfClass(ViewSheet)]

# 3. Get title block from the currently open sheet
def getTitleBlock(sheet):
    titleblocks = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
    return titleblocks[0] if titleblocks else None

# 4. Custom form for input
class SheetDuplicationForm(Form):
    def __init__(self):
        self.Text, self.Size = "Duplicate Sheet", Size(400, 370)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = self.MinimizeBox = False

        labelFont = Font("Segoe UI", 9, FontStyle.Bold)

        label1 = Label()
        label1.Text, label1.Location = "New sheet number(s)", Point(20, 10)
        label1.AutoSize, label1.Font = True, labelFont
        self.Controls.Add(label1)

        label2 = Label()
        label2.Text, label2.Location = "New sheet name(s)", Point(200, 10)
        label2.AutoSize, label2.Font = True, labelFont
        self.Controls.Add(label2)

        # Input fields
        self.sheet_number_boxes = []
        self.sheetNameBoxes = []

        for i in range(10):
            textboxNumber = TextBox()
            textboxNumber.Size, textboxNumber.Location = Size(170, 10), Point(20, 40 + i * 25)
            self.Controls.Add(textboxNumber)
            self.sheet_number_boxes.append(textboxNumber)

            textboxName = TextBox()
            textboxName.Size, textboxName.Location = Size(170, 10), Point(200, 40 + i * 25)
            self.Controls.Add(textboxName)
            self.sheetNameBoxes.append(textboxName)

        # Buttons
        applyButton = Button()
        applyButton.Text, applyButton.Location = "Apply", Point(200, 300)
        applyButton.Click += self.onApply
        self.Controls.Add(applyButton)

        closeButton = Button()
        closeButton.Text, closeButton.Location = "Close", Point(115, 300)
        closeButton.Click += self.onClose
        self.Controls.Add(closeButton)

    def onApply(self, sender, args):
        inputs = []
        errorRows = []

        for idx, (numberBox, nameBox) in enumerate(zip(self.sheet_number_boxes, self.sheetNameBoxes)):
            num = numberBox.Text.strip()
            name = nameBox.Text.strip()

            if num and name:
                if num in existingSheetNumbers:
                    forms.alert("Sheet number '{}' already exists.\nPlease choose a different number.".format(num), title="Duplicate Error")
                    return
                inputs.append((num, name))
            elif num or name: errorRows.append(idx + 1)

        if errorRows:
            message = "Please enter both a sheet number and name for row(s):\n{}".format(", ".join(str(r) for r in errorRows))
            forms.alert(message, title="Incomplete Entries")
            return  # Keep the form open until the user corrects the data

        if not inputs:
            forms.alert("No valid inputs provided.")
            return

        currentSheet = activeView
        titleBlock = getTitleBlock(currentSheet)

        t = Transaction(doc, "Duplicate Sheets")
        t.Start()

        for sheet_number, sheetName in inputs:
            newSheet = ViewSheet.Create(doc, titleBlock.GetTypeId() if titleBlock else ElementId.InvalidElementId)
            newSheet.SheetNumber = sheet_number
            newSheet.Name = sheetName

            # 5. Copy parameters (except Sheet Number and Name; they will be taken from the input values)
            for param in currentSheet.Parameters:
                if not param.IsReadOnly:
                    name = param.Definition.Name
                    if name in ["Sheet Number", "Sheet Name"]: continue
                    targetParameter = newSheet.LookupParameter(name)
                    if targetParameter and not targetParameter.IsReadOnly:
                        try:
                            if param.StorageType == DB.StorageType.String: targetParameter.Set(param.AsString())
                            elif param.StorageType == DB.StorageType.Integer: targetParameter.Set(param.AsInteger())
                            elif param.StorageType == DB.StorageType.Double: targetParameter.Set(param.AsDouble())
                            elif param.StorageType == DB.StorageType.ElementId: targetParameter.Set(param.AsElementId())
                        except: pass  # ignore failed param copy
        t.Commit()

        forms.alert("Sheet(s) duplicated successfully.")
        self.Close()  # Close form after user confirms

    def onClose(self, sender, args):
        self.Close()

# Run the form
Application.EnableVisualStyles()
form = SheetDuplicationForm()
form.ShowDialog()