import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Windows.Forms import *
from System.Drawing import *
from System.Collections.Generic import List
from System import Array, Object

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class TextSelectorForm(Form):
    def __init__(self):
        self.Text = "Select Text"
        self.Size = Size(380, 220)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False

        # Create controls and add them as a .NET array
        controlsList = [
            self.createLabel("Scope", 20, 20),
            self.createCombo(["Active View", "Entire Model"], 150, 18, 200),
            self.createLabel("Filter Type:", 20, 60),
            self.createRadio("Exact Text Value", 20, 85, True),
            self.createRadio("OR Text Contains", 20, 115, False),
        ]
        self.Controls.AddRange(Array[Control](controlsList))

        # Set up controls references
        self.spaceCombo = self.Controls[1]
        self.exactRadio = self.Controls[3]
        self.containsRadio = self.Controls[4]

        # Exact text dropdown
        self.exactCombo = ComboBox()
        self.exactCombo.Location = Point(150, 85)
        self.exactCombo.Size = Size(200, 25)
        self.exactCombo.DropDownStyle = ComboBoxStyle.DropDownList
        self.Controls.Add(self.exactCombo)

        # Contains textbox
        self.containsText = TextBox()
        self.containsText.Location = Point(150, 115)
        self.containsText.Size = Size(200, 110)
        self.containsText.Text = "<None>"
        self.containsText.Enabled = False
        self.containsText.Enter += lambda s, e: self.clearPlaceholder()
        self.containsText.Leave += lambda s, e: self.restoreNone()
        self.Controls.Add(self.containsText)

        # Buttons
        buttonsList = [self.createButton("Cancel", 200, 145, self.cancelButton),self.createButton("Select", 280, 145, self.selectButton)]
        self.Controls.AddRange(Array[Control](buttonsList))

        # Events
        self.exactRadio.CheckedChanged += lambda s, e: self.toggleControls()
        self.containsRadio.CheckedChanged += lambda s, e: self.toggleControls()

        self.populateTextValues()

    def createLabel(self, text, x, y):
        label = Label()
        label.Text = text
        label.Location = Point(x, y)
        label.Size = Size(80, 20)
        return label

    def createCombo(self, items, x, y, width):
        combo = ComboBox()
        combo.Location = Point(x, y)
        combo.Size = Size(width, 25)
        combo.DropDownStyle = ComboBoxStyle.DropDownList
        netArray = Array[Object](items)
        combo.Items.AddRange(netArray)
        combo.SelectedIndex = 0
        return combo

    def createRadio(self, text, x, y, checked):
        radio = RadioButton()
        radio.Text = text
        radio.Location = Point(x, y)
        radio.Size = Size(120, 20)
        radio.Checked = checked
        return radio

    def createButton(self, text, x, y, handler):
        button = Button()
        button.Text = text
        button.Location = Point(x, y)
        button.Size = Size(70, 22)
        button.Click += handler
        return button

    def populateTextValues(self):
        collector = FilteredElementCollector(doc)
        textNotes = collector.OfCategory(BuiltInCategory.OST_TextNotes).WhereElementIsNotElementType().ToElements()
        # Get unique, non-system text values
        textValues = set()
        for note in textNotes:
            if hasattr(note, 'Text') and note.Text:
                text = note.Text.strip()
                if text and not self.isSystemText(text): textValues.add(text)

        sortedValues = sorted(textValues)
        netArray = Array[Object](sortedValues)
        self.exactCombo.Items.AddRange(netArray)
        if sortedValues: self.exactCombo.SelectedIndex = 0

    def isSystemText(self, text):
        return (len(text) > 150 or
                '123456789' in text.replace(' ', '') or
                sum(1 for w in text.split() if len(w) == 1 and w.isalpha()) > 10)

    def toggleControls(self):
        self.exactCombo.Enabled = self.exactRadio.Checked
        self.containsText.Enabled = self.containsRadio.Checked

    def clearPlaceholder(self):
        if self.containsText.Text == "<None>": self.containsText.Text = ""

    def restoreNone(self):
        if not self.containsText.Text.strip(): self.containsText.Text = "<None>"

    def cancelButton(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def selectButton(self, sender, args):
        try:
            # Get text notes based on scope
            if self.spaceCombo.SelectedItem == "Active View": collector = FilteredElementCollector(doc, doc.ActiveView.Id)
            else: collector = FilteredElementCollector(doc)
            textNotes = collector.OfCategory(BuiltInCategory.OST_TextNotes).WhereElementIsNotElementType().ToElements()
            # Filter notes
            if self.exactRadio.Checked and self.exactCombo.SelectedItem:
                target = str(self.exactCombo.SelectedItem)
                matching = [n for n in textNotes if n.Text and n.Text.strip() == target]
            else:
                contains = self.containsText.Text
                if contains and contains != "<None>": matching = [n for n in textNotes if n.Text and contains.lower() in n.Text.lower()]
                else: matching = []

            if matching:
                elementIds = List[ElementId]([n.Id for n in matching])
                uidoc.Selection.SetElementIds(elementIds)
                MessageBox.Show("Selected {} text note(s).".format(len(matching)), "Selection Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)

                self.DialogResult = DialogResult.OK
                self.Close()
            else: MessageBox.Show("No text notes found. Please adjust your criteria.", "No Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
        except Exception as e: MessageBox.Show("Error: {}".format(str(e)), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

# Main execution
if __name__ == "__main__":
    form = TextSelectorForm()
    form.ShowDialog()