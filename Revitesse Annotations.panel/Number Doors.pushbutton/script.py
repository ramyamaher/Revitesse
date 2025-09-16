# -*- coding: utf-8 -*-
import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms
from System.Windows.Forms import *
from System.Drawing import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Backup parameter name & group
backupParamName = "Revitesse Old Marks"
backupParamGroup = "Revitesse"

# Create and bind the backup parameter
def createAndBindBackupParameter():
    app = doc.Application
    bindings = doc.ParameterBindings

    # Check if parameter already bound
    def isParameterBound(parameterName):
        it = bindings.ForwardIterator()
        it.Reset()
        while it.MoveNext():
            if it.Key.Name.strip().lower() == parameterName.strip().lower(): return True
        return False

    sharedParameterFilePath = app.SharedParametersFilename
    sharedParameterFile = app.OpenSharedParameterFile()
    # Check if the shared parameter file exists
    if not sharedParameterFilePath or not sharedParameterFile:
        import os
        revitPath = doc.PathName
        folder = os.path.dirname(revitPath) if revitPath else os.environ.get("TEMP")
        sharedParameterFilePath = os.path.join(folder, "sharedParameters.txt")
        if not os.path.exists(sharedParameterFilePath):
            with open(sharedParameterFilePath, 'w'): pass

        app.SharedParametersFilename = sharedParameterFilePath
        sharedParameterFile = app.OpenSharedParameterFile()
        if not sharedParameterFile:
            forms.alert("Unable to create or open shared parameter file.", "Error")
            return None
    # Check if the parameter and its group exist in the shared parameter file
    group = next((g for g in sharedParameterFile.Groups if g.Name == backupParamGroup), None)
    if not group: group = sharedParameterFile.Groups.Create(backupParamGroup)

    definition = next((d for d in group.Definitions if d.Name == backupParamName), None)
    if not definition:
        opt = ExternalDefinitionCreationOptions(backupParamName, SpecTypeId.String.Text)
        definition = group.Definitions.Create(opt)

    categories = CategorySet()
    for cat in doc.Settings.Categories:
        if cat.AllowsBoundParameters: categories.Insert(cat)
    # Bind the parameter to all categories that allow binding
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

# UI Form
class DoorNumberingForm(Form):
    def __init__(self):
        self.Text = "Door Numbering"
        self.Size = Size(500, 220)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 10
        # Prefix & Suffix
        self.Controls.Add(Label(Text="Prefix:", Location=Point(220, y), Size=Size(40, 25)))
        self.prefixBox = TextBox(Location=Point(260, y), Size=Size(70, 25))
        self.Controls.Add(self.prefixBox)
        self.Controls.Add(Label(Text="Suffix:", Location=Point(350, y), Size=Size(40, 25)))
        self.suffixBox = TextBox(Location=Point(400, y), Size=Size(70, 25))
        self.Controls.Add(self.suffixBox)

        y += 30
        # Room reference
        self.Controls.Add(Label(Text="Room reference:", Location=Point(20, y), Size=Size(150, 25)))
        self.roomCombo = ComboBox(Location=Point(220, y), Size=Size(250, 25))
        self.roomCombo.DropDownStyle = ComboBoxStyle.DropDownList
        self.roomCombo.Items.Add("Room To")
        self.roomCombo.Items.Add("Room From")
        self.roomCombo.SelectedIndex = 0
        self.Controls.Add(self.roomCombo)

        y += 30
        # Separator
        self.Controls.Add(Label(Text="Separator if multiple doors:", Location=Point(20, y), Size=Size(200, 25)))
        self.sepBox = TextBox(Location=Point(220, y), Size=Size(250, 25))
        self.Controls.Add(self.sepBox)

        y += 30
        # Sort mode
        self.Controls.Add(Label(Text="Arrange if same room number:", Location=Point(20, y), Size=Size(200, 25)))
        self.sortCombo = ComboBox(Location=Point(220, y), Size=Size(250, 25))
        self.sortCombo.DropDownStyle = ComboBoxStyle.DropDownList
        self.sortCombo.Items.Add("Alphabetically: capital letters")
        self.sortCombo.Items.Add("Alphabetically: small letters")
        self.sortCombo.Items.Add("Numerically")
        self.sortCombo.SelectedIndex = 0
        self.Controls.Add(self.sortCombo)

        y += 40
        # Buttons
        self.backupButton = Button(Text="Backup Old Marks", Location=Point(10, y), Size=Size(150, 30))
        self.cancelButton = Button(Text="Cancel", Location=Point(320, y), Size=Size(70, 30))
        self.okButton = Button(Text="OK", Location=Point(400, y), Size=Size(70, 30))

        self.okButton.Click += self.okClicked
        self.backupButton.Click += self.backupClicked
        self.cancelButton.Click += self.cancelClicked

        self.Controls.Add(self.okButton)
        self.Controls.Add(self.backupButton)
        self.Controls.Add(self.cancelButton)

    def okClicked(self, sender, args):
        self.DialogResult = DialogResult.OK
        self.Close()

    def cancelClicked(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()
        
    def backupClicked(self, sender, args):
        try:
            backupOldMarks()
            forms.alert("Backup completed successfully.", "Backup")
        except Exception as e: forms.alert("Backup failed:\n{}".format(str(e)), "Error")

# Backup function: copies ALL doors' current Mark parameter value to "Revitesse Old Marks"
def backupOldMarks():
    # Ensure parameter is created and bound
    definition = createAndBindBackupParameter()
    if not definition: raise Exception("Failed to create or bind backup parameter.")

    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType()
    doors = list(collector)
    if not doors: raise Exception("No doors found to backup.")

    t2 = Transaction(doc, "Backup Door Marks")
    t2.Start()
    for door in doors:
        currentMark = door.get_Parameter(BuiltInParameter.DOOR_NUMBER)
        backupParam = door.LookupParameter(backupParamName)
        if currentMark and backupParam: backupParam.Set(currentMark.AsString() or "")
    t2.Commit()

# Get active phase with multiple fallback scenarios
def getActivePhase():
    # Try to get phase from active view first
    try:
        phaseParameter = doc.ActiveView.get_Parameter(BuiltInParameter.VIEW_PHASE)
        if phaseParameter and phaseParameter.AsElementId() != ElementId.InvalidElementId: return doc.GetElement(phaseParameter.AsElementId())
    except: pass
    
    # Fallback: get the last phase from the document's phases
    try:
        phases = FilteredElementCollector(doc).OfClass(Phase).ToElements()
        if phases: return list(phases)[-1]
    except: pass    
    # Final fallback: return None and handle doors without phase
    return None

# Door Numbering
def numberDoors(prefix, suffix, roomReference, separator, sortMode):
    if separator is None or separator.strip() == "": separator = " "

    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType()
    doors = list(collector)
    if not doors: forms.alert("No doors found.", exitscript=True)
    # Get the active phase
    phase = getActivePhase()   
    # Group doors by room number with fallback logic
    doorGroups = {}
    doorsWithoutRooms = []

    for door in doors:
        room = None # Try to get room based on phase if available
        if phase:
            try:
                if roomReference == "Room To": room = door.ToRoom[phase] or door.FromRoom[phase]
                else: room = door.FromRoom[phase] or door.ToRoom[phase]
            except: pass
        
        # Fallback: try to get room without phase (for models without phases)
        if not room:
            try:
                if hasattr(door, 'ToRoom') and hasattr(door, 'FromRoom'):
                    # Try to access rooms directly if the door has these properties
                    if roomReference == "Room To": room = getattr(door, 'ToRoom', None) or getattr(door, 'FromRoom', None)
                    else: room = getattr(door, 'FromRoom', None) or getattr(door, 'ToRoom', None)
            except: pass

        if room:
            try:
                roomNumberParam = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                roomNumber = roomNumberParam.AsString() if roomNumberParam else None
                if roomNumber:
                    if roomNumber not in doorGroups: doorGroups[roomNumber] = []
                    doorGroups[roomNumber].append(door)
                else: doorsWithoutRooms.append(door)
            except: doorsWithoutRooms.append(door)
        else: doorsWithoutRooms.append(door)

    t3 = Transaction(doc, "Number Doors")
    t3.Start()  
    # Number doors grouped by room number
    for roomNumber, doorList in doorGroups.items():
        # Sort doors based on selected mode
        if sortMode == "Alphabetically: capital letters":   doorList.sort(key=lambda d: d.Name.upper())
        elif sortMode == "Alphabetically: small letters":   doorList.sort(key=lambda d: d.Name.lower())
        elif sortMode == "Numerically":   doorList.sort(key=lambda d: d.Id)
        # Generate suffixes based on sort mode
        for idx, door in enumerate(doorList):
            markParameter = door.get_Parameter(BuiltInParameter.DOOR_NUMBER)
            if markParameter and not markParameter.IsReadOnly:
                # Generate suffix based on sort mode and number of doors
                if len(doorList) == 1: numberSuffix = ""
                else:
                    if sortMode == "Numerically": numberSuffix = separator + str(idx + 1).zfill(2)
                    else:
                        letter = chr(ord('A') + idx)
                        if sortMode == "Alphabetically: small letters": letter = letter.lower()
                        numberSuffix = separator + letter              
                roomPart = roomNumber or ""
                markParameter.Set(prefix + roomPart + numberSuffix + suffix)

    # Clear door number for doors without rooms on either side
    for door in doorsWithoutRooms:
        markParameter = door.get_Parameter(BuiltInParameter.DOOR_NUMBER)
        if markParameter and not markParameter.IsReadOnly:  markParameter.Set("")    
    t3.Commit()    
    forms.alert("Door numbering completed.", title="Success")

def main():
    form = DoorNumberingForm()
    if form.ShowDialog() == DialogResult.OK:
        prefix = form.prefixBox.Text or ""
        suffix = form.suffixBox.Text or ""
        roomReference = form.roomCombo.SelectedItem
        separator = form.sepBox.Text
        sortMode = form.sortCombo.SelectedItem
        numberDoors(prefix, suffix, roomReference, separator, sortMode)

if __name__ == "__main__": main()