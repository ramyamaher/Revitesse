# -*- coding: utf-8 -*-
import clr, os
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit import DB
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Windows.Forms import *
from System.Drawing import *
from System import Array

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# Parameter Type Mapping
TYPES = {
    "Angle": DB.SpecTypeId.Angle,
    "Area": DB.SpecTypeId.Area,
    "Boolean (Yes/No)": DB.SpecTypeId.Boolean.YesNo,
    "Currency": DB.SpecTypeId.Currency,
    "Distance": DB.SpecTypeId.Length,
    "Fill Pattern": DB.SpecTypeId.Reference.FillPattern,
    "Image": DB.SpecTypeId.Reference.Image,
    "Integer": DB.SpecTypeId.Int.Integer,
    "Length": DB.SpecTypeId.Length,
    "Mass Density": DB.SpecTypeId.MassDensity,
    "Material": DB.SpecTypeId.Reference.Material,
    "Multiline Text": DB.SpecTypeId.String.MultilineText,
    "Number": DB.SpecTypeId.Number,
    "Rotation Angle": DB.SpecTypeId.Angle,
    "Slope": DB.SpecTypeId.Slope,
    "Speed": DB.SpecTypeId.Speed,
    "Text": DB.SpecTypeId.String.Text,
    "Time": DB.SpecTypeId.Time,
    "URL": DB.SpecTypeId.String.Url,
    "Volume": DB.SpecTypeId.Volume,
}

# GroupTypeId options for "Group Parameter Under" (Revit 2025)
def getAvailableParameterGroups():
    groups = []
    potentialGroups = [
        'Analysis Results', 'Analytical Alignment', 'Analytical Model', 'Constraints', 'Construction', 'Data', 'Dimensions', 'Division Geometry', 'Electrical', 
        'ElectricalCircuiting', 'ElectricalLighting', 'ElectricalLoads', 'ElectricalAnalysis', 'ElectricalEngineering', 
        'EnergyAnalysis', 'FireProtection', 'Forces', 'General', 'Graphics',
        'GreenBuilding', 'IdentityData', 'IFCParameters', 'Layers', 'LifeSafety', 'Materials', 'Mechanical',
        'MechanicalFlow', 'MechanicalLoads', 'Model', 'Moments', 'Other', 'OverallLegend',
        'Phasing', 'Photometrics', 'Plumbing', 'PrimaryEnd', 'RebarSet', 'SegmentsFinishes', 'Set', 'SlabShapeEdit', 'Structural', 'StructuralAnalysis',
        'StructuralSectionDimensions', 'Sub-division', 'Text', 'TitleText', 'ViewToSheetPositioning', 'Visibility', 'Visualization'
    ]
    for grName in potentialGroups:
        try:
            groupId = getattr(DB.GroupTypeId, grName)
            groups.append(groupId)
        except AttributeError: continue
    groups.sort(key=lambda x: LabelUtils.GetLabelForGroup(x))
    return groups

PARAMETER_GROUPS = getAvailableParameterGroups()

# Shared Parameters helpers
def sharedParameterFileExists(application):
    spath = application.SharedParametersFilename
    if not spath or not os.path.exists(spath):
        temporaryPath = os.path.join(os.environ.get('TEMP', os.getcwd()), 'PyRevitSharedParams.txt')
        if not os.path.exists(temporaryPath):
            with open(temporaryPath, 'w') as f: f.write('# Shared Parameter File created by pyRevit\n')
        application.SharedParametersFilename = temporaryPath
        spath = temporaryPath
    return application.OpenSharedParameterFile()

def groupExists(defFile, groupName):
    grp = defFile.Groups.get_Item(groupName)
    if grp is None: grp = defFile.Groups.Create(groupName)
    return grp

def definitionExists(group, parameterName, parameterType):
    for d in group.Definitions:
        if d.Name == parameterName: return d
    options = ExternalDefinitionCreationOptions(parameterName, parameterType)
    return group.Definitions.Create(options)

def bindableCategories(document):
    return sorted([cat for cat in document.Settings.Categories if cat.AllowsBoundParameters], key=lambda cat: cat.Name)

# Form
class SharedParamsForm(Form):
    def __init__(self, document):
        Form.__init__(self)
        self.doc = document
        self.Text = 'Shared Parameter Setup'
        self.Width = 900
        self.Height = 500
        self.StartPosition = FormStartPosition.CenterScreen

        # Main container with left (categories) and right (parameters)
        splitPanel = SplitContainer()
        splitPanel.Dock = DockStyle.Fill
        splitPanel.Orientation = Orientation.Vertical
        splitPanel.SplitterDistance = 43   # left panel width
        self.Controls.Add(splitPanel)

        # LEFT: categories with scroll
        catPanel = Panel()
        catPanel.Dock = DockStyle.Fill
        lblCategories = Label(Text="Categories", AutoSize=False, Width=220, Height=30, Dock=DockStyle.Top, 
                                TextAlign=ContentAlignment.MiddleLeft, Font=Font(Control.DefaultFont, FontStyle.Bold), Padding=Padding(10,0,0,0))

        catPanel.Controls.Add(lblCategories)
        
        scrollCats = Panel()
        scrollCats.Dock = DockStyle.Fill
        scrollCats.AutoScroll = True

        table = TableLayoutPanel()
        table.Dock = DockStyle.None
        table.AutoSize = True
        table.ColumnCount = 1
        table.Padding = Padding(10,30,0,0)

        self.chkCategories = []
        
        for cat in bindableCategories(self.doc):
            cb = CheckBox(Text=cat.Name)
            cb.AutoSize = False
            cb.Width = 220
            cb.Height = 20
            cb.Margin = Padding(2, 0, 2, 0)   # tight vertical spacing
            self.chkCategories.append((cat, cb))
            table.Controls.Add(cb)

        scrollCats.Controls.Add(table)

        # Buttons outside scroll
        btnPanel = FlowLayoutPanel()
        btnPanel.Dock = DockStyle.Bottom
        btnPanel.FlowDirection = FlowDirection.LeftToRight
        btnPanel.Height = 35
        btnPanel.Padding = Padding(5, 5, 5, 5)

        btnAll = Button(Text="Select All", Width=90)
        btnNone = Button(Text="Select None", Width=90)

        btnAll.Click += lambda s, e: [setattr(cb, "Checked", True) for _, cb in self.chkCategories]
        btnNone.Click += lambda s, e: [setattr(cb, "Checked", False) for _, cb in self.chkCategories]

        btnPanel.Controls.AddRange(Array[Control]([btnAll, btnNone]))

        # Add scrollable list + fixed buttons
        catPanel.Controls.Add(scrollCats)
        catPanel.Controls.Add(btnPanel)

        splitPanel.Panel1.Controls.Add(catPanel)

        # RIGHT: parameters
        rightPanel = Panel(Dock=DockStyle.Fill)
        splitPanel.Panel2.Controls.Add(rightPanel)

        # Header row
        headerPanel = Panel(Dock=DockStyle.Top, Height=30, BackColor=SystemColors.Control)
        headerFlow = FlowLayoutPanel(Dock=DockStyle.Fill, FlowDirection=FlowDirection.LeftToRight, WrapContents=False, Padding=Padding(0, 5, 5, 0))
        lbls = [("Parameter Name", 105), ("Parameter Group", 105), ("Data Type", 100), ("Group Parameter Under", 150), ("Binding", 130)]
        for text, w in lbls:
            lbl = Label(Text=text, AutoSize=False, Width=w, TextAlign=ContentAlignment.MiddleLeft, Font=Font(Control.DefaultFont, FontStyle.Bold))
            headerFlow.Controls.Add(lbl)
        headerPanel.Controls.Add(headerFlow)
        rightPanel.Controls.Add(headerPanel)

        # Scroll area for parameter rows
        scrollPanel = Panel()
        scrollPanel.Top = headerPanel.Height
        scrollPanel.Left = 0
        scrollPanel.Dock = DockStyle.Fill
        scrollPanel.AutoScroll = True
        scrollPanel.Padding = Padding(10, 30, 10, 0)

        mainPanel = FlowLayoutPanel()
        mainPanel.FlowDirection = FlowDirection.TopDown
        mainPanel.WrapContents = False
        mainPanel.AutoSize = True
        mainPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink
        mainPanel.Padding = Padding(0, 30, 0, 0)

        self.paramRows = []
        for i in range(10):
            rowPanel = self.createParameterRow(i)
            mainPanel.Controls.Add(rowPanel)

        scrollPanel.Controls.Add(mainPanel)
        rightPanel.Controls.Add(scrollPanel)

        # Bottom buttons
        buttonPanel = FlowLayoutPanel(Dock=DockStyle.Bottom, Height=40, FlowDirection=FlowDirection.RightToLeft, Padding=Padding(0, 0, 10, 0))
        self.btnApply = Button(Text='Apply', Width=90)
        self.btnOK = Button(Text='OK', Width=90)
        self.btnCancel = Button(Text='Cancel', Width=90)
        self.btnApply.Click += self.onApply
        self.btnOK.Click += self.onOK
        self.btnCancel.Click += self.onCancel
        buttonPanel.Controls.AddRange(Array[Control]([self.btnApply, self.btnCancel, self.btnOK]))
        self.Controls.Add(buttonPanel)

    def createParameterRow(self, rowIndex):
        rowPanel = FlowLayoutPanel()
        rowPanel.FlowDirection = FlowDirection.LeftToRight
        rowPanel.WrapContents = False
        rowPanel.AutoSize = True
        rowPanel.Margin = Padding(0, 0, 0, 0)
        # Parameter Name and Group
        txtParam = TextBox(Width=105)
        txtGroup = TextBox(Width=105)
        # Data Type
        cmbParamType = ComboBox(Width=100, DropDownStyle=ComboBoxStyle.DropDownList)
        sortedTypes = sorted(list(TYPES.keys()))
        cmbParamType.Items.AddRange(Array[object](sortedTypes))
        cmbParamType.SelectedIndex = sortedTypes.index("Text")
        #Group Parameter Under
        cmbUnder = ComboBox(Width=150, DropDownStyle=ComboBoxStyle.DropDownList)
        localizedNames = [LabelUtils.GetLabelForGroup(groupId) for groupId in PARAMETER_GROUPS]
        cmbUnder.Items.AddRange(Array[object](localizedNames))
        textGroupName = LabelUtils.GetLabelForGroup(DB.GroupTypeId.IdentityData)
        cmbUnder.SelectedIndex = next((i for i, name in enumerate(localizedNames) if name == textGroupName), 0)
        #Binding
        bindingPanel = FlowLayoutPanel(FlowDirection=FlowDirection.LeftToRight, AutoSize=True, Width=130)
        radioInstance = RadioButton(Text='Instance', Checked=True, AutoSize=True)
        radioType = RadioButton(Text='Type', AutoSize=True)
        bindingPanel.Controls.Add(radioInstance)
        bindingPanel.Controls.Add(radioType)

        rowPanel.Controls.AddRange(Array[Control]([txtParam, txtGroup, cmbParamType, cmbUnder, bindingPanel]))
        self.paramRows.append({'txtParam': txtParam, 'txtGroup': txtGroup, 'cmbParamType': cmbParamType, 'cmbUnder': cmbUnder, 
                                'radioInstance': radioInstance, 'radioType': radioType})
        return rowPanel

    def selectedParameterType(self, rowIndex):
        row = self.paramRows[rowIndex]
        name = row['cmbParamType'].SelectedItem
        return TYPES.get(name, DB.SpecTypeId.String.Text)

    def selectedParameterGroupUnder(self, rowIndex):
        row = self.paramRows[rowIndex]
        idx = row['cmbUnder'].SelectedIndex
        if idx >= 0 and idx < len(PARAMETER_GROUPS): return PARAMETER_GROUPS[idx]
        return DB.GroupTypeId.Data

    def validateAndGetParametersToCreate(self):
        parametersToCreate = []
        for i, row in enumerate(self.paramRows):
            paramName = row['txtParam'].Text.strip()
            groupName = row['txtGroup'].Text.strip()
            if not paramName and not groupName: continue
            if paramName and not groupName:
                MessageBox.Show('Row {}: Please enter a Parameter Group name.'.format(i + 1), 'Validation Error')
                return None
            if not paramName and groupName:
                MessageBox.Show('Row {}: Please enter a Parameter Name.'.format(i + 1), 'Validation Error')
                return None
            parametersToCreate.append({
                'paramName': paramName,
                'groupName': groupName,
                'paramType': self.selectedParameterType(i),
                'paramGroup': self.selectedParameterGroupUnder(i),
                'isTypeBinding': row['radioType'].Checked
            })
        if not parametersToCreate:
            MessageBox.Show('Please enter at least one parameter name and group.', 'Info')
            return None
        return parametersToCreate

    def createAndBind(self):
        parametersToCreate = self.validateAndGetParametersToCreate()
        if parametersToCreate is None: return
        defFile = sharedParameterFileExists(app)
        if defFile is None:
            MessageBox.Show('Could not open or create the shared parameter file.', 'Error')
            return
        t = Transaction(self.doc, 'Bind Shared Parameters')
        t.Start()
        try:
            successCount, errors = 0, []
            for param in parametersToCreate:
                try:
                    catset = app.Create.NewCategorySet()
                    for cat, cb in self.chkCategories:
                        if cb.Checked: catset.Insert(cat)
                    if catset.IsEmpty: continue
                    binding = app.Create.NewTypeBinding(catset) if param['isTypeBinding'] else app.Create.NewInstanceBinding(catset)
                    grp = groupExists(defFile, param['groupName'])
                    definition = definitionExists(grp, param['paramName'], param['paramType'])
                    bmap = self.doc.ParameterBindings
                    if not bmap.Contains(definition): bmap.Insert(definition, binding, param['paramGroup'])
                    else: bmap.ReInsert(definition, binding, param['paramGroup'])
                    successCount += 1
                except Exception as ex: errors.append('Parameter "{}": {}'.format(param['paramName'], str(ex)))
            t.Commit()
            if successCount > 0:
                msg = 'Successfully created {} parameter(s).'.format(successCount)
                if errors: msg += '\n\nErrors:\n' + '\n'.join(errors)
                MessageBox.Show(msg, 'Results')
            elif errors: MessageBox.Show('All parameters failed:\n' + '\n'.join(errors), 'Error')
        except Exception as ex:
            t.RollBack()
            MessageBox.Show('Failed to bind parameters:\n' + str(ex), 'Error')

    def onApply(self, s, e):
        self.createAndBind()

    def onOK(self, s, e):
        self.createAndBind()
        self.DialogResult = DialogResult.OK
        self.Close()

    def onCancel(self, s, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()

# Entry point
def main():
    form = SharedParamsForm(doc)
    form.ShowDialog()

if __name__ == '__main__': main()