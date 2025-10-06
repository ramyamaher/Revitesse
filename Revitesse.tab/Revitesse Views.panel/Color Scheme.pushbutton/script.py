from pyrevit import revit, DB
from Autodesk.Revit.UI.Selection import ObjectType
import clr, random

clr.AddReference('PresentationFramework')
from System.Windows.Controls import StackPanel, ComboBox, Label, Button, WrapPanel, Border, Orientation, RadioButton
from System.Windows import Window, Thickness, WindowStartupLocation, SizeToContent, HorizontalAlignment
from System.Windows.Media import SolidColorBrush, Color
from System.Collections.Generic import List

# Alert comes always on top
def alertTopmost(message, title="Alert"):
    win = Window()
    win.Title = title
    win.SizeToContent = SizeToContent.WidthAndHeight
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen
    win.Topmost = True
    win.ResizeMode = 0

    panel = StackPanel(Margin=Thickness(10))
    panel.Children.Add(Label(Content=message, Width=250, HorizontalAlignment=HorizontalAlignment.Left))
    btn = Button(Content="OK", Width=70, Margin=Thickness(0,10,0,0), HorizontalAlignment=HorizontalAlignment.Center)
    btn.Click += lambda sender,args: win.Close()
    panel.Children.Add(btn)
    win.Content = panel
    win.ShowDialog()

# Gradient definitions
GRADIENTS = {
    "Red":      [((255- (i*25)), 0, 0) for i in range(10)],
    "Orange":   [(200 + i*5, 100 + i*20, 0) for i in range(10)],
    "Yellow":   [(200 + i*5, 200 + i*5, 50 + i*20) for i in range(10)],
    "Green":    [(0, (255- (i*25)), 0) for i in range(10)],
    "Cyan":     [(0, 128 + i*12, 128 + i*12) for i in range(10)],
    "Blue":     [(0, 0, (255- (i*25))) for i in range(10)],
    "Purple":   [(90 + i*15, 0, 90 + i*15) for i in range(10)],
    "Brown":    [(101 + i*15, 67 + i*15, 33 + i*10) for i in range(10)],
}

# Colors
def randomColor():
    return DB.Color(random.randint(50, 255), random.randint(50, 255), random.randint(50, 255))

def rgbToColor(r,g,b):
    return Color.FromRgb(byte(r), byte(g), byte(b))

def byte(v):
    return max(0, min(255, int(v)))

def getParameterByName(elem, name):
    for p in elem.Parameters:
        if p.Definition and p.Definition.Name == name: return p
    return None

def getParameterValue(param):
    if not param: return None
    st = param.StorageType
    if st == DB.StorageType.String: return param.AsString()
    elif st == DB.StorageType.Integer: return param.AsInteger()
    elif st == DB.StorageType.Double: return param.AsDouble()
    elif st == DB.StorageType.ElementId:
        eid = param.AsElementId()
        if eid and eid != DB.ElementId.InvalidElementId: return eid
        return None
    return None

def createFilterRule(param, value):
    st = param.StorageType
    provider = DB.ParameterValueProvider(param.Id)
    if st == DB.StorageType.String: return DB.FilterStringRule(provider, DB.FilterStringEquals(), str(value))
    elif st == DB.StorageType.Integer: return DB.FilterIntegerRule(provider, DB.FilterNumericEquals(), int(value))
    elif st == DB.StorageType.ElementId:
        if isinstance(value, DB.ElementId): return DB.FilterElementIdRule(provider, DB.FilterNumericEquals(), value)
        try: return DB.FilterElementIdRule(provider, DB.FilterNumericEquals(), DB.ElementId(int(value)))
        except: return None
    elif st == DB.StorageType.Double: return DB.FilterDoubleRule(provider, DB.FilterNumericEquals(), float(value), 1e-6)
    return None

def canonicalKey(val):
    if isinstance(val, DB.ElementId): return "eid:" + str(val)
    if isinstance(val, int): return "int:" + str(val)
    if isinstance(val, float): return "dbl:" + repr(val)
    if isinstance(val, str): return "str:" + val
    return "repr:" + str(val)

def valueLabel(val):
    if isinstance(val, DB.ElementId):
        try:
            el = revit.doc.GetElement(val)
            if el is not None and hasattr(el, "Name"): return el.Name
        except: pass
        return str(val)
    return str(val)

def isFilterable(param):
    if not param or not param.Definition: return False
    if param.IsReadOnly: return False
    st = param.StorageType
    if st not in (DB.StorageType.String, DB.StorageType.Integer, DB.StorageType.Double, DB.StorageType.ElementId): return False
    return True

# Gradient generator
def generateGradientColors(name, count):
    if name == "Random" or name not in GRADIENTS: 
        return [randomColor() for _ in range(count)]
    colors = GRADIENTS[name]
    return [DB.Color(r,g,b) for r,g,b in colors[:count]]

# UI
class ParameterPicker(Window):
    def __init__(self, parameterNames):
        self.Title = "Choose Parameter for Color Scheme"
        self.Width = 400
        self.Height = 400
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True
        self.selectedGradient = None
        self.random_preview_colors = [randomColor() for _ in range(10)]

        panel = StackPanel(Margin=Thickness(10))
        panel.Children.Add(Label(Content="Select parameter for color scheme:"))

        self.parameterCombo = ComboBox(ItemsSource=parameterNames)
        panel.Children.Add(self.parameterCombo)

        panel.Children.Add(Label(Content="Select a gradient:"))
        self.radio_buttons = []

        for gradientName, colors in list(GRADIENTS.items()) + [("Random", self.random_preview_colors)]:
            gradPanel = StackPanel(Orientation=Orientation.Horizontal, Margin=Thickness(0,5,0,0))

            # Radio button
            radio = RadioButton(Content=gradientName, GroupName="Gradients", Width=140)
            radio.Checked += self.gradientChecked
            gradPanel.Children.Add(radio)
            self.radio_buttons.append(radio)

            # Color boxes
            colorBoxPanel = WrapPanel()
            colorBoxPanel.Margin = Thickness(0)
            colorBoxPanel.HorizontalAlignment = 0
            if colors:
                for c in colors:
                    if isinstance(c, DB.Color): r, g, b = c.Red, c.Green, c.Blue
                    else: r, g, b = c
                    col = SolidColorBrush(Color.FromRgb(byte(r), byte(g), byte(b)))
                    border = Border(Width=20, Height=20, Background=col, Margin=Thickness(0))
                    colorBoxPanel.Children.Add(border)

            gradPanel.Children.Add(colorBoxPanel)
            panel.Children.Add(gradPanel)

        # Action buttons
        buttonPanel = StackPanel(Orientation=Orientation.Horizontal, Margin=Thickness(0,10,0,0))
        buttonApply = Button(Content="Apply Colors", Width=100, Margin=Thickness(10))
        buttonApply.Click += self.applyClicked
        buttonPanel.Children.Add(buttonApply)

        buttonLegend = Button(Content="Legend", Width=100, Margin=Thickness(10))
        buttonLegend.Click += self.legendClicked
        buttonPanel.Children.Add(buttonLegend)

        buttonClose = Button(Content="Close", Width=100, Margin=Thickness(10))
        buttonClose.Click += self.closeClicked
        buttonPanel.Children.Add(buttonClose)

        panel.Children.Add(buttonPanel)
        self.Content = panel

    def gradientChecked(self, sender, args):
        self.selectedGradient = sender.Content

    def applyClicked(self, sender, args):
        if not self.parameterCombo.SelectedItem:
            alertTopmost("Please select a parameter.")
            return
        if not self.selectedGradient:
            alertTopmost("Please select a gradient.")
            return
        applyColors(self.parameterCombo.SelectedItem, self.selectedGradient)

    def legendClicked(self, sender, args):
        if not self.parameterCombo.SelectedItem:
            alertTopmost("Please select a parameter.")
            return
        if not self.selectedGradient:
            alertTopmost("Please select a gradient.")
            return
        create_legend(self.parameterCombo.SelectedItem, self.selectedGradient)

    def closeClicked(self, sender, args):
        self.Close()

# Pick element and get category
picked = revit.uidoc.Selection.PickObject(ObjectType.Element, "Pick an element")
sourceElement = revit.doc.GetElement(picked.ElementId)
if not sourceElement or not sourceElement.Category: alertTopmost("Selected element has no category.", exitscript=True)
categoryId = sourceElement.Category.Id
categoryName = sourceElement.Category.Name

allCategoryElements = [
    e for e in DB.FilteredElementCollector(revit.doc).WhereElementIsNotElementType()
    if e.Category and e.Category.Id == categoryId
]
if not allCategoryElements: alertTopmost("No elements found for this category.", exitscript=True)

# Collect parameters
parameterCounts = {}
parameterValuesByName = {}
for e in allCategoryElements:
    for p in e.Parameters:
        if p.Definition:
            name = p.Definition.Name
            if name not in parameterCounts:
                parameterCounts[name] = 0
                parameterValuesByName[name] = set()
            parameterCounts[name] += 1
            val = getParameterValue(p)
            if val is not None: parameterValuesByName[name].add(val)

parameterNames = [
    n for n in parameterCounts
    if parameterCounts[n] == len(allCategoryElements)
    and isFilterable(getParameterByName(allCategoryElements[0], n))
]
parameterNames.sort()
if not parameterNames: alertTopmost("No filterable parameters exist in all elements of this category.", exitscript=True)

# Solid drafting pattern
solidDraftingPatternId = None
for f in DB.FilteredElementCollector(revit.doc).OfClass(DB.FillPatternElement).ToElements():
    try:
        fp = f.GetFillPattern()
        if fp and fp.IsSolidFill and fp.Target == DB.FillPatternTarget.Drafting:
            solidDraftingPatternId = f.Id
            break
    except: continue
if not solidDraftingPatternId: alertTopmost("Could not find a solid drafting pattern."); exit()

# Gradient generator
def generateGradientColors(name, count):
    if name == "Random" or name not in GRADIENTS: return [randomColor() for _ in range(count)]
    colors = GRADIENTS[name]
    return [DB.Color(r,g,b) for r,g,b in colors[:count]]

# Apply Colors
def applyColors(paramName, gradName):
    global forms
    t = DB.Transaction(revit.doc, "Apply Color Overrides")
    t.Start()
    try:
        param = getParameterByName(allCategoryElements[0], paramName)
        if not param: alertTopmost("Cannot find parameter object.", exitscript=True)

        view = revit.doc.ActiveView
        rawValues = list(parameterValuesByName[paramName])
        rawValues.sort(key=lambda v: valueLabel(v).lower())
        values = [(v, canonicalKey(v)) for v in rawValues]
        colorsList = generateGradientColors(gradName, len(values))
        colorMap = {}

        for (originalValue, key), col in zip(values, colorsList):
            filterName = categoryName + "_" + paramName + "_" + valueLabel(originalValue)
            existingFilter = next((f for f in DB.FilteredElementCollector(revit.doc).OfClass(DB.ParameterFilterElement) if f.Name == filterName), None)

            if existingFilter:
                if existingFilter.Id not in view.GetFilters(): view.AddFilter(existingFilter.Id)
                filterElem = existingFilter
            else:
                rule = createFilterRule(param, originalValue)
                if not rule: continue
                rulesList = List[DB.FilterRule]([rule])
                paramFilter = DB.ElementParameterFilter(rulesList)
                categories = List[DB.ElementId]([categoryId])
                filterElem = DB.ParameterFilterElement.Create(revit.doc, filterName, categories, paramFilter)
                view.AddFilter(filterElem.Id)

            # Graphics Overrides
            ogs = DB.OverrideGraphicSettings()
            colorMap[key] = col

            if sourceElement.Category.CategoryType == DB.CategoryType.Model:
                ogs.SetCutForegroundPatternColor(col)
                ogs.SetCutForegroundPatternId(solidDraftingPatternId)
                ogs.SetCutLineColor(col)
                ogs.SetProjectionLineColor(col)
                ogs.SetSurfaceForegroundPatternColor(col)
            else:
                ogs.SetProjectionLineColor(col)
                ogs.SetSurfaceForegroundPatternColor(col)
                ogs.SetSurfaceForegroundPatternId(solidDraftingPatternId)
                ogs.SetCutForegroundPatternColor(col)
                ogs.SetCutForegroundPatternId(solidDraftingPatternId)
                ogs.SetCutLineColor(col)

            view.SetFilterOverrides(filterElem.Id, ogs)

        t.Commit()
        alertTopmost("Colors applied successfully.")
    except Exception as e:
        if t.HasStarted() and not t.HasEnded(): t.RollBack()
        alertTopmost("Failed to apply colors:\n" + str(e))

# Legend Creation
def create_legend(paramName, gradName):
    legends = [vw for vw in DB.FilteredElementCollector(revit.doc).OfClass(DB.View).ToElements() if vw.ViewType == DB.ViewType.Legend]
    if not legends:
        alertTopmost("No legend view exists. Please create one manually first.")
        return

    t2 = DB.Transaction(revit.doc, "Create Color Legend")
    t2.Start()
    try:
        newIdLegend = legends[0].Duplicate(DB.ViewDuplicateOption.Duplicate)
        newLegend = revit.doc.GetElement(newIdLegend)
        newLegend.Name = "Color Scheme - " + categoryName + " - " + paramName

        textnoteType = DB.FilteredElementCollector(revit.doc).OfClass(DB.TextNoteType).FirstElement()
        filledType = DB.FilteredElementCollector(revit.doc).OfClass(DB.FilledRegionType).FirstElement()
        if not textnoteType or not filledType: raise Exception("Missing TextNoteType or FilledRegionType")

        y = 0.0
        spacing = 0.15
        height = 0.35
        width = height * 1.6
        initialX = 2.0

        rawValues = list(parameterValuesByName[paramName])
        rawValues.sort(key=lambda v: valueLabel(v).lower())
        values = [(v, canonicalKey(v)) for v in rawValues]

        colorsList = generateGradientColors(gradName, len(values))

        for (originalValue, key), col in zip(values, colorsList):
            lbl = valueLabel(originalValue)
            pt = DB.XYZ(0, y, 0)
            DB.TextNote.Create(revit.doc, newLegend.Id, pt, lbl, textnoteType.Id)

            pt0 = DB.XYZ(initialX, y, 0)
            pt1 = DB.XYZ(initialX, y + height, 0)
            pt2 = DB.XYZ(initialX + width, y + height, 0)
            pt3 = DB.XYZ(initialX + width, y, 0)

            loop = DB.CurveLoop()
            loop.Append(DB.Line.CreateBound(pt0, pt1))
            loop.Append(DB.Line.CreateBound(pt1, pt2))
            loop.Append(DB.Line.CreateBound(pt2, pt3))
            loop.Append(DB.Line.CreateBound(pt3, pt0))
            loops = List[DB.CurveLoop]([loop])

            reg = DB.FilledRegion.Create(revit.doc, filledType.Id, newLegend.Id, loops)
            ogs = DB.OverrideGraphicSettings()
            ogs.SetSurfaceForegroundPatternColor(col)
            ogs.SetCutForegroundPatternColor(col)
            ogs.SetSurfaceForegroundPatternId(solidDraftingPatternId)
            ogs.SetCutForegroundPatternId(solidDraftingPatternId)
            newLegend.SetElementOverrides(reg.Id, ogs)

            y -= (height + spacing)

        t2.Commit()
        alertTopmost("Legend updated: " + newLegend.Name)
    except Exception as e:
        if t2.HasStarted() and not t2.HasEnded(): t2.RollBack()
        alertTopmost("Failed creating legend:\n" + str(e))

# Show UI
form = ParameterPicker(parameterNames)
form.ShowDialog()
