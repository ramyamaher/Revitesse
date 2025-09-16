from pyrevit import revit, DB, forms
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc

# 1. Get model categories
modelCategories = [cat for cat in doc.Settings.Categories if cat.CategoryType == DB.CategoryType.Model]
categoryNames = sorted([cat.Name for cat in modelCategories])
categoryNameToCategory = {cat.Name: cat for cat in modelCategories}

# 2. Create grouped categories for single dialog
entireModelCategoryNames = [name + " " for name in categoryNames]
groupedCategories = {"Active View": categoryNames, "Entire Model": entireModelCategoryNames}

# 3. Single dialog with scope dropdown and category selection
chosenCategoryNames = forms.SelectFromList.show(groupedCategories, title="Select Scope and Categories", groupSelectorTitle="Scope", multiselect=True)

if not chosenCategoryNames: forms.alert("No categories selected.", exitscript=True)

# 4. Determine scope based on suffix and clean category names
scopeChoice = "Active View"  # Default
cleanedCategoryNames = []

for categoryName in chosenCategoryNames:
    if categoryName.endswith(" "):
        scopeChoice = "Entire Model"
        cleanedCategoryNames.append(categoryName.rstrip())
    else: cleanedCategoryNames.append(categoryName)

chosenCategories = [categoryNameToCategory[name] for name in cleanedCategoryNames]

# 5. Collect elements based on scope
def getTargetElements(scope, categories):
    if scope == "Active View": collector = DB.FilteredElementCollector(doc, doc.ActiveView.Id)
    else: collector = DB.FilteredElementCollector(doc)
    
    elements = collector.WhereElementIsNotElementType().ToElements()
    
    return [
        e for e in elements
        if e.Category and any(e.Category.Id == cat.Id for cat in categories)
    ]

targetElements = getTargetElements(scopeChoice, chosenCategories)

if not targetElements:   forms.alert("No instances found for selected categories in " + scopeChoice.lower() + ".", exitscript=True)

elementIds = List[DB.ElementId]([e.Id for e in targetElements])
uidoc.Selection.SetElementIds(elementIds)