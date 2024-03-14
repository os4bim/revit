import clr
import System
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import * 

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document


#Is element type

print('#'*20)
print('Is element type')
print('-'*20)

PA1 = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsElementType().FirstElement()

for p in PA1.Parameters:
    print(p.Definition.Name)
    try:
        print(p.GUID)
    except:
        print(p.Definition.BuiltInParameter)
    print('-'*20)
    
#Is NOT element type

print('#'*20)
print('Is NOT element type')
print('-'*20)

PA2 = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().FirstElement()

for p in PA2.Parameters:
    print(p.Definition.Name)
    try:
        print(p.GUID)
    except:
        print(p.Definition.BuiltInParameter)
    print('-'*20)
