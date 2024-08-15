using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;
using System.Threading.Tasks;

namespace OutputInJSON
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            ElementCategoryFilter windowFilter = new ElementCategoryFilter(BuiltInCategory.OST_Windows);
            ElementCategoryFilter doorFilter = new ElementCategoryFilter(BuiltInCategory.OST_Doors);
            LogicalOrFilter filter = new LogicalOrFilter(windowFilter, doorFilter);

            List<Element> openingElements = new FilteredElementCollector(doc)
                .WherePasses(filter)
                 .WhereElementIsNotElementType()
                .Cast<Element>()
                .ToList();

            var listOpening = new List<Opening>();
            foreach (Element element in openingElements)
            {
                var openingElement = element as FamilyInstance;
                listOpening.Add(new Opening
                {
                    Name = openingElement.Category.Name,
                    Width = openingElement.Symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH).AsValueString(),
                    Height = openingElement.Symbol.get_Parameter(BuiltInParameter.GENERIC_HEIGHT).AsValueString()

                });
            }

            JsonSerializerOptions options = new JsonSerializerOptions()
            {
                Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic),
                WriteIndented = true
            };
            string jsonString = JsonSerializer.Serialize(listOpening, options);

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Текстовый файл (*.json)|*.json";

            if (saveFileDialog.ShowDialog() == true)
            {
                File.WriteAllText(saveFileDialog.FileName, jsonString);
                TaskDialog.Show("!", "Запись в json завершина!");
            }
            return Result.Succeeded;
        }
    }
}
