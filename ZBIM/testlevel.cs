using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;


namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    public class testlevel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            Document document = commandData.Application.ActiveUIDocument.Document;
            StringBuilder stringBuilder = new StringBuilder();
            FilteredElementCollector collectorLevels = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Levels);

            ElementId viewid = document.ActiveView.Id;


            List<Element> elementsLevels = collectorLevels.OfCategory(BuiltInCategory.OST_Levels).ToList<Element>();

            foreach (var item in collectorLevels)
            {
                stringBuilder.AppendLine(item.Name);
            }

            TaskDialog.Show("H00018强条检测", stringBuilder.ToString());


            return Result.Succeeded;
        }
    }
}
