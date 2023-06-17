using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    class SwitchTo2D : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;
            Type type = doc.ActiveView.GetType();
            if (!type.Equals(typeof(ViewPlan)))
            {
                ViewPlan viewPlan = null;
                //过滤所有的3D视图
                FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan));
                foreach (ViewPlan view in collector)
                {
                    if (view.Name.Contains("1#-23#、27#20F56.400（建）"))
                    {
                        viewPlan = view;
                    }
                }
                uiDoc.ActiveView = viewPlan;//设置所取得的3D视图为当前视图
            }
            TaskDialog.Show("测试", "我是2D！");
            return Result.Succeeded;
        }
    }
}
