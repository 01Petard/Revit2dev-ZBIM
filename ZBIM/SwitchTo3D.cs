using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    class SwitchTo3D : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;
            Type type = doc.ActiveView.GetType();
            if (!type.Equals(typeof(View3D)))
            {
                View3D view3D = null;
                //过滤所有的3D视图
                FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(View3D));
                foreach (View3D view in collector)
                {
                    if (view.Name.Contains("{3D}"))
                    {
                        view3D = view;
                    }
                }
                uiDoc.ActiveView = view3D;//设置所取得的3D视图为当前视图
            }
            TaskDialog.Show("测试", "我是3D！");
            return Result.Succeeded;
        }
    }
}
