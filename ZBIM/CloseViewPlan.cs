using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using ZBIMUtils;


namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CloseViewPlan : IExternalCommand
    {
        

        public static void switchto2D(UIDocument uiDocument, Document document)
        {
            Type type = document.ActiveView.GetType();
            if (!type.Equals(typeof(ViewPlan)))
            {
                ViewPlan viewPlan = null;
                //过滤所有的3D视图
                FilteredElementCollector collector = new FilteredElementCollector(document).OfClass(typeof(ViewPlan));
                foreach (ViewPlan view in collector)
                {
                    if (view.Name.Contains("1F"))
                    {
                        viewPlan = view;
                    }
                }
                uiDocument.ActiveView = viewPlan;//设置所取得的3D视图为当前视图
            }
        }

        public static void switchto3D(UIDocument uiDocument, Document document)
        {
            Type type2 = document.ActiveView.GetType();
            if (!type2.Equals(typeof(View3D)))
            {
                View3D view3D = null;
                //过滤所有的3D视图
                FilteredElementCollector collector = new FilteredElementCollector(document).OfClass(typeof(View3D));
                foreach (View3D view in collector)
                {
                    if (view.Name.Contains("{三维}"))
                    {
                        view3D = view;
                    }
                }
                uiDocument.ActiveView = view3D;//设置所取得的3D视图为当前视图
            }
        }

        public static void CloseCurrentView(UIDocument uiDocument)
        {
            //获得当前的视图
            var activeView = uiDocument.ActiveGraphicalView;
            //获得当前打开的视图
            var openUIViews = uiDocument.GetOpenUIViews();
            if (openUIViews.Count > 1)//如果当前已打开的视图个数少于2的话 关闭当前视图会抛异常
            {
                var targetView = openUIViews.FirstOrDefault(v => v.ViewId == activeView.Id);
                if (targetView != null)
                {
                    targetView.Close();
                }
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            Document document = commandData.Application.ActiveUIDocument.Document;






            //假设我们现在在视图A，下面我们随便跳转一个视图B
            View viewPlan1F = null;
            View viewPlan2F = null;
            foreach (ViewPlan view in new FilteredElementCollector(document).OfClass(typeof(ViewPlan)))
            {
                if (view.Name.Contains("1F")) { viewPlan1F = view; }
                if (view.Name.Contains("15F")) { viewPlan2F = view; }
            }
            bool switched = true;
            //跳转视图B
            //Utils.SwitchTo2D(document, uiDocument, ref viewPlan1F, ref switched);
            //关闭当前这个视图B
            //CloseCurrentView(uiDocument);
            //跳转视图B
            Utils.SwitchTo2D(document, uiDocument, ref viewPlan2F, ref switched);
            //关闭当前这个视图B
            //CloseCurrentView(uiDocument);
            //跳回视图A
            //Utils.SwitchBack(uiDocument, viewPlan, switched);














            return Result.Succeeded;
        }
    }
}
