using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using ZBIMUtils;


namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]

    public class TestStairRunDirection99 : IExternalCommand
    {
        private static double Get_Right_Number(double number)
        {
            if (number.ToString().Contains("E"))
            {
                return 0;
            }
            else
            {
                return number;
            }
        }
        public static bool IsNumeral(string input)//shuz数字 numeral
        {
            foreach (char ch in input)
            {
                if (ch < '0' || ch > '9')
                {
                    return false;
                }
            }
            return true;
        }
        //获取楼梯相关栏杆
        private static List<ElementId> getStairsRailings(Document doc, Stairs ele, FilteredElementCollector StairsRailingCollector)
        {
            Dictionary<string, List<ElementId>> stairMap = new Dictionary<string, List<ElementId>>();
            foreach (var item in StairsRailingCollector)
            {
                Element stairsRailing = item as Element;
                Railing railing = item as Railing;
                //List<ElementId> stairsRailingList = new List<ElementId>();
                if (railing != null && railing.HasHost)
                {
                    Element stairEle = doc.GetElement(railing.HostId);
                    if (stairEle.Category.Id.IntegerValue == (int)(BuiltInCategory.OST_Stairs))
                    {
                        if (stairMap.ContainsKey(stairEle.Id.ToString()))
                        {
                            stairMap[stairEle.Id.ToString()].Add(stairsRailing.Id);
                        }
                        else
                        {
                            List<ElementId> stairsRailingList = new List<ElementId>();
                            stairsRailingList.Add(stairsRailing.Id);
                            stairMap.Add(stairEle.Id.ToString(), stairsRailingList);
                        }
                    }
                }
            }
            if (stairMap.ContainsKey(ele.Id.ToString()))
            {
                return stairMap[ele.Id.ToString()];
            }
            else
            {
                return null;
            }
        }

        private static List<Element> GetItemById(ElementId elementId, FilteredElementCollector elementCollector)
        {
            List<Element> element_list = new List<Element>();

            foreach (Element element in elementCollector)
            {
                if (element.Id.Equals(elementId))
                {
                    element_list.Add(element);
                    break;
                }
            }
            return element_list;
        }

        private static List<Element> GetItemByIds(List<ElementId> elementIds, FilteredElementCollector elementCollector)
        {

            List<ElementId> temp_element_list = elementIds;
            List<Element> element_list = new List<Element>();

            foreach (ElementId elementId in temp_element_list)
            {
                foreach (Element element in elementCollector)
                {
                    if (element.Id.Equals(elementId))
                    {
                        element_list.Add(element);
                    }
                }

            }
            return element_list;
        }




        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            Document document = commandData.Application.ActiveUIDocument.Document;
            StringBuilder stringBuilder = new StringBuilder();
            StringBuilder stringBuilder_pass = new StringBuilder();
            StringBuilder stringBuilder_notpass_basic = new StringBuilder(); //不符合最基本净宽要求的楼梯
            StringBuilder stringBuilder_notpass_exist_element = new StringBuilder(); //梯段改变方向，无实体端，不符合的楼梯
            StringBuilder stringBuilder_notpass_notexist_element = new StringBuilder(); //梯段改变方向，有实体端，不符合的楼梯
            StringBuilder stringBuilder_notpass_runstairs = new StringBuilder(); //直跑楼梯，不符合的楼梯
            StringBuilder stringBuilder_result = new StringBuilder();
            bool H00099 = true;


            //过滤：栏杆
            ElementCategoryFilter StairsRailingFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRailing);
            //过滤：楼梯
            ElementCategoryFilter StairsFilter = new ElementCategoryFilter(BuiltInCategory.OST_Stairs);
            //过滤：平面
            ElementCategoryFilter StairsLandingsFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsLandings);
            //过滤：梯段
            ElementCategoryFilter StairsRunsFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRuns);
            //过滤：墙体
            ElementCategoryFilter WallsFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);


            //收集器：栏杆
            FilteredElementCollector StairsRailingCollector = new FilteredElementCollector(document).WherePasses(StairsRailingFilter);
            //FilteredElementCollector StairsRailingCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StairsRailing);
            //收集器：楼梯
            FilteredElementCollector StairsCollector = new FilteredElementCollector(document).WherePasses(StairsFilter);
            //FilteredElementCollector StairsCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Stairs);
            //收集器：平面
            FilteredElementCollector StairsLandingsCollector = new FilteredElementCollector(document).WherePasses(StairsLandingsFilter);
            //FilteredElementCollector StairsLandingsCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StairsLandings);
            //收集器：梯段
            FilteredElementCollector StairsRunsCollector = new FilteredElementCollector(document).WherePasses(StairsRunsFilter);
            //收集器：墙体
            FilteredElementCollector WallsCollector = new FilteredElementCollector(document).WherePasses(WallsFilter);
            //FilteredElementCollector WallsCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Walls);



            int num = 1;
            foreach (var item in StairsCollector)
            {

                try
                {
                    Stairs stairs = item as Stairs;
                    //List<ElementId> StairsRoliingsId_List = getStairsRailings(document, stairs, StairsRailingCollector);
                    List<ElementId> StairsRunsId_List = stairs.GetStairsRuns().ToList();
                    //List<ElementId> StairsLandingsId_List = stairs.GetStairsLandings().ToList();

                    //List<Element> StairsRoliings_List = GetItemByIds(StairsRoliingsId_List, StairsRailingCollector);
                    List<Element> StairsRuns_List = GetItemByIds(StairsRunsId_List, StairsRunsCollector);
                    //List<Element> StairsLandings_List = GetItemByIds(StairsLandingsId_List, StairsLandingsCollector);
                    if (true
                        //StairsLandingsId_List.Count != 0 
                        //!(StairsRunsId_List.Count == 2 && StairsLandingsId_List.Count == 2) &&
                        //!(StairsRunsId_List.Count == 1 && StairsLandingsId_List.Count == 1) &&
                        //!(StairsRunsId_List.Count == 2 && StairsLandingsId_List.Count == 1) 
                      )
                    {
                        //stringBuilder.Append($"楼梯：{stairs.Id}，梯段：{StairsRunsId_List.Count}，平台：{StairsLandingsId_List.Count}");
                        //stringBuilder.AppendLine("");

                        stringBuilder.Append($"{num}、楼梯：{stairs.Id}");
                        foreach (StairsRun stairsruns in StairsRuns_List)
                        {
                            double X_direction = 0;
                            double Y_direction = 0;
                            CurveLoop curveLoop = stairsruns.GetStairsPath();
                            foreach (Curve curve in curveLoop)
                            {
                                if (curve is Line line) //判断Curve对象是否为Line对象
                                {
                                    line = (Line)curve;
                                    //最重要的部分！！！！！！！！！！！！！！！！！！！！！！！！！！
                                    X_direction = Get_Right_Number(line.Direction.X);
                                    Y_direction = Get_Right_Number(line.Direction.Y);
                                }
                            }
                            if (Math.Abs(X_direction) == 1)
                            {
                                stringBuilder.Append($"，梯段：{stairsruns.Id}，X");
                            }
                            else if (Math.Abs(Y_direction) == 1)
                            {
                                stringBuilder.Append($"，梯段：{stairsruns.Id}，Y");
                            }
                            else 
                            {
                                stringBuilder.Append($"，梯段：{stairsruns.Id}，Unknown");
                            }
                        }
                        stringBuilder.AppendLine("");

                        num++;
                    }
                    //foreach (var stairsroliings in StairsRoliings_List)   stringBuilder.Append($"，栏杆：{stairsroliings.Id}");
                    //foreach (var stairsruns in StairsRuns_List)
                    //stringBuilder.Append($"，梯段：{stairsruns.Id}");
                    

                    
                    //foreach (var stairslandings in StairsLandings_List)
                    //stringBuilder.Append($"，平台：{stairslandings.Id}");


                }
                catch (Exception e)
                {
                    //stringBuilder.Append("," + e.Message);
                }
            }
            



            TaskDialog.Show("FilterStairs", stringBuilder_result + stringBuilder.ToString());
            return Result.Succeeded;
        }
    }
}
