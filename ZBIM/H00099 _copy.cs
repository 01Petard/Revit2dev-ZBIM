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

    public class H00099_copy : IExternalCommand
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


            //过滤：栏杆
            ElementCategoryFilter StairsRailingFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRailing);
            //过滤：楼梯
            ElementCategoryFilter StairsFilter = new ElementCategoryFilter(BuiltInCategory.OST_Stairs);
            //过滤：平台
            ElementCategoryFilter StairsLandingsFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsLandings);
            //过滤：梯段
            ElementCategoryFilter StairsRunsFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRuns);


            //收集器：栏杆
            FilteredElementCollector StairsRailingCollector = new FilteredElementCollector(document).WherePasses(StairsRailingFilter);
            //收集器：楼梯
            FilteredElementCollector StairsCollector = new FilteredElementCollector(document).WherePasses(StairsFilter);
            //收集器：平台
            FilteredElementCollector StairsLandingsCollector = new FilteredElementCollector(document).WherePasses(StairsLandingsFilter);
            //收集器：梯段
            FilteredElementCollector StairsRunsCollector = new FilteredElementCollector(document).WherePasses(StairsRunsFilter);


            int num_stairs = 0;
            foreach (var item in StairsCollector)
            {
                try
                {
                    Stairs stairs = item as Stairs;
                    //获得Id
                    //List<ElementId> StairsRoliingsId_List = getStairsRailings(document, stairs, StairsRailingCollector);
                    List<ElementId> StairsRunsId_List = stairs.GetStairsRuns().ToList();
                    List<ElementId> StairsLandingsId_List = stairs.GetStairsLandings().ToList();

                    //通过Id获得Element
                    //List<Element> StairsRoliings_List = GetItemByIds(StairsRoliingsId_List, StairsRailingCollector);
                    //List<Element> StairsRuns_List = GetItemByIds(StairsRunsId_List, StairsRunsCollector);
                    //List<Element> StairsLandings_List = GetItemByIds(StairsLandingsId_List, StairsLandingsCollector);

                    int num_stairsruns = StairsRunsId_List.Count;
                    int num_stairslandings = StairsLandingsId_List.Count;

                    num_stairs++;
                    //打印查看楼梯，以及楼梯所有的栏杆、梯段、平台
                    stringBuilder.Append($"{num_stairs}、楼梯：{stairs.Id}，梯段：{num_stairsruns}，平台：{num_stairslandings}");
                    //foreach (var stairsroliings in StairsRoliings_List)
                    //{
                        //stringBuilder.Append($"，栏杆：{stairsroliings.Id}");
                    //}
                    foreach (var stairsruns in StairsRunsId_List)
                    {
                        stringBuilder.Append($"，梯段：{stairsruns}");
                    }
                    foreach (var stairslandings in StairsLandingsId_List)
                    {
                        stringBuilder.Append($"，平台：{stairslandings}");
                    } 
                    stringBuilder.AppendLine("");
                }
                catch (Exception e)
                {
                    //stringBuilder.Append("," + e.Message);
                }
            }




            //Utils.PrintLog(stringBuilder.ToString(), "H00099", document);
            TaskDialog.Show("FilterStairs", stringBuilder.ToString());
            return Result.Succeeded;
            
        }
    }
}
