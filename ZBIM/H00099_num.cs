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

    public class H00099_num : IExternalCommand
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
            StringBuilder stringBuilder_later = new StringBuilder(); //直跑楼梯，不符合的楼梯
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


            /*
            foreach (var item in StairsCollector)
            {
                try
                {
                    Stairs stairs = item as Stairs;
                    List<ElementId> StairsRoliingsId_List = getStairsRailings(document, stairs, StairsRailingCollector);
                    List<ElementId> StairsRunsId_List = stairs.GetStairsRuns().ToList();
                    List<ElementId> StairsLandingsId_List = stairs.GetStairsLandings().ToList();

                    List<Element> StairsRoliings_List = GetItemByIds(StairsRoliingsId_List, StairsRailingCollector);
                    List<Element> StairsRuns_List = GetItemByIds(StairsRunsId_List, StairsRunsCollector);
                    List<Element> StairsLandings_List = GetItemByIds(StairsLandingsId_List, StairsLandingsCollector);

                    stringBuilder.Append($"楼梯：{stairs.Id}");
                    foreach (var stairsroliings in StairsRoliings_List)
                    {
                        stringBuilder.Append($"，栏杆：{stairsroliings.Id}");
                    }
                    foreach (var stairsruns in StairsRuns_List)
                    {
                        stringBuilder.Append($"，梯段：{stairsruns.Id}");
                    }
                    foreach (var stairslandings in StairsLandings_List)
                    {
                        stringBuilder.Append($"，平面：{stairslandings.Id}");
                    }
                    
                    stringBuilder.AppendLine("");
                }
                catch (Exception e)
                {
                    //stringBuilder.Append("," + e.Message);
                }
            }
            */


            //==========找墙体Walls==========
            //根据每块墙的BoundingXYZ坐标判断
            //键：墙的id，值：坐标信息，例如：{Xmin:0,Xmax:1,Ymin:2,Ymax:3,...}
            Dictionary<Element, Dictionary<string, double>> WallBoundingXYZ_Dict = new Dictionary<Element, Dictionary<string, double>>();
            double WallMinX;
            double WallMaxX;
            double WallMinY;
            double WallMaxY;
            double WallMinZ;
            double WallMaxZ;
            double WallXLength;
            double WallYLength;
            double WallXY;  // 表示墙体走向，0表示X方向更长，1表示Y方向更长
            int Wall_findNum = 0;
            int Wall_notfindNum = 0;

            foreach (var wall in WallsCollector)
            {
                if (!WallBoundingXYZ_Dict.ContainsKey(wall))
                {
                    WallBoundingXYZ_Dict.Add(wall, new Dictionary<string, double>());
                }
                try
                {
                    BoundingBoxXYZ boundingBoxXYZ = wall.get_BoundingBox(document.ActiveView);
                    XYZ max = boundingBoxXYZ.Max;
                    XYZ min = boundingBoxXYZ.Min;
                    WallMinX = Get_Right_Number(min.X);
                    WallMaxX = Get_Right_Number(max.X);
                    WallMinY = Get_Right_Number(min.Y);
                    WallMaxY = Get_Right_Number(max.Y);
                    WallMinZ = Get_Right_Number(min.Z);
                    WallMaxZ = Get_Right_Number(max.Z);
                    WallXLength = Math.Round(Utils.FootToMeter(Math.Abs(WallMinX - WallMaxX)), 2);
                    WallYLength = Math.Round(Utils.FootToMeter(Math.Abs(WallMinY - WallMaxY)), 2);
                    if (WallXLength > WallYLength)
                    {
                        WallXY = 0;  // 如果是X轴方向上的长条形，就赋值为0，平台净宽通过X轴计算
                    }
                    else
                    {
                        WallXY = 1; // 如果是Y轴方向上的长条形，就赋值为1，平台净宽通过Y轴计算
                    }
                    WallBoundingXYZ_Dict[wall].Add("WallMinX", WallMinX);
                    WallBoundingXYZ_Dict[wall].Add("WallMaxX", WallMaxX);
                    WallBoundingXYZ_Dict[wall].Add("WallMinY", WallMinY);
                    WallBoundingXYZ_Dict[wall].Add("WallMaxY", WallMaxY);
                    WallBoundingXYZ_Dict[wall].Add("WallMinZ", WallMinZ);
                    WallBoundingXYZ_Dict[wall].Add("WallMaxZ", WallMaxZ);
                    WallBoundingXYZ_Dict[wall].Add("WallXLength", WallXLength);
                    WallBoundingXYZ_Dict[wall].Add("WallYLength", WallYLength);
                    WallBoundingXYZ_Dict[wall].Add("WallXY", WallXY);

                    //stringBuilder.AppendLine($"楼梯名：{wall.Name}，{wall.Id}，X：({WallMinX},{WallMaxX})，Y：({WallMinY},{WallMaxY})，Z：({WallMinZ},{WallMaxZ})");
                    Wall_findNum++;

                }
                catch (Exception)
                {
                    Wall_notfindNum++;
                    //stringBuilder.AppendLine("找不到的墙体：" + wall.Id.ToString() + "，" + wall.Name.ToString());
                    WallBoundingXYZ_Dict.Remove(wall);
                }
            }
            //stringBuilder.AppendLine($"找得到的墙体数量：{Wall_findNum}，找不到的墙体数量：{Wall_notfindNum}");



            //==========找楼梯Stairs==========
            //==========找楼梯对应的梯段和平面==========
            Dictionary<Element, Dictionary<string, double>> StairsBoundingXYZ_Dict = new Dictionary<Element, Dictionary<string, double>>();
            double StairsMinX;
            double StairsMaxX;
            double StairsMinY;
            double StairsMaxY;
            double StairsMinZ;
            double StairsMaxZ;
            double StairsXLength;
            double StairsYLength;
            double StairsWidth;

            int find_Stairs = 0;
            int notfind_Stairs = 0;

            foreach (Element item_stairs in StairsCollector)
            {
                try
                {
                    Stairs stairs = item_stairs as Stairs;
                    List<ElementId> StairsLandingsId_List = stairs.GetStairsLandings().ToList();
                    if (StairsLandingsId_List.Count != 0)  //只有当楼梯的平台数量为0时，才把它当成一个楼梯看待
                    {
                        if (!StairsBoundingXYZ_Dict.ContainsKey(stairs))
                        {
                            StairsBoundingXYZ_Dict.Add(stairs, new Dictionary<string, double>());
                        }
                        try
                        {
                            BoundingBoxXYZ boundingBoxXYZ = stairs.get_BoundingBox(document.ActiveView);
                            XYZ max = boundingBoxXYZ.Max;
                            XYZ min = boundingBoxXYZ.Min;
                            StairsMinX = Get_Right_Number(min.X);
                            StairsMaxX = Get_Right_Number(max.X);
                            StairsMinY = Get_Right_Number(min.Y);
                            StairsMaxY = Get_Right_Number(max.Y);
                            StairsMinZ = Get_Right_Number(min.Z);
                            StairsMaxZ = Get_Right_Number(max.Z);
                            StairsXLength = Math.Round(Utils.FootToMeter(Math.Abs(StairsMinX - StairsMaxX)), 2);
                            StairsYLength = Math.Round(Utils.FootToMeter(Math.Abs(StairsMinY - StairsMaxY)), 2);
                            StairsWidth = Math.Min(StairsXLength, StairsYLength);

                            StairsBoundingXYZ_Dict[stairs].Add("StairsMinX", StairsMinX);
                            StairsBoundingXYZ_Dict[stairs].Add("StairsMaxX", StairsMaxX);
                            StairsBoundingXYZ_Dict[stairs].Add("StairsMinY", StairsMinY);
                            StairsBoundingXYZ_Dict[stairs].Add("StairsMaxY", StairsMaxY);
                            StairsBoundingXYZ_Dict[stairs].Add("StairsMinZ", StairsMinZ);
                            StairsBoundingXYZ_Dict[stairs].Add("StairsMaxZ", StairsMaxZ);
                            StairsBoundingXYZ_Dict[stairs].Add("StairsXLength", StairsXLength);
                            StairsBoundingXYZ_Dict[stairs].Add("StairsYLength", StairsYLength);
                            StairsBoundingXYZ_Dict[stairs].Add("StairsWidth", StairsWidth);

                            //stringBuilder.AppendLine($"楼梯名：{stairs.Name}，{stairs.Id}，X：({StairsMinX},{StairsMaxX})，Y：({StairsMinY},{StairsMaxY})，Z：({StairsMinZ},{StairsMaxZ})");
                            find_Stairs++;
                        }
                        catch (Exception)
                        {
                            notfind_Stairs++;
                            //stringBuilder.AppendLine("找不到的楼梯：" + stairs.Id.ToString() + "，" + stairs.Name.ToString());
                            StairsBoundingXYZ_Dict.Remove(stairs);
                        }
                    }
                }
                catch { }
            }
            //stringBuilder.AppendLine($"找得到的楼梯数量：{find_Stairs}，找不到的楼梯数量：{notfind_Stairs}");



            //========== 遍历每个楼梯 ==========
            bool exist_wall;            //表示一个楼梯内是否存在墙体
            bool exist_stairsroliings;  //表示一个楼梯内是否存在栏杆
            int num_stairsruns;         //梯段数量
            int num_stairslandings;     //平台数量


            foreach (KeyValuePair<Element, Dictionary<string, double>> item_stairs in StairsBoundingXYZ_Dict)
            {
                Stairs stairs = item_stairs.Key as Stairs;

                exist_wall = false;
                exist_stairsroliings = false;
                //========== 先判断楼梯中是否存在墙和栏杆 ==========
                //判断楼梯中是否存在墙
                StairsMinX = item_stairs.Value["StairsMinX"];
                StairsMaxX = item_stairs.Value["StairsMaxX"];
                StairsMinY = item_stairs.Value["StairsMinY"];
                StairsMaxY = item_stairs.Value["StairsMaxY"];
                StairsMinZ = item_stairs.Value["StairsMinZ"];
                StairsMaxZ = item_stairs.Value["StairsMaxZ"];
                StairsXLength = item_stairs.Value["StairsXLength"];
                StairsYLength = item_stairs.Value["StairsYLength"];
                StairsWidth = item_stairs.Value["StairsWidth"];
                //stringBuilder.AppendLine($"楼梯：{stairs.Key.Id}，{stairs.Key.Name}，X轴长度：{StairsXLength}，Y轴长度：{StairsYLength}，X：({StairsMinX},{StairsMaxX})，Y：({StairsMinY},{StairsMaxY})，Z：({StairsMinZ},{StairsMaxZ})");
                //stringBuilder.AppendLine($"楼梯：{stairs.Key.Id}");
                //判断楼梯中是否存在墙
                foreach (KeyValuePair<Element, Dictionary<string, double>> wall in WallBoundingXYZ_Dict)
                {
                    //2、获取墙体的坐标
                    WallMinX = wall.Value["WallMinX"];
                    WallMaxX = wall.Value["WallMaxX"];
                    WallMinY = wall.Value["WallMinY"];
                    WallMaxY = wall.Value["WallMaxY"];
                    WallMinZ = wall.Value["WallMinZ"];
                    WallMaxZ = wall.Value["WallMaxZ"];
                    WallXLength = wall.Value["WallXLength"];
                    WallYLength = wall.Value["WallYLength"];
                    WallXY = wall.Value["WallXY"];
                    if (WallMinX > StairsMaxX + 1 && WallMinY > StairsMaxY + 1) //转向侧存在墙体时，平台净宽应≥1.3m
                    {
                        exist_wall = true;
                    }
                }
                //判断楼梯中是否存在栏杆
                try
                {
                    List<ElementId> StairsRoliingsId_List = getStairsRailings(document, stairs, StairsRailingCollector);
                    //List<Element> StairsRoliings_List = GetItemByIds(StairsRoliingsId_List, StairsRailingCollector);
                    if (StairsRoliingsId_List.Count > 0)
                    {
                        exist_stairsroliings = true;
                    }
                }
                catch (Exception e) { }



                //========== 然后获得楼梯所拥有的梯段、平台、栏杆 ==========
                List<ElementId> StairsRunsId_List = stairs.GetStairsRuns().ToList();
                List<ElementId> StairsLandingsId_List = stairs.GetStairsLandings().ToList();
                List<Element> StairsRuns_List = GetItemByIds(StairsRunsId_List, StairsRunsCollector);
                List<Element> StairsLandings_List = GetItemByIds(StairsLandingsId_List, StairsLandingsCollector);
                num_stairsruns = StairsRunsId_List.Count;
                num_stairslandings = StairsLandingsId_List.Count;

                //梯段数量：1，一定是直跑楼梯
                if (num_stairsruns == 1)
                {
                    double stairsrunsWidth;
                    double X_direction = 0;
                    double Y_direction = 0;
                    double StairsRunsWidth; //梯段宽度
                    double StairsLandingsWidth; //平台宽度
                    //获得梯段的走向
                    foreach (Element item_stairsruns in StairsRuns_List)
                    {
                        StairsRun stairsruns = item_stairsruns as StairsRun;
                        CurveLoop curveLoop = stairsruns.GetStairsPath();
                        foreach (Curve curve in curveLoop)
                        {
                            if (curve is Line line) //判断Curve对象是否为Line对象
                            {
                                line = (Line)curve;
                                X_direction = Get_Right_Number(line.Direction.X);
                                Y_direction = Get_Right_Number(line.Direction.Y);
                            }
                        }
                        StairsRunsWidth = Math.Round(Utils.FootToMeter(stairsruns.LookupParameter("实际梯段宽度").AsDouble()), 2);
                    }
                    //平台数量：1，直跑楼梯
                    if (num_stairslandings == 1)
                    {
                        foreach (Element item_stairsLanding in StairsLandings_List)
                        {
                            StairsLanding stairslandings = item_stairsLanding as StairsLanding;
                            BoundingBoxXYZ boundingBoxXYZ = stairslandings.get_BoundingBox(document.ActiveView);
                            XYZ max = boundingBoxXYZ.Max;
                            XYZ min = boundingBoxXYZ.Min;
                            double StairsLandingsMinX = Get_Right_Number(min.X);
                            double StairsLandingsMaxX = Get_Right_Number(max.X);
                            double StairsLandingsMinY = Get_Right_Number(min.Y);
                            double StairsLandingsMaxY = Get_Right_Number(max.Y);
                            if (X_direction != 0)
                            {
                                StairsLandingsWidth = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinX - StairsLandingsMaxX)), 2);
                            }
                            else
                            {
                                StairsLandingsWidth = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinY - StairsLandingsMaxY)), 2);
                            }

                            if (StairsLandingsWidth < 0.90)
                            {
                                stringBuilder_notpass_runstairs.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < 0.90");
                            }
                            else
                            {
                                stringBuilder_pass.Append(stairs.Id.ToString() + "  ");
                            }
                        }
                    }
                    //平台数量：2，直跑楼梯
                    else if (num_stairslandings == 2)
                    {
                        foreach (Element item_stairsLanding in StairsLandings_List)
                        {
                            StairsLanding stairslandings = item_stairsLanding as StairsLanding;
                            BoundingBoxXYZ boundingBoxXYZ = stairslandings.get_BoundingBox(document.ActiveView);
                            XYZ max = boundingBoxXYZ.Max;
                            XYZ min = boundingBoxXYZ.Min;
                            double StairsLandingsMinX = Get_Right_Number(min.X);
                            double StairsLandingsMaxX = Get_Right_Number(max.X);
                            double StairsLandingsMinY = Get_Right_Number(min.Y);
                            double StairsLandingsMaxY = Get_Right_Number(max.Y);
                            if (X_direction != 0)
                            {
                                StairsLandingsWidth = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinX - StairsLandingsMaxX)), 2);
                            }
                            else
                            {
                                StairsLandingsWidth = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinY - StairsLandingsMaxY)), 2);
                            }

                            if (StairsLandingsWidth < 0.90)
                            {
                                stringBuilder_notpass_runstairs.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < 0.90");
                            }
                            else
                            {
                                stringBuilder_pass.Append(stairs.Id.ToString() + "  ");
                            }
                        }
                    }
                }
                //梯段数量：2，按照平台数量再确定
                else if (num_stairsruns == 2)
                {

                    double X_direction_1 = -2;  //梯段1的X走向
                    double Y_direction_1 = -2;  //梯段1的Y走向
                    double X_direction_2 = -2;  //梯段2的X走向
                    double Y_direction_2 = -2;  //梯段2的Y走向
                    double StairsRunsWidth_1 = 0; //梯段1的宽度
                    double StairsRunsWidth_2 = 0; //梯段2的宽度

                    //获得第一个梯段的走向
                    StairsRun stairsruns_1 = StairsRuns_List[0] as StairsRun;
                    CurveLoop curveLoop_1 = stairsruns_1.GetStairsPath();
                    foreach (Curve curve in curveLoop_1)
                    {
                        if (curve is Line line) //判断Curve对象是否为Line对象
                        {
                            line = (Line)curve;
                            X_direction_1 = Get_Right_Number(line.Direction.X);
                            Y_direction_1 = Get_Right_Number(line.Direction.Y);
                        }
                    }
                    StairsRunsWidth_1 = Math.Round(Utils.FootToMeter(stairsruns_1.LookupParameter("实际梯段宽度").AsDouble()), 2);


                    //获得第二个梯段的走向
                    StairsRun stairsruns_2 = StairsRuns_List[1] as StairsRun;
                    CurveLoop curveLoop_2 = stairsruns_2.GetStairsPath();
                    foreach (Curve curve in curveLoop_2)
                    {
                        if (curve is Line line) //判断Curve对象是否为Line对象
                        {
                            line = (Line)curve;
                            X_direction_2 = Get_Right_Number(line.Direction.X);
                            Y_direction_2 = Get_Right_Number(line.Direction.Y);
                        }
                    }
                    StairsRunsWidth_2 = Math.Round(Utils.FootToMeter(stairsruns_2.LookupParameter("实际梯段宽度").AsDouble()), 2);

                    //平台数量：1，直跑楼梯或L型楼梯.如果两个梯段的方向一样就是直跑楼梯，不一样的就是L型楼梯
                    if (num_stairslandings == 1)
                    {
                        //梯段方向一样，就是直跑楼梯
                        if (X_direction_1 == X_direction_2 || Y_direction_1 == Y_direction_2)
                        {
                            double StairsLandingsWidth; //平台宽度
                            foreach (Element item_stairsLanding in StairsLandings_List)
                            {
                                StairsLanding stairslandings = item_stairsLanding as StairsLanding;
                                BoundingBoxXYZ boundingBoxXYZ = stairslandings.get_BoundingBox(document.ActiveView);
                                XYZ max = boundingBoxXYZ.Max;
                                XYZ min = boundingBoxXYZ.Min;
                                double StairsLandingsMinX = Get_Right_Number(min.X);
                                double StairsLandingsMaxX = Get_Right_Number(max.X);
                                double StairsLandingsMinY = Get_Right_Number(min.Y);
                                double StairsLandingsMaxY = Get_Right_Number(max.Y);
                                if (Math.Abs(X_direction_1) == Math.Abs(X_direction_2))  //两个梯段沿X走向
                                {
                                    StairsLandingsWidth = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinX - StairsLandingsMaxX)), 2);
                                }
                                else  //两个梯段沿Y走向
                                {
                                    StairsLandingsWidth = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinY - StairsLandingsMaxY)), 2);
                                }
                                if (StairsLandingsWidth < 0.90 || StairsLandingsWidth < StairsRunsWidth_1 || StairsLandingsWidth < StairsRunsWidth_2)
                                {
                                    stringBuilder_notpass_runstairs.AppendLine($"{stairs.Id}：{stairslandings.Id}：平台宽度：{StairsLandingsWidth} < 梯段宽度_1：{StairsRunsWidth_1}，梯段宽度_2：{StairsRunsWidth_2}");
                                }
                                else
                                {
                                    stringBuilder_pass.Append(stairs.Id.ToString() + "  ");
                                }
                            }
                        }
                        //梯段方向的绝对值相等，则为X型楼梯
                        else if (Math.Abs(X_direction_1) == Math.Abs(X_direction_2) || Math.Abs(Y_direction_1) == Math.Abs(Y_direction_2))
                        {
                            double StairsLandingsWidth; //平台宽度
                            foreach (Element item_stairsLanding in StairsLandings_List)
                            {
                                StairsLanding stairslandings = item_stairsLanding as StairsLanding;
                                BoundingBoxXYZ boundingBoxXYZ = stairslandings.get_BoundingBox(document.ActiveView);
                                XYZ max = boundingBoxXYZ.Max;
                                XYZ min = boundingBoxXYZ.Min;
                                double StairsLandingsMinX = Get_Right_Number(min.X);
                                double StairsLandingsMaxX = Get_Right_Number(max.X);
                                double StairsLandingsMinY = Get_Right_Number(min.Y);
                                double StairsLandingsMaxY = Get_Right_Number(max.Y);
                                if (Math.Abs(X_direction_1) == Math.Abs(X_direction_2))  //两个梯段沿X走向
                                {
                                    StairsLandingsWidth = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinX - StairsLandingsMaxX)), 2);
                                }
                                else  //两个梯段沿Y走向
                                {
                                    StairsLandingsWidth = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinY - StairsLandingsMaxY)), 2);
                                }
                                if (exist_wall)
                                {
                                    if (StairsLandingsWidth < 1.30 || StairsLandingsWidth < StairsRunsWidth_1 || StairsLandingsWidth < StairsRunsWidth_2)
                                    {
                                        stringBuilder_notpass_exist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：平台宽度：{StairsLandingsWidth} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                    }
                                    else
                                    {
                                        stringBuilder_pass.Append(stairs.Id.ToString() + "  ");
                                    }
                                }
                                else if (exist_stairsroliings)
                                {
                                    if (StairsLandingsWidth < 1.20 || StairsLandingsWidth < StairsRunsWidth_1 || StairsLandingsWidth < StairsRunsWidth_2)
                                    {
                                        stringBuilder_notpass_exist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：平台宽度：{StairsLandingsWidth} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                    }
                                    else
                                    {
                                        stringBuilder_pass.Append(stairs.Id.ToString() + "  ");
                                    }
                                }
                                else
                                {
                                    if (StairsLandingsWidth < 1.20 || StairsLandingsWidth < StairsRunsWidth_1 || StairsLandingsWidth < StairsRunsWidth_2)
                                    {
                                        stringBuilder_notpass_notexist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：平台宽度：{StairsLandingsWidth} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                    }
                                    else
                                    {
                                        stringBuilder_pass.Append(stairs.Id.ToString() + "  ");
                                    }
                                }
                            }
                        }
                        //梯段方向分别为（1，0，0）和（0，1，0）或者是（0，1，0）和（1，0，0），就是L型楼梯，这里就直接归纳为else了
                        else
                        {
                            foreach (Element item_stairsLanding in StairsLandings_List)
                            {
                                double StairsLandingsWidth_X; //平台宽度
                                double StairsLandingsWidth_Y; //平台另一个方向的宽度
                                StairsLanding stairslandings = item_stairsLanding as StairsLanding;
                                BoundingBoxXYZ boundingBoxXYZ = stairslandings.get_BoundingBox(document.ActiveView);
                                XYZ max = boundingBoxXYZ.Max;
                                XYZ min = boundingBoxXYZ.Min;
                                double StairsLandingsMinX = Get_Right_Number(min.X);
                                double StairsLandingsMaxX = Get_Right_Number(max.X);
                                double StairsLandingsMinY = Get_Right_Number(min.Y);
                                double StairsLandingsMaxY = Get_Right_Number(max.Y);
                                StairsLandingsWidth_X = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinX - StairsLandingsMaxX)), 2);
                                StairsLandingsWidth_Y = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinY - StairsLandingsMaxY)), 2);
                                if (Math.Abs(X_direction_1) == 1 && Math.Abs(Y_direction_2) == 1)  //如果梯段1沿着X，梯段2沿着Y
                                {
                                    if (exist_wall)
                                    {
                                        if (StairsLandingsWidth_X < 1.30 || StairsLandingsWidth_Y < 1.30 || StairsLandingsWidth_X < StairsRunsWidth_1 || StairsLandingsWidth_Y < StairsRunsWidth_2)
                                        {
                                            stringBuilder_notpass_exist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：平台宽度：{StairsLandingsWidth_X},{StairsLandingsWidth_Y} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                        }
                                        else
                                        {
                                            stringBuilder_pass.Append(stairs.Id.ToString() + "  ");
                                        }
                                    }
                                    else if (exist_stairsroliings)
                                    {
                                        if (StairsLandingsWidth_X < 1.20 || StairsLandingsWidth_Y < 1.20 || StairsLandingsWidth_X < StairsRunsWidth_1 || StairsLandingsWidth_Y < StairsRunsWidth_2)
                                        {
                                            stringBuilder_notpass_exist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：平台宽度：{StairsLandingsWidth_X},{StairsLandingsWidth_Y} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                        }
                                        else
                                        {
                                            stringBuilder_pass.Append(stairs.Id.ToString() + "  ");
                                        }
                                    }
                                    else
                                    {
                                        if (StairsLandingsWidth_X < 1.20 || StairsLandingsWidth_Y < 1.20 || StairsLandingsWidth_X < StairsRunsWidth_1 || StairsLandingsWidth_Y < StairsRunsWidth_2)
                                        {
                                            stringBuilder_notpass_notexist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：平台宽度：{StairsLandingsWidth_X},{StairsLandingsWidth_Y} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                        }
                                        else
                                        {
                                            stringBuilder_pass.Append(stairs.Id.ToString() + "  ");
                                        }
                                    }
                                }
                                else if (Math.Abs(Y_direction_1) == 1 && Math.Abs(X_direction_2) == 1)  //反之，如果梯段1沿着Y，梯段2沿着X
                                {
                                    if (exist_wall)
                                    {
                                        if (StairsLandingsWidth_X < 1.30 || StairsLandingsWidth_Y < 1.30 || StairsLandingsWidth_X < StairsRunsWidth_2 || StairsLandingsWidth_Y < StairsRunsWidth_1)
                                        {
                                            stringBuilder_notpass_exist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：平台宽度：{StairsLandingsWidth_X},{StairsLandingsWidth_Y} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                        }
                                        else
                                        {
                                            stringBuilder_pass.Append(stairs.Id.ToString() + "  ");
                                        }
                                    }
                                    else if (exist_stairsroliings)
                                    {
                                        if (StairsLandingsWidth_X < 1.20 || StairsLandingsWidth_Y < 1.20 || StairsLandingsWidth_X < StairsRunsWidth_2 || StairsLandingsWidth_Y < StairsRunsWidth_1)
                                        {
                                            stringBuilder_notpass_exist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：平台宽度：{StairsLandingsWidth_X},{StairsLandingsWidth_Y} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                        }
                                        else
                                        {
                                            stringBuilder_pass.Append(stairs.Id.ToString() + "  ");
                                        }
                                    }
                                    else
                                    {
                                        if (StairsLandingsWidth_X < 1.20 || StairsLandingsWidth_Y < 1.20 || StairsLandingsWidth_X < StairsRunsWidth_2 || StairsLandingsWidth_Y < StairsRunsWidth_1)
                                        {
                                            stringBuilder_notpass_notexist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：平台宽度：{StairsLandingsWidth_X},{StairsLandingsWidth_Y} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                        }
                                        else
                                        {
                                            stringBuilder_pass.Append(stairs.Id.ToString() + "  ");
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //平台数量：2，X型楼梯或蛇型楼梯，如果梯段方向相反，就是X型楼梯，如果不是相反数就是蛇型楼梯
                    else if (num_stairslandings == 2) //交叉型楼梯或蜿蜒型楼梯
                    {
                        //获得平台1的宽度
                        StairsLanding stairslandings_1 = StairsLandings_List[0] as StairsLanding;
                        BoundingBoxXYZ boundingBoxXYZ_1 = stairslandings_1.get_BoundingBox(document.ActiveView);
                        XYZ max_1 = boundingBoxXYZ_1.Max;
                        XYZ min_1 = boundingBoxXYZ_1.Min;
                        double StairsLandingsMinX_1 = Get_Right_Number(min_1.X);
                        double StairsLandingsMaxX_1 = Get_Right_Number(max_1.X);
                        double StairsLandingsMinY_1 = Get_Right_Number(min_1.Y);
                        double StairsLandingsMaxY_1 = Get_Right_Number(max_1.Y);
                        double StairsLandingsWidth_1 = 0;

                        //获得平台2的宽度
                        StairsLanding stairslandings_2 = StairsLandings_List[1] as StairsLanding;
                        BoundingBoxXYZ boundingBoxXYZ_2 = stairslandings_2.get_BoundingBox(document.ActiveView);
                        XYZ max_2 = boundingBoxXYZ_2.Max;
                        XYZ min_2 = boundingBoxXYZ_1.Min;
                        double StairsLandingsMinX_2 = Get_Right_Number(min_2.X);
                        double StairsLandingsMaxX_2 = Get_Right_Number(max_2.X);
                        double StairsLandingsMinY_2 = Get_Right_Number(min_2.Y);
                        double StairsLandingsMaxY_2 = Get_Right_Number(max_2.Y);
                        double StairsLandingsWidth_2 = 0;


                        //梯段1和梯段2的方向一样，就是直跑楼梯
                        if (X_direction_1 == X_direction_2 || Y_direction_1 == Y_direction_2)
                        {
                            if (X_direction_1 == X_direction_2)
                            {
                                StairsLandingsWidth_1 = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinX_1 - StairsLandingsMaxX_1)), 2);
                                StairsLandingsWidth_2 = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinX_2 - StairsLandingsMaxX_2)), 2);
                                if (StairsLandingsWidth_1 < 0.90 || StairsLandingsWidth_1 < StairsRunsWidth_1 || StairsLandingsWidth_1 < StairsRunsWidth_2)
                                {
                                    stringBuilder_notpass_runstairs.AppendLine($"{stairs.Id}：{stairslandings_1.Id}：平台宽度：{StairsLandingsWidth_1} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                }
                                if (StairsLandingsWidth_2 < 0.90 || StairsLandingsWidth_2 < StairsRunsWidth_1 || StairsLandingsWidth_2 < StairsRunsWidth_2)
                                {
                                    stringBuilder_notpass_runstairs.AppendLine($"{stairs.Id}：{stairslandings_2.Id}：平台宽度：{StairsLandingsWidth_2} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                }
                            }
                            else if (Y_direction_1 == Y_direction_2)
                            {
                                StairsLandingsWidth_1 = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinY_1 - StairsLandingsMaxY_1)), 2);
                                StairsLandingsWidth_2 = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinY_2 - StairsLandingsMaxY_2)), 2);
                                if (StairsLandingsWidth_1 < 0.90 || StairsLandingsWidth_1 < StairsRunsWidth_1 || StairsLandingsWidth_1 < StairsRunsWidth_2)
                                {
                                    stringBuilder_notpass_runstairs.AppendLine($"{stairs.Id}：{stairslandings_1.Id}：平台宽度：{StairsLandingsWidth_1} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                }
                                if (StairsLandingsWidth_2 < 0.90 || StairsLandingsWidth_2 < StairsRunsWidth_1 || StairsLandingsWidth_2 < StairsRunsWidth_2)
                                {
                                    stringBuilder_notpass_runstairs.AppendLine($"{stairs.Id}：{stairslandings_2.Id}：平台宽度：{StairsLandingsWidth_2} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                }
                            }
                        }
                        //梯段1和梯段2的方向的绝对值一样，就是X型楼梯

                        //else if ((Math.Abs(X_direction_1) == 1 && Math.Abs(X_direction_2) == 1) || (Math.Abs(Y_direction_1) == 1 && Math.Abs(Y_direction_2) == 1))
                        else
                        {
                            stringBuilder_later.AppendLine($"回旋楼梯：{stairs.Id}，梯段数量:{num_stairsruns}，平台数量:{num_stairslandings}");
                            if (Math.Abs(X_direction_1) == 1 && Math.Abs(X_direction_2) == 1)  //两个梯段沿X走向
                            {
                                StairsLandingsWidth_1 = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinX_1 - StairsLandingsMaxX_1)), 2);
                                StairsLandingsWidth_2 = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinX_2 - StairsLandingsMaxX_2)), 2);
                            }
                            else  //两个梯段沿Y走向
                            {
                                StairsLandingsWidth_1 = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinY_1 - StairsLandingsMaxY_1)), 2);
                                StairsLandingsWidth_2 = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinY_2 - StairsLandingsMaxY_2)), 2);
                            }
                            if (exist_wall)
                            {
                                if (StairsLandingsWidth_1 < 1.30 || StairsLandingsWidth_1 < StairsRunsWidth_1 || StairsLandingsWidth_1 < StairsRunsWidth_2)
                                {
                                    stringBuilder_notpass_exist_element.AppendLine($"{stairs.Id}：{stairslandings_1.Id}：平台宽度：{StairsLandingsWidth_1} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                }
                                if (StairsLandingsWidth_2 < 1.30 || StairsLandingsWidth_2 < StairsRunsWidth_1 || StairsLandingsWidth_2 < StairsRunsWidth_2)
                                {
                                    stringBuilder_notpass_exist_element.AppendLine($"{stairs.Id}：{stairslandings_2.Id}：平台宽度：{StairsLandingsWidth_2} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                }
                            }
                            else if (exist_stairsroliings)
                            {
                                if (StairsLandingsWidth_1 < 1.20 || StairsLandingsWidth_1 < StairsRunsWidth_1 || StairsLandingsWidth_1 < StairsRunsWidth_2)
                                {
                                    stringBuilder_notpass_exist_element.AppendLine($"{stairs.Id}：{stairslandings_1.Id}：平台宽度：{StairsLandingsWidth_1} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                }
                                if (StairsLandingsWidth_2 < 1.20 || StairsLandingsWidth_2 < StairsRunsWidth_1 || StairsLandingsWidth_2 < StairsRunsWidth_2)
                                {
                                    stringBuilder_notpass_exist_element.AppendLine($"{stairs.Id}：{stairslandings_2.Id}：平台宽度：{StairsLandingsWidth_2} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                }
                            }
                            else
                            {
                                if (StairsLandingsWidth_1 < 1.20 || StairsLandingsWidth_1 < StairsRunsWidth_1 || StairsLandingsWidth_1 < StairsRunsWidth_2)
                                {
                                    stringBuilder_notpass_notexist_element.AppendLine($"{stairs.Id}：{stairslandings_1.Id}：平台宽度：{StairsLandingsWidth_1} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                }
                                if (StairsLandingsWidth_2 < 1.20 || StairsLandingsWidth_2 < StairsRunsWidth_1 || StairsLandingsWidth_2 < StairsRunsWidth_2)
                                {
                                    stringBuilder_notpass_notexist_element.AppendLine($"{stairs.Id}：{stairslandings_2.Id}：平台宽度：{StairsLandingsWidth_2} < 梯段1_宽度：{StairsRunsWidth_1}，梯段2_宽度：{StairsRunsWidth_2}");
                                }
                            }
                        }
                        //梯段1和梯段2的方向的不一样，就是蛇型楼梯
                        //else
                        // {
                        // stringBuilder_later.Append($"蛇型楼梯：{stairs.Id}");
                        //    }
                    }
                }
                else
                {
                    //回旋楼梯
                    stringBuilder_later.Append($"回旋楼梯：{stairs.Id}");
                }
            }

            stringBuilder.AppendLine("平台净宽不符合基本要求的有：\n" + stringBuilder_notpass_basic);
            stringBuilder.AppendLine("梯段改变方向，无实体端，平台宽度不符合要求的有：\n" + stringBuilder_notpass_notexist_element);
            stringBuilder.AppendLine("梯段改变方向，有实体端，平台宽度不符合要求的有：\n" + stringBuilder_notpass_exist_element);
            stringBuilder.AppendLine("直跑楼梯，平台宽度不符合要求的有：\n" + stringBuilder_notpass_runstairs);
            stringBuilder.AppendLine("还未检测到的有：\n" + stringBuilder_later);
            stringBuilder.AppendLine("符合的楼梯：\n" + stringBuilder_pass);




            if (H00099 == true)
            {
                stringBuilder_result.AppendLine("符合5.3.5民用建筑通用规范 GB55031-2022");
            }
            else
            {
                stringBuilder_result.AppendLine("不符合5.3.5民用建筑通用规范 GB55031-2022");
            }
            stringBuilder = stringBuilder_result.AppendLine(stringBuilder.ToString());

            Utils.PrintLog(stringBuilder.ToString(), "H00099", document);
            TaskDialog.Show("H00099强条检测", stringBuilder.ToString());
            return Result.Succeeded;
            
        }
    }
}
