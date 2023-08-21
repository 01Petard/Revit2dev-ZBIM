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

    public class H00099 : IExternalCommand
    {
        #region H00099强条涉及的静态方法
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
        #endregion



        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            #region 步骤0.初始化的过滤器、字符串等值
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            Document document = commandData.Application.ActiveUIDocument.Document;
            StringBuilder stringBuilder = new StringBuilder();
            StringBuilder stringBuilder_pass = new StringBuilder();
            StringBuilder stringBuilder_notpass_basic = new StringBuilder(); //不符合最基本净宽要求的楼梯
            StringBuilder stringBuilder_notpass_exist_element = new StringBuilder(); //梯段改变方向，无实体端，不符合的楼梯
            StringBuilder stringBuilder_notpass_notexist_element = new StringBuilder(); //梯段改变方向，有实体端，不符合的楼梯
            StringBuilder stringBuilder_notpass_runstairs = new StringBuilder(); //直跑楼梯，不符合的楼梯
            //StringBuilder stringBuilder_later = new StringBuilder(); //直跑楼梯，不符合的楼梯
            StringBuilder stringBuilder_result = new StringBuilder();
            bool H00099 = true;

            //切换到二维平面
            switchto2D(uiDocument, document);


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

            #endregion

            #region 步骤1.获得每块墙的坐标、每个楼梯的坐标，将坐标存入字典中
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
            
            //寻找楼梯的坐标
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
            #endregion


            //========== 遍历每个楼梯 ==========
            foreach (KeyValuePair<Element, Dictionary<string, double>> item_stairs in StairsBoundingXYZ_Dict)
            {
                Stairs stairs = item_stairs.Key as Stairs;
                #region 步骤2.判断楼梯中是否存在墙 
                //========== 1、判断楼梯中是否存在墙 ==========
                bool exist_wall = false;
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
                foreach (KeyValuePair<Element, Dictionary<string, double>> wall in WallBoundingXYZ_Dict)
                {
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
                #endregion

                #region 步骤3.判断楼梯中是否存在栏杆，并获得栏杆
                //========== 2、判断楼梯中是否存在栏杆，并获得栏杆 ==========
                bool exist_stairsroliings = false;
                try
                {
                    List<ElementId> StairsRoliingsId_List = getStairsRailings(document, stairs, StairsRailingCollector);
                    List<Element> StairsRoliings_List = GetItemByIds(StairsRoliingsId_List, StairsRailingCollector);
                    int num_stairsroliings = StairsRoliingsId_List.Count;
                    if (StairsRoliingsId_List.Count > 0)
                    {
                        exist_stairsroliings = true;
                    }
                }catch (Exception e) { }
                #endregion

                #region 步骤4.获得楼梯所拥有的梯段、平台
                //========== 3、获得楼梯所拥有的梯段、平台 ==========
                List<ElementId> StairsRunsId_List = stairs.GetStairsRuns().ToList();
                List<Element> StairsRuns_List = GetItemByIds(StairsRunsId_List, StairsRunsCollector);
                int num_stairsruns = StairsRunsId_List.Count;

                List<ElementId> StairsLandingsId_List = stairs.GetStairsLandings().ToList();
                List<Element> StairsLandings_List = GetItemByIds(StairsLandingsId_List, StairsLandingsCollector);
                int num_stairslandings = StairsLandingsId_List.Count;
                #endregion

                
                //========== 4、遍历梯段和平台 ==========
                foreach (Element item_stairsruns in StairsRuns_List)
                {
                    #region 步骤5.获得楼梯梯段的走向
                    StairsRun stairsruns = item_stairsruns as StairsRun;
                    double X_direction = 999;
                    double Y_direction = 999;
                    //获得梯段的X、Y走向
                    //CurveLoop curveLoop = stairsruns.GetStairsPath();
                    foreach (Curve curve in stairsruns.GetStairsPath())
                    {
                        if (curve is Line line) //判断Curve对象是否为Line对象
                        {
                            line = (Line)curve;
                            X_direction = Get_Right_Number(line.Direction.X);
                            Y_direction = Get_Right_Number(line.Direction.Y);
                        }
                    }
                    #endregion

                    //获得梯段的宽度
                    double StairsRunsWidth = Math.Round(Utils.FootToMeter(stairsruns.LookupParameter("实际梯段宽度").AsDouble()), 2);
                    //遍历楼梯的所有台面
                    foreach (Element item_stairsLanding in StairsLandings_List)
                    {
                        #region 步骤6.根据楼梯走向获得楼梯平台的净宽
                        StairsLanding stairslandings = item_stairsLanding as StairsLanding;
                        double StairsLandingsWidth;  //平台宽度
                        BoundingBoxXYZ boundingBoxXYZ = stairslandings.get_BoundingBox(document.ActiveView);
                        XYZ max = boundingBoxXYZ.Max;
                        XYZ min = boundingBoxXYZ.Min;
                        double StairsLandingsMinX = Get_Right_Number(min.X);
                        double StairsLandingsMaxX = Get_Right_Number(max.X);
                        double StairsLandingsMinY = Get_Right_Number(min.Y);
                        double StairsLandingsMaxY = Get_Right_Number(max.Y);
                        if (X_direction == 1)
                        {
                            StairsLandingsWidth = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinX - StairsLandingsMaxX)), 2);
                        }
                        else 
                        {
                            StairsLandingsWidth = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinY - StairsLandingsMaxY)), 2);
                        }
                        #endregion

                        #region 步骤7.根据楼梯走向判断连接的梯段和平台数量，检测楼梯中的平台净宽是否符合要求
                        //========== 5、进行判断 ==========
                        if (num_stairsruns == 1 && num_stairslandings == 1) //特殊情况：直跑楼梯
                        {
                            if (StairsLandingsWidth < 0.9)
                            {
                                stringBuilder_notpass_runstairs.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < 0.9：{stairsruns.Id}，直跑楼梯，平台净宽 < 0.9m");
                                H00099 = false;
                            }
                            else if (StairsLandingsWidth < StairsRunsWidth)
                            {
                                //将不符合的整合到下面中
                                //stringBuilder_notpass_basic.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < {StairsRunsWidth}：{stairsruns.Id}");
                                stringBuilder_notpass_runstairs.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < {StairsRunsWidth}：{stairsruns.Id}，直跑楼梯，平台净宽 < 梯段净宽");
                                H00099 = false;
                            }
                            else
                            {
                                stringBuilder_pass.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} >= {StairsRunsWidth}：{stairsruns.Id}，直跑楼梯");
                            }
                        }
                        else 
                        {
                            if (exist_wall)
                            {
                                if (StairsLandingsWidth < 1.3) //先判断是否符合基本要求
                                {
                                    stringBuilder_notpass_exist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < 1.3：{stairsruns.Id}，转角存在墙，平台净宽 < 1.3m");
                                    H00099 = false;
                                }
                                else //如果平台宽度符合要求再判断平台宽度是否大于等于梯段宽度
                                {
                                    if (StairsLandingsWidth < StairsRunsWidth)
                                    {
                                        //将不符合的整合到下面中
                                        //stringBuilder_notpass_basic.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < {StairsRunsWidth}：{stairsruns.Id}");
                                        stringBuilder_notpass_exist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < {StairsRunsWidth}：{stairsruns.Id}，转角存在墙，平台净宽 < 梯段净宽");
                                        H00099 = false;
                                    }
                                    else
                                    {
                                        stringBuilder_pass.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} >= {StairsRunsWidth}：{stairsruns.Id}，转角存在墙");
                                    }
                                }
                            }
                            else if (exist_stairsroliings)
                            {
                                if (StairsLandingsWidth < 1.2) //先判断是否符合基本要求
                                {
                                    stringBuilder_notpass_notexist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < 1.2：{stairsruns.Id}，转角存在栏杆，平台净宽 < 1.2m");
                                    H00099 = false;
                                }
                                else //如果平台宽度符合要求再判断平台宽度是否大于等于梯段宽度
                                {
                                    if (StairsLandingsWidth < StairsRunsWidth)
                                    {
                                        //将不符合的整合到下面中
                                        //stringBuilder_notpass_basic.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < {StairsRunsWidth}：{stairsruns.Id}");
                                        stringBuilder_notpass_notexist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < {StairsRunsWidth}：{stairsruns.Id}，转角存在栏杆，平台净宽 < 梯段净宽");
                                        H00099 = false;
                                    }
                                    else
                                    {
                                        stringBuilder_pass.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} >= {StairsRunsWidth}：{stairsruns.Id}，转角存在栏杆");
                                    }
                                }
                            }
                            else //如果既不存在墙也不存在栏杆
                            {
                                if (StairsLandingsWidth < 1.2) //先判断是否符合基本要求
                                {
                                    stringBuilder_notpass_notexist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < 1.2：{stairsruns.Id}，转角无实体，平台净宽 < 1.2m");
                                    H00099 = false;
                                }
                                else //如果平台宽度符合要求再判断平台宽度是否大于等于梯段宽度
                                {
                                    if (StairsLandingsWidth < StairsRunsWidth)
                                    {
                                        //将不符合的整合到下面中
                                        //stringBuilder_notpass_basic.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < {StairsRunsWidth}：{stairsruns.Id}");
                                        stringBuilder_notpass_notexist_element.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} < {StairsRunsWidth}：{stairsruns.Id}，转角无实体，平台净宽 < 梯段净宽");
                                        H00099 = false;
                                    }
                                    else
                                    {
                                        stringBuilder_pass.AppendLine($"{stairs.Id}：{stairslandings.Id}：{StairsLandingsWidth} >= {StairsRunsWidth}：{stairsruns.Id}，转角无实体");
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                }
            }

            #region 输出结果
            //stringBuilder.AppendLine("平台净宽小于梯段宽度，平台宽度不符合要求的有：\n" + stringBuilder_notpass_basic);
            stringBuilder.AppendLine("梯段改变方向，无实体端，平台宽度不符合要求的有：\n" + stringBuilder_notpass_notexist_element);
            stringBuilder.AppendLine("梯段改变方向，有实体端，平台宽度不符合要求的有：\n" + stringBuilder_notpass_exist_element);
            stringBuilder.AppendLine("直跑楼梯，平台宽度不符合要求的有：\n" + stringBuilder_notpass_runstairs);
            //stringBuilder.AppendLine("还未检测到的有：\n" + stringBuilder_later);
            //stringBuilder.AppendLine("符合的楼梯：\n" + stringBuilder_pass);




            if (H00099 == true)
            {
                stringBuilder_result.AppendLine("符合5.3.5民用建筑通用规范 GB55031-2022");
            }
            else
            {
                stringBuilder_result.AppendLine("不符合5.3.5民用建筑通用规范 GB55031-2022");
            }
            stringBuilder_result.AppendLine("————————————————————————————————");
            stringBuilder_result.AppendLine("输出格式说明：{楼梯Id}：{平台Id}：平台净宽 < 梯段宽度：{梯段Id}，{备注}");
            stringBuilder_result.AppendLine("————————————————————————————————");

            stringBuilder = stringBuilder_result.AppendLine(stringBuilder.ToString());

            //切换到三维平面
            switchto3D(uiDocument, document);


            Utils.PrintLog(stringBuilder.ToString(), "H00099", document);
            TaskDialog.Show("H00099强条检测", stringBuilder.ToString());

            #endregion

            return Result.Succeeded;
            
        }
    }
}
