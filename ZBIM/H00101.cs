using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ZBIMUtils;

namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    public class H00101 : IExternalCommand
    {
        private IEnumerable<Element> railings;

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

        public static IEnumerable<int> FindIndexList(string text, string word)
        {
            if (text.Length < 1 || word.Length < 1)
            {
                yield break;
            }
            int index = 0 - word.Length;
            while ((index = text.IndexOf(word, index + word.Length)) > -1)
            {
                yield return (index);
            }
            yield break;
        }
        public static int FindIndex(string text, string word)
        {
            if (text.Length < 1 || word.Length < 1)
            {
                return -1;
            }
            int index = 0 - word.Length;
            if ((index = text.IndexOf(word, index + word.Length)) > -1)
            {
                return index;
            }
            return -1;
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



        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            Document document = commandData.Application.ActiveUIDocument.Document;
            StringBuilder stringBuilder = new StringBuilder();
            StringBuilder stringBuilder_pass = new StringBuilder();
            StringBuilder stringBuilder_null = new StringBuilder();
            StringBuilder stringBuilder_notpass = new StringBuilder();
            StringBuilder stringBuilder_result = new StringBuilder();
            bool H00101 = true;

            //切换到二维平面
            switchto2D(uiDocument, document);


            //过滤：栏杆
            ElementCategoryFilter StairsRailingFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRailing);
            //过滤：楼梯
            //ElementCategoryFilter StairsFilter = new ElementCategoryFilter(BuiltInCategory.OST_Stairs);
            //过滤：平面
            //ElementCategoryFilter StairsLandingsFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsLandings);
            //过滤：梯段
            //ElementCategoryFilter StairsRunsFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRuns);
            //过滤：墙体
            ElementCategoryFilter WallsFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);


            //收集器：栏杆
            FilteredElementCollector StairsRailingCollector = new FilteredElementCollector(document).WherePasses(StairsRailingFilter);
            //FilteredElementCollector StairsRailingCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StairsRailing);
            //收集器：楼梯
            //FilteredElementCollector StairsCollector = new FilteredElementCollector(document).WherePasses(StairsFilter);
            //FilteredElementCollector StairsCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Stairs);
            //收集器：平面
            //FilteredElementCollector StairsLandingsCollector = new FilteredElementCollector(document).WherePasses(StairsLandingsFilter);
            //FilteredElementCollector StairsLandingsCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StairsLandings);
            //收集器：梯段
            //FilteredElementCollector StairsRunsCollector = new FilteredElementCollector(document).WherePasses(StairsRunsFilter);
            //收集器：墙体
            FilteredElementCollector WallsCollector = new FilteredElementCollector(document).WherePasses(WallsFilter);
            //FilteredElementCollector WallsCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Walls);



            
            //==========获得墙体Walls的BoundingBox==========
            //根据每块墙的BoundingXYZ坐标判断
            //键：墙的id，值：坐标信息，例如：{Xmin:0,Xmax:1,Ymin:2,Ymax:3,...}
            Dictionary<Element, Dictionary<string, double>> WallBoundingXYZ_Dict = new Dictionary<Element, Dictionary<string, double>>();
            double WallMinX;
            double WallMaxX;
            double WallMinY;
            double WallMaxY;
            double WallMinZ;
            double WallMaxZ;
            int Wall_findNum = 0;
            int Wall_notfindNum = 0;
            
            foreach (var wall in WallsCollector)
            {
                //if (wall.Id.ToString() != "3739118" || wall.Id.ToString() != "3468983")
                if (wall.Id.ToString() != "010101")
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
                        WallMinX = Math.Round(Get_Right_Number(min.X), 2);
                        WallMaxX = Math.Round(Get_Right_Number(max.X), 2);
                        WallMinY = Math.Round(Get_Right_Number(min.Y), 2);
                        WallMaxY = Math.Round(Get_Right_Number(max.Y), 2);
                        WallMinZ = Math.Round(Get_Right_Number(min.Z), 2);
                        WallMaxZ = Math.Round(Get_Right_Number(max.Z), 2);
                        WallBoundingXYZ_Dict[wall].Add("WallMinX", WallMinX);
                        WallBoundingXYZ_Dict[wall].Add("WallMaxX", WallMaxX);
                        WallBoundingXYZ_Dict[wall].Add("WallMinY", WallMinY);
                        WallBoundingXYZ_Dict[wall].Add("WallMaxY", WallMaxY);
                        WallBoundingXYZ_Dict[wall].Add("WallMinZ", WallMinZ);
                        WallBoundingXYZ_Dict[wall].Add("WallMaxZ", WallMaxZ);
                        //stringBuilder.AppendLine($"墙体名：{wall.Name}，{wall.Id}，X：({WallMinX},{WallMaxX})，Y：({WallMinY},{WallMaxY})，Z：({WallMinZ},{WallMaxZ})");
                        Wall_findNum++;

                    }
                    catch (Exception)
                    {
                        Wall_notfindNum++;
                        //stringBuilder.AppendLine("找不到的墙体：" + wall.Id.ToString() + "，" + wall.Name.ToString());
                        WallBoundingXYZ_Dict.Remove(wall);
                    }
                }
                
            }
            //stringBuilder.AppendLine($"找得到的墙体数量：{Wall_findNum}，找不到的墙体数量：{Wall_notfindNum}");



            //==========获得栏杆Railing的BoundingBox==========
            List<Element> filterd_railings = new List<Element>();
            //先过滤，去除“玻璃”栏杆
            foreach (Element item_railing in StairsRailingCollector)
            {
                try
                {
                    Railing railing = item_railing as Railing;
                    if (!(railing.Name.Contains("玻璃") || railing.GetPath().Count() <= 1))
                    {
                        filterd_railings.Add(item_railing);
                    }
                }
                catch { }
            }
            Dictionary<Element, Dictionary<string, double>> RailingBoundingXYZ_Dict = new Dictionary<Element, Dictionary<string, double>>();
            double RailingMinX;
            double RailingMaxX;
            double RailingMinY;
            double RailingMaxY;
            double RailingMinZ;
            double RailingMaxZ;
            double RailingXLength;
            double RailingYLength;
            double RailingZLength;
            int find_Railing = 0;
            int notfind_Railing = 0;
            
            foreach (Element item_railing in filterd_railings)
            {
                try
                {
                    Railing railing = item_railing as Railing;
                    if (!RailingBoundingXYZ_Dict.ContainsKey(railing))
                    {
                        RailingBoundingXYZ_Dict.Add(railing, new Dictionary<string, double>());
                    }
                    try
                    {
                        BoundingBoxXYZ boundingBoxXYZ = railing.get_BoundingBox(document.ActiveView);
                        XYZ max = boundingBoxXYZ.Max;
                        XYZ min = boundingBoxXYZ.Min;
                        RailingMinX = Math.Round(Get_Right_Number(min.X), 2);
                        RailingMaxX = Math.Round(Get_Right_Number(max.X), 2);
                        RailingMinY = Math.Round(Get_Right_Number(min.Y), 2);
                        RailingMaxY = Math.Round(Get_Right_Number(max.Y), 2);
                        RailingMinZ = Math.Round(Get_Right_Number(min.Z), 2);
                        RailingMaxZ = Math.Round(Get_Right_Number(max.Z), 2);
                        RailingXLength = Math.Round(Utils.FootToMeter(Math.Abs(RailingMinX - RailingMaxX)), 2);
                        RailingYLength = Math.Round(Utils.FootToMeter(Math.Abs(RailingMinY - RailingMaxY)), 2);
                        RailingZLength = Math.Round(Utils.FootToMeter(Math.Abs(RailingMinZ - RailingMaxZ)), 2);
                        RailingBoundingXYZ_Dict[railing].Add("RailingMinX", RailingMinX);
                        RailingBoundingXYZ_Dict[railing].Add("RailingMaxX", RailingMaxX);
                        RailingBoundingXYZ_Dict[railing].Add("RailingMinY", RailingMinY);
                        RailingBoundingXYZ_Dict[railing].Add("RailingMaxY", RailingMaxY);
                        RailingBoundingXYZ_Dict[railing].Add("RailingMinZ", RailingMinZ);
                        RailingBoundingXYZ_Dict[railing].Add("RailingMaxZ", RailingMaxZ);
                        RailingBoundingXYZ_Dict[railing].Add("RailingXLength", RailingXLength);
                        RailingBoundingXYZ_Dict[railing].Add("RailingYLength", RailingYLength);
                        RailingBoundingXYZ_Dict[railing].Add("RailingZLength", RailingZLength);
                        //stringBuilder.AppendLine($"栏杆名：{railing.Name}，{railing.Id}，X：({RailingMinX},{RailingMaxX})，Y：({RailingMinY},{RailingMaxY})，Z：({RailingMinZ},{RailingMaxZ})");
                        find_Railing++;
                    }
                    catch (Exception)
                    {
                        notfind_Railing++;
                        //stringBuilder.AppendLine("找不到的栏杆：" + railing.Id.ToString() + "，" + railing.Name.ToString());
                        RailingBoundingXYZ_Dict.Remove(railing);
                    }
                }
                catch { }
            }
            //stringBuilder.AppendLine($"找得到的栏杆数量：{find_Railing}，找不到的栏杆数量：{notfind_Railing}");
            


            /*
            foreach (Element item_railing in StairsRailingCollector)
            {
                double X_direction = 999;
                double Y_direction = 999;
                double Z_direction = 999;
                double X_origin = 0;
                double Y_origin = 0;
                double Z_origin = 0;
                int line_num = 0;
                Railing railing = item_railing as Railing;
                try
                {
                    foreach (Curve curve in railing.GetPath())
                    {
                        if (curve is Line line) //判断Curve对象是否为Line对象
                        {
                            line = (Line)curve;
                            X_direction = Get_Right_Number(line.Direction.X);
                            Y_direction = Get_Right_Number(line.Direction.Y);
                            Z_direction = Get_Right_Number(line.Direction.Z);
                            X_origin = Math.Round(line.Origin.X, 2);
                            Y_origin = Math.Round(line.Origin.Y, 2);
                            Z_origin = Math.Round(line.Origin.Z, 2);
                            line_num++;
                            stringBuilder.AppendLine($"{line_num}#：{X_direction}，{Y_direction}，{Z_direction}，({X_origin}，{Y_origin}，{Z_origin})");
                            
                        }
                        else
                        {
                            stringBuilder.Append($"不是线条：{curve}，");
                        }
                    }
                    stringBuilder.Append("========================================================================\n");

                }
                catch (Exception)
                {
                    stringBuilder.Append($"异常栏杆：{item_railing.Id}，{item_railing.Name}，也许是不存在？");
                    stringBuilder.AppendLine("\n========================================================================");
                }
            }
            */


            //==========开始遍历==========

            foreach (KeyValuePair<Element, Dictionary<string, double>> item_stairsRoliings in RailingBoundingXYZ_Dict)
            {
                
                Railing railing = item_stairsRoliings.Key as Railing;


                //获得栏杆的BoundingBox
                RailingMinX = item_stairsRoliings.Value["RailingMinX"];
                RailingMaxX = item_stairsRoliings.Value["RailingMaxX"];
                RailingMinY = item_stairsRoliings.Value["RailingMinY"];
                RailingMaxY = item_stairsRoliings.Value["RailingMaxY"];
                RailingMinZ = item_stairsRoliings.Value["RailingMinZ"];
                RailingMaxZ = item_stairsRoliings.Value["RailingMaxZ"];
                RailingXLength = item_stairsRoliings.Value["RailingXLength"];
                RailingYLength = item_stairsRoliings.Value["RailingYLength"];
                RailingZLength = item_stairsRoliings.Value["RailingZLength"];



                //找栏杆的每条线的方向
                double X_direction = 999;
                double Y_direction = 999;
                double Z_direction = 999;
                double X_origin = 0;
                double Y_origin = 0;
                double Z_origin = 0;
                double X_end = 0;
                double Y_end = 0;
                double Z_end = 0;
                double line_length = 0;
                double X_min = 0;
                double X_max = 0;
                double Y_min = 0;
                double Y_max = 0;
                double Z_min = 0;
                double Z_max = 0;
                int line_num = 0;
                try
                {
                    //stringBuilder.AppendLine($"{railing.Id}，{railing.Name}");
                    foreach (Curve curve in railing.GetPath())
                    {
                        if (curve is Line line) //判断Curve对象是否为Line对象
                        {
                            line = (Line)curve;
                            X_direction = Get_Right_Number(line.Direction.X);
                            Y_direction = Get_Right_Number(line.Direction.Y);
                            line_length = Math.Round(line.Length, 2);
                            X_origin = Math.Round(line.Origin.X, 2);
                            Y_origin = Math.Round(line.Origin.Y, 2);
                            Z_origin = Math.Round(line.Origin.Z, 2);

                            //将Z_origin修正到栏杆的Z轴高度区间
                            if (!(Z_origin > RailingMinZ && Z_origin < RailingMaxZ))
                            {
                                Z_origin = RailingMinZ;
                            }

                            X_end = X_origin + X_direction * line_length;
                            Y_end = Y_origin + Y_direction * line_length;
                            Z_end = Math.Round(Z_origin + RailingZLength, 2);


                            //线段的最大、最小XY值
                            X_min = Math.Min(X_origin, X_end);
                            X_max = Math.Max(X_origin, X_end);
                            Y_min = Math.Min(Y_origin, Y_end);
                            Y_max = Math.Max(Y_origin, Y_end);
                            Z_min = Z_origin;
                            Z_max = Z_end;

                            

                            line_num++;
                            //stringBuilder.AppendLine($"{railing.Id}，{railing.Name}，{line_num}#，{X_direction}，{Y_direction}，({X_min}，{X_max})，({Y_min}，{Y_max})，({Z_min}，{Z_max})，{line_length}");

                            /*
                            //遍历所有墙，计算线与墙垂直方向的距离
                            //double up_distance = 0;
                            //double bottom_distance = 0;
                            //double left_distance = 0;
                            //double right_distance = 0;
                            //double z_bia_up_distance = 0;    // 墙与栏杆的顶面差距距离
                            //double z_bia_bottom_distance = 0;// 墙与栏杆的底面差距距离
                            */

                            if (line_length > 0.2)
                            {
                                foreach (KeyValuePair<Element, Dictionary<string, double>> item_wall in WallBoundingXYZ_Dict)
                                {
                                    Wall wall = item_wall.Key as Wall;
                                    //获得墙的BoundingBox
                                    WallMinX = item_wall.Value["WallMinX"];
                                    WallMaxX = item_wall.Value["WallMaxX"];
                                    WallMinY = item_wall.Value["WallMinY"];
                                    WallMaxY = item_wall.Value["WallMaxY"];
                                    WallMinZ = item_wall.Value["WallMinZ"];
                                    WallMaxZ = item_wall.Value["WallMaxZ"];

                                    /*
                                    //判断墙在栏杆的哪个相对方向（以栏杆为中心）
                                    //up_distance = Math.Round(Utils.FootToMeter(WallMinY - RailingMaxY), 2);      //北
                                    //bottom_distance = Math.Round(Utils.FootToMeter(RailingMinY - WallMaxY), 2);  //南
                                    //left_distance = Math.Round(Utils.FootToMeter(RailingMinX - WallMaxX), 2);    //西
                                    //right_distance = Math.Round(Utils.FootToMeter(WallMinX - RailingMaxX), 2);   //东
                                    //z_bia_up_distance = Math.Round(Utils.FootToMeter(Math.Abs(RailingMinZ - WallMinZ)), 2);      //Z轴的顶部偏移量
                                    //z_bia_bottom_distance = Math.Round(Utils.FootToMeter(Math.Abs(WallMaxZ - RailingMaxZ)), 2);   //Z轴的底部偏移量
                                    */

                                    double line2wall_distance = 0;  //线段到墙的垂直距离
                                    double line_x_min_pre = 0;
                                    double line_x_max_pre = 0;
                                    double line_y_min_pre = 0;
                                    double line_y_max_pre = 0;
                                    double line_x_min_kuozhan = 0;
                                    double line_x_max_kuozhan = 0;
                                    double line_y_min_kuozhan = 0;
                                    double line_y_max_kuozhan = 0;

                                    double temp_distance1 = 0;
                                    double temp_distance2 = 0;
                                    double temp_distance3 = 0;
                                    double temp_distance4 = 0;


                                    //扩展前计算是否碰撞
                                    if (Math.Abs(X_direction) == 1 && Math.Abs(Y_direction) == 0)
                                    {
                                        //线段为X方向，则向Y方向扩展，比较Y方向的距离
                                        line_x_min_pre = X_min - 0.0005;
                                        line_x_max_pre = X_max - 0.0005;
                                        line_y_min_pre = Y_origin;
                                        line_y_max_pre = Y_origin;

                                        line_x_min_kuozhan = X_min - 0.0005;
                                        line_x_max_kuozhan = X_max - 0.0005;
                                        line_y_min_kuozhan = Y_origin - Utils.MeterToFoot(0.04);
                                        line_y_max_kuozhan = Y_origin + Utils.MeterToFoot(0.04);

                                        //计算得到线段到墙的距离
                                        temp_distance1 = Math.Abs(line_y_min_pre - WallMinY);
                                        temp_distance2 = Math.Abs(line_y_min_pre - WallMaxY);
                                        temp_distance3 = Math.Abs(line_y_max_pre - WallMinY);
                                        temp_distance4 = Math.Abs(line_y_max_pre - WallMaxY);
                                        line2wall_distance = Math.Min(temp_distance1, temp_distance2);
                                        line2wall_distance = Math.Min(line2wall_distance, temp_distance3);
                                        line2wall_distance = Math.Min(line2wall_distance, temp_distance4);

                                    }
                                    else if (Math.Abs(X_direction) == 0 && Math.Abs(Y_direction) == 1)
                                    {
                                        //线段为Y方向，则向X方向扩展（相隔很近的存在约0.0005的空隙）
                                        line_x_min_pre = X_origin;
                                        line_x_max_pre = X_origin;
                                        line_y_min_pre = Y_min - 0.0005;
                                        line_y_max_pre = Y_max - 0.0005;

                                        line_x_min_kuozhan = X_origin - Utils.MeterToFoot(0.04);
                                        line_x_max_kuozhan = X_origin + Utils.MeterToFoot(0.04);
                                        line_y_min_kuozhan = Y_min - 0.0005;
                                        line_y_max_kuozhan = Y_max - 0.0005;

                                        //计算得到线段到墙的距离
                                        temp_distance1 = Math.Abs(line_x_min_pre - WallMinX);
                                        temp_distance2 = Math.Abs(line_x_min_pre - WallMaxX);
                                        temp_distance3 = Math.Abs(line_x_max_pre - WallMinX);
                                        temp_distance4 = Math.Abs(line_x_max_pre - WallMaxX);
                                        line2wall_distance = Math.Min(temp_distance1, temp_distance2);
                                        line2wall_distance = Math.Min(line2wall_distance, temp_distance3);
                                        line2wall_distance = Math.Min(line2wall_distance, temp_distance4);
                                    }
                                    if ((Z_max > WallMinZ && Z_max < WallMaxZ && Z_min < WallMinZ) || (Z_min < WallMaxZ && Z_min > WallMinZ) || (Z_max > WallMaxZ && Z_min < WallMinZ) || (Z_max < WallMaxZ && Z_min > WallMinZ))  //必要条件：墙与线条在差不多的高度
                                    {
                                        //如果线条没扩展前线条和墙没有碰撞，扩展后才碰撞的才认为不符合
                                        if (line_x_min_pre < WallMaxX && line_x_max_pre > WallMinX && line_y_min_pre < WallMaxY && line_y_max_pre > WallMinY)
                                        //if (item1_x_min < item2_x_max && item1_x_max > item2_x_min && item1_y_min < item2_y_max && item1_y_max > item2_y_min)
                                        {
                                            //这是扩展前就和线条碰撞的，不能再进行判定了
                                        }
                                        else
                                        {
                                            // 线条扩展后与墙不存在重叠，说明距离符合要求
                                            //stringBuilder_notpass.AppendLine($"{railing.Id}，{railing.Name}，{X_direction}，{Y_direction}，({X_min}，{X_max})，({Y_min}，{Y_max})，({Z_min}，{Z_max})，{wall.Id}，({WallMinX}，{WallMaxX})，({WallMinY}，{WallMaxY})，({WallMinZ}，{WallMaxZ})，距离：{line2wall_distance}");
                                            //stringBuilder_pass.AppendLine($"{railing.Id}，{wall.Id}");

                                            //若扩展后存在重合，则才认为是不符合的
                                            if (line_x_min_kuozhan < WallMaxX && line_x_max_kuozhan > WallMinX && line_y_min_kuozhan < WallMaxY && line_y_max_kuozhan > WallMinY)
                                            {
                                                stringBuilder_notpass.AppendLine($"{railing.Id}，{line_length}，{wall.Id}，{Math.Round(Utils.FootToMeter(line2wall_distance) * 1000, 4)}");
                                                H00101 = false;
                                                // 线条向外扩展0.04mm，若与墙存在重叠，则不符合
                                                //if (line2wall_distance > 0)
                                                //{
                                                //stringBuilder_notpass.AppendLine($"{railing.Id}，{railing.Name}，{X_direction}，{Y_direction}，({X_min}，{X_max})，({Y_min}，{Y_max})，({Z_min}，{Z_max})，{wall.Id}，({WallMinX}，{WallMaxX})，({WallMinY}，{WallMaxY})，({WallMinZ}，{WallMaxZ})，距离：{line2wall_distance}");

                                                //}
                                            }

                                        }
                                    }



                                    /*
                                    if ((Z_max > WallMinZ && Z_max < WallMaxZ && Z_min < WallMinZ) || (Z_min < WallMaxZ && Z_min > WallMinZ) || (Z_max > WallMaxZ && Z_min < WallMinZ) || (Z_max < WallMaxZ && Z_min > WallMinZ))  //墙与栏杆线条在差不多的高度
                                    {
                                        if (Math.Abs(X_direction) == 1 && Math.Abs(Y_direction) == 0)
                                        {
                                            if (Y_origin > WallMaxY) //墙在线段北侧，且在栏杆的正上方
                                            {
                                                if ((RailingMinX < WallMaxX && RailingMaxX > WallMaxX) || (RailingMaxX > WallMinX && RailingMinX < WallMinX) || (RailingMinX < WallMinX && RailingMaxX > WallMaxX) || (RailingMinX > WallMinX && RailingMaxX < WallMaxX))
                                                {
                                                    stringBuilder.AppendLine($"X北，{X_direction}，{Y_direction}，({X_min}，{X_max})，({Y_min}，{Y_max})，({Z_min}，{Z_max})，{wall.Id}，({WallMinX}，{WallMaxX})，({WallMinY}，{WallMaxY})，({WallMinZ}，{WallMaxZ})");
                                                }
                                            }
                                            else if (Y_origin < WallMinY)  //墙在线段南侧
                                            {
                                                if (((RailingMinX < WallMaxX && RailingMaxX > WallMaxX) || (RailingMaxX > WallMinX && RailingMinX < WallMinX) || (RailingMinX < WallMinX && RailingMaxX > WallMaxX) || (RailingMinX > WallMinX && RailingMaxX < WallMaxX)))
                                                {
                                                    stringBuilder.AppendLine($"X南，{X_direction}，{Y_direction}，({X_min}，{X_max})，({Y_min}，{Y_max})，({Z_min}，{Z_max})，{wall.Id}，({WallMinX}，{WallMaxX})，({WallMinY}，{WallMaxY})，({WallMinZ}，{WallMaxZ})");
                                                }
                                            }
                                            else
                                            { 
                                                stringBuilder.AppendLine($"X，不知道是北侧还是南侧，{X_direction}，{Y_direction}，({X_min}，{X_max})，({Y_min}，{Y_max})，{wall.Id}，({Z_min}，{Z_max})，({WallMinX}，{WallMaxX})，({WallMinY}，{WallMaxY})，({WallMinZ}，{WallMaxZ})");
                                            }
                                        }
                                        else if (Math.Abs(X_direction) == 0 && Math.Abs(Y_direction) == 1)
                                        {
                                            if (X_origin > WallMaxX) //墙在线段西侧
                                            {
                                                if ((RailingMinY < WallMaxY && RailingMaxY > WallMaxY) || (RailingMaxY > WallMinY && RailingMinY < WallMinY) || (RailingMinY < WallMinY && RailingMaxY > WallMaxY) || (RailingMinY > WallMinY && RailingMaxY < WallMaxY))
                                                {
                                                    stringBuilder.AppendLine($"Y西，{X_direction}，{Y_direction}，({X_min}，{X_max})，({Y_min}，{Y_max})，({Z_min}，{Z_max})，{wall.Id}，({WallMinX}，{WallMaxX})，({WallMinY}，{WallMaxY})，({WallMinZ}，{WallMaxZ})");
                                                }
                                            }
                                            else if (X_origin < WallMinX) //墙在线段东侧
                                            {
                                                if ((RailingMinY < WallMaxY && RailingMaxY > WallMaxY) || (RailingMaxY > WallMinY && RailingMinY < WallMinY) || (RailingMinY < WallMinY && RailingMaxY > WallMaxY) || (RailingMinY > WallMinY && RailingMaxY < WallMaxY))
                                                {
                                                    stringBuilder.AppendLine($"Y东，{X_direction}，{Y_direction}，({X_min}，{X_max})，({Y_min}，{Y_max})，({Z_min}，{Z_max})，{wall.Id}，({WallMinX}，{WallMaxX})，({WallMinY}，{WallMaxY})，({WallMinZ}，{WallMaxZ})");
                                                }
                                            }
                                            else
                                            { 
                                                stringBuilder.AppendLine($"Y，不知道是西侧还是东侧，{X_direction}，{Y_direction}，({X_min}，{X_max})，({Y_min}，{Y_max})，{wall.Id}，({Z_min}，{Z_max})，({WallMinX}，{WallMaxX})，({WallMinY}，{WallMaxY})，({WallMinZ}，{WallMaxZ})");
                                            }
                                        }
                                        else
                                        { 
                                            stringBuilder.AppendLine($"未知方向，{X_direction}，{Y_direction}，({X_min}，{X_max})，({Y_min}，{Y_max})，({Z_min}，{Z_max})，{wall.Id}，({WallMinX}，{WallMaxX})，({WallMinY}，{WallMaxY})，({WallMinZ}，{WallMaxZ})");
                                        }


                                    }

                                    */
                                    /*
                                    if (!((WallMinZ > RailingMaxZ) || (RailingMinZ > WallMaxZ)))    // 墙首先必须跟栏杆的高度大约齐平
                                    {
                                        stringBuilder_notpass.AppendLine($"{railing.Id}，{line.GetHashCode()}，({X_direction}，{Y_direction})，{wall.Id}");
                                        if (Math.Abs(X_direction) == 1 && Math.Abs(Y_direction) == 0)
                                        {
                                            if (Y_origin > WallMaxY && X_min < WallMaxX && X_max > WallMinX) //墙在线段北侧，且正上方，非斜上方
                                            {
                                                line2wall_distance = Math.Round(Utils.FootToMeter(Math.Abs(WallMaxY - Y_origin)), 2);
                                                if (line2wall_distance < 0.04)
                                                {
                                                    stringBuilder_notpass.AppendLine($"{railing.Id}，{line.GetHashCode()}，({X_direction}，{Y_direction})，{wall.Id}，{line2wall_distance}m，北");
                                                }
                                            }
                                            else if (Y_origin < WallMinY && X_min < WallMaxX && X_max > WallMinX) //墙在线段南侧，且正下方，非斜下方
                                            {
                                                line2wall_distance = Math.Round(Utils.FootToMeter(Math.Abs(WallMinY - Y_origin)), 2);
                                                if (line2wall_distance < 0.04)
                                                    stringBuilder_notpass.AppendLine($"{railing.Id}，{line.GetHashCode()}，({X_direction}，{Y_direction})，{wall.Id}，{line2wall_distance}m，南");
                                            }
                                            else //线段的墙的中间
                                            {
                                                //stringBuilder_notpass.AppendLine($"{railing.Id}，{line.GetHashCode()}，({X_direction}，{Y_direction})，{wall.Id}，在墙中间");
                                            }
                                        }
                                        else if (Math.Abs(X_direction) == 0 && Math.Abs(Y_direction) == 1)
                                        {
                                            if (X_origin > WallMaxX && Y_min < WallMaxY && Y_max > WallMinY) //墙在线段西侧
                                            {
                                                line2wall_distance = Math.Round(Utils.FootToMeter(Math.Abs(WallMinX - X_origin)), 2);

                                                if (line2wall_distance < 0.04)
                                                    stringBuilder_notpass.AppendLine($"{railing.Id}，{line.GetHashCode()}，({X_direction}，{Y_direction})，{wall.Id}，{line2wall_distance}m，西");

                                            }
                                            else if (X_origin < WallMinX && Y_min < WallMaxY && Y_max > WallMinY) //墙在线段东侧
                                            {
                                                line2wall_distance = Math.Round(Utils.FootToMeter(Math.Abs(WallMaxX - X_origin)), 2);

                                                if (line2wall_distance < 0.04)
                                                    stringBuilder_notpass.AppendLine($"{railing.Id}，{line.GetHashCode()}，({X_direction}，{Y_direction})，{wall.Id}，{line2wall_distance}m，东");
                                            }
                                            else //线段的墙的中间
                                            {
                                                //stringBuilder_notpass.AppendLine($"{railing.Id}，{line.GetHashCode()}，({X_direction}，{Y_direction})，{wall.Id}，在墙中间");
                                            }
                                        }
                                    }
                                    */

                                }
                            }
                            
                            
                            

                        }
                        else
                        {
                            //stringBuilder.Append($"不是线条：{curve}，");
                        }
                    }
                    //stringBuilder.Append("========================================================================\n");

                }
                catch (Exception)
                {
                    //stringBuilder.Append($"异常栏杆：{railing.Id}，{railing.Name}，也许是不存在？");
                    //stringBuilder.AppendLine("\n========================================================================");
                }


            }




            stringBuilder.AppendLine("不符合要求的栏杆：（栏杆扶手到墙面距离小于40mm）\n" + stringBuilder_notpass);
            //stringBuilder.AppendLine("符合要求的栏杆：\n" + stringBuilder_pass);



            if (H00101 == true)
            {
                stringBuilder_result.AppendLine("符合5.3.3民用建筑通用规范 GB55031-2022");
            }
            else
            {
                stringBuilder_result.AppendLine("不符合5.3.3民用建筑通用规范 GB55031-2022");
            }

            stringBuilder_result.AppendLine("—————————————————");
            stringBuilder_result.AppendLine("输出格式说明：{栏杆Id}，{栏杆扶手的长度（米/m），{墙Id}，{栏杆扶手到墙面的距离（毫米/mm）");
            stringBuilder_result.AppendLine("—————————————————");

            stringBuilder = stringBuilder_result.AppendLine(stringBuilder.ToString());

            //切换到三维平面
            Utils.CloseCurrentView(uiDocument);
            switchto3D(uiDocument, document);




            Utils.PrintLog(stringBuilder.ToString(), "H00101", document);
            TaskDialog.Show("H00101强条检测", stringBuilder.ToString());
            return Result.Succeeded;

        }
    }
}
