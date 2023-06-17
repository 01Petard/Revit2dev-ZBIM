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



    public class FilterStairs : IExternalCommand
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
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            Document document = commandData.Application.ActiveUIDocument.Document;
            StringBuilder stringBuilder = new StringBuilder();
            StringBuilder stringBuilder_pass = new StringBuilder();
            StringBuilder stringBuilder_notpass = new StringBuilder();
            StringBuilder stringBuilder_result = new StringBuilder();
            bool H00099 = true;

            stringBuilder.AppendLine("==============================================");
            //==========找栏杆StairsRailing==========
            FilteredElementCollector StairsRailingCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StairsRailing);
            Dictionary<Element, Dictionary<string, double>> StairsRailingBoundingXYZ_Dict = new Dictionary<Element, Dictionary<string, double>>();
            double StairsRailingMinX = 0;
            double StairsRailingMaxX = 0;
            double StairsRailingMinY = 0;
            double StairsRailingMaxY = 0;
            double StairsRailingMinZ = 0;
            double StairsRailingMaxZ = 0;
            double StairsRailingXLength = 0;
            double StairsRailingYLength = 0;
            double StairsRailingXY = 0;  // 0 表示X方向更长，1表示Y方向更长
            int find_StairsRailing = 0;
            int notfind_StairsRailing = 0;

            foreach (var stairsrailing in StairsRailingCollector)
            {
                if (!StairsRailingBoundingXYZ_Dict.ContainsKey(stairsrailing))
                {
                    StairsRailingBoundingXYZ_Dict.Add(stairsrailing, new Dictionary<string, double>());
                }
                try
                {
                    BoundingBoxXYZ boundingBoxXYZ = stairsrailing.get_BoundingBox(document.ActiveView);
                    XYZ max = boundingBoxXYZ.Max;
                    XYZ min = boundingBoxXYZ.Min;
                    StairsRailingMinX = Get_Right_Number(min.X);
                    StairsRailingMaxX = Get_Right_Number(max.X);
                    StairsRailingMinY = Get_Right_Number(min.X);
                    StairsRailingMaxY = Get_Right_Number(max.X);
                    StairsRailingMinZ = Get_Right_Number(min.X);
                    StairsRailingMaxZ = Get_Right_Number(max.X);
                    StairsRailingXLength = Math.Round(Utils.FootToMeter(Math.Abs(StairsRailingMinX - StairsRailingMaxX)), 2);
                    StairsRailingYLength = Math.Round(Utils.FootToMeter(Math.Abs(StairsRailingMinY - StairsRailingMaxY)), 2);
                    if (StairsRailingXLength > StairsRailingYLength)
                    {
                        StairsRailingXY = 0;  // 如果是X轴方向上的长条形，就赋值为0，平台净宽通过X轴计算
                    }
                    else {
                        StairsRailingXY = 1; // 如果是Y轴方向上的长条形，就赋值为1，平台净宽通过Y轴计算
                    }
                    StairsRailingBoundingXYZ_Dict[stairsrailing].Add("StairsRailingMinX", StairsRailingMinX);
                    StairsRailingBoundingXYZ_Dict[stairsrailing].Add("StairsRailingMaxX", StairsRailingMaxX);
                    StairsRailingBoundingXYZ_Dict[stairsrailing].Add("StairsRailingMinY", StairsRailingMinY);
                    StairsRailingBoundingXYZ_Dict[stairsrailing].Add("StairsRailingMaxY", StairsRailingMaxY);
                    StairsRailingBoundingXYZ_Dict[stairsrailing].Add("StairsRailingMinZ", StairsRailingMinZ);
                    StairsRailingBoundingXYZ_Dict[stairsrailing].Add("StairsRailingMaxZ", StairsRailingMaxZ);
                    StairsRailingBoundingXYZ_Dict[stairsrailing].Add("StairsRailingXLength", StairsRailingXLength);
                    StairsRailingBoundingXYZ_Dict[stairsrailing].Add("StairsRailingYLength", StairsRailingYLength);
                    StairsRailingBoundingXYZ_Dict[stairsrailing].Add("StairsRailingXY", StairsRailingXY);

                    //stringBuilder.AppendLine($"楼梯名：{stairs.Name}，{stairs.Id}，X：({StairsMinX},{StairsMaxX})，Y：({StairsMinY},{StairsMaxY})，Z：({StairsMinZ},{StairsMaxZ})");
                    find_StairsRailing++;

                }
                catch (Exception e)
                {
                    notfind_StairsRailing++;
                    //stringBuilder.AppendLine("找不到的栏杆：" + stairsrailing.Id.ToString() + "，" + stairsrailing.Name.ToString());
                    StairsRailingBoundingXYZ_Dict.Remove(stairsrailing);
                }
            }
            stringBuilder.AppendLine($"找得到的栏杆数量：{find_StairsRailing}，找不到的栏杆数量：{notfind_StairsRailing}");



            stringBuilder.AppendLine("==============================================");
            //==========找墙体Walls==========
            FilteredElementCollector WallsCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Walls);
            Dictionary<Element, Dictionary<string, double>> WallBoundingXYZ_Dict = new Dictionary<Element, Dictionary<string, double>>();
            double WallMinX = 0;
            double WallMaxX = 0;
            double WallMinY = 0;
            double WallMaxY = 0;
            double WallMinZ = 0;
            double WallMaxZ = 0;
            double WallXLength = 0;
            double WallYLength = 0;
            double WallXY = 0;  // 0 表示X方向更长，1表示Y方向更长
            int find_Wall = 0;
            int notfind_Wall = 0;

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

                    //stringBuilder.AppendLine($"楼梯名：{stairs.Name}，{stairs.Id}，X：({StairsMinX},{StairsMaxX})，Y：({StairsMinY},{StairsMaxY})，Z：({StairsMinZ},{StairsMaxZ})");
                    find_Wall++;

                }
                catch (Exception)
                {
                    notfind_Wall++;
                    //stringBuilder.AppendLine("找不到的墙体：" + wall.Id.ToString() + "，" + wall.Name.ToString());
                    WallBoundingXYZ_Dict.Remove(wall);
                }
            }
            stringBuilder.AppendLine($"找得到的墙体数量：{find_Wall}，找不到的墙体数量：{notfind_Wall}");



            stringBuilder.AppendLine("==============================================");
            //==========找楼梯Stairs==========
            //==========找楼梯对应的梯段和平面==========
            FilteredElementCollector StairsCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Stairs);
            ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_Stairs);

            FilteredElementCollector Stairs2collector = new FilteredElementCollector(document).WherePasses(elementCategoryFilter);
            foreach (Stairs stairs in StairsCollector)
            {
                ICollection<ElementId> StairsRun_List = stairs.GetStairsRuns();
                ICollection<ElementId> StairsLanding_List = stairs.GetStairsLandings();
                stringBuilder.Append($"楼梯：{stairs.Id},{stairs.Name}");
                foreach (var stairsrun in StairsRun_List)
                {
                    stringBuilder.Append($"，梯段：{stairsrun}");
                }

                foreach (var stairslanding in StairsLanding_List)
                {
                    stringBuilder.Append($"，平面：{stairslanding}");
                }

                stringBuilder.AppendLine("");
            }
            Dictionary<Stairs, Dictionary<string, double>> StairsBoundingXYZ_Dict = new Dictionary<Stairs, Dictionary<string, double>>();

            double StairsMinX = 0;
            double StairsMaxX = 0;
            double StairsMinY = 0;
            double StairsMaxY = 0;
            double StairsMinZ = 0;
            double StairsMaxZ = 0;
            double StairsXLength = 0;
            double StairsYLength = 0;

            double StairsWidth = 0;


            int find_Stairs = 0;
            int notfind_Stairs = 0;

            foreach (Stairs stairs in StairsCollector)
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

                    StairsBoundingXYZ_Dict[stairs].Add("StairsMinX", StairsMinX);
                    StairsBoundingXYZ_Dict[stairs].Add("StairsMaxX", StairsMaxX);
                    StairsBoundingXYZ_Dict[stairs].Add("StairsMinY", StairsMinY);
                    StairsBoundingXYZ_Dict[stairs].Add("StairsMaxY", StairsMaxY);
                    StairsBoundingXYZ_Dict[stairs].Add("StairsMinZ", StairsMinZ);
                    StairsBoundingXYZ_Dict[stairs].Add("StairsMaxZ", StairsMaxZ);
                    StairsBoundingXYZ_Dict[stairs].Add("StairsXLength", StairsXLength);
                    StairsBoundingXYZ_Dict[stairs].Add("StairsYLength", StairsYLength);

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
            stringBuilder.AppendLine($"找得到的楼梯数量：{find_Stairs}，找不到的楼梯数量：{notfind_Stairs}");

            stringBuilder.AppendLine("==============================================");
            //==========找楼梯平台StairsLandings==========
            FilteredElementCollector StairsLandingsCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StairsLandings);
            Dictionary<Element, Dictionary<string, double>> StairsLandingsBoundingXYZ_Dict = new Dictionary<Element, Dictionary<string, double>>();
            double StairsLandingsMinX = 0;
            double StairsLandingsMaxX = 0;
            double StairsLandingsMinY = 0;
            double StairsLandingsMaxY = 0;
            double StairsLandingsMinZ = 0;
            double StairsLandingsMaxZ = 0;
            double StairsLandingsXLength = 0;
            double StairsLandingsYLength = 0;

            double StairsLandingsWidth = 0;  // 平台净宽

            int find_StairsLandings = 0;
            int notfind_StairsLandings = 0;

            foreach (var stairslanding in StairsLandingsCollector)
            {
                if (!StairsLandingsBoundingXYZ_Dict.ContainsKey(stairslanding))
                {
                    StairsLandingsBoundingXYZ_Dict.Add(stairslanding, new Dictionary<string, double>());
                }
                try
                {
                    BoundingBoxXYZ boundingBoxXYZ = stairslanding.get_BoundingBox(document.ActiveView);
                    XYZ max = boundingBoxXYZ.Max;
                    XYZ min = boundingBoxXYZ.Min;
                    
                    StairsLandingsMinX = Get_Right_Number(min.X);
                    StairsLandingsMaxX = Get_Right_Number(max.X);
                    StairsLandingsMinY = Get_Right_Number(min.Y);
                    StairsLandingsMaxY = Get_Right_Number(max.Y);
                    StairsLandingsMinZ = Get_Right_Number(min.Z);
                    StairsLandingsMaxZ = Get_Right_Number(max.Z);
                    StairsLandingsXLength = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinX - StairsLandingsMaxX)), 2);
                    StairsLandingsYLength = Math.Round(Utils.FootToMeter(Math.Abs(StairsLandingsMinY - StairsLandingsMaxY)), 2);

                    StairsLandingsBoundingXYZ_Dict[stairslanding].Add("StairsLandingsMinX", StairsLandingsMinX);
                    StairsLandingsBoundingXYZ_Dict[stairslanding].Add("StairsLandingsMaxX", StairsLandingsMaxX);
                    StairsLandingsBoundingXYZ_Dict[stairslanding].Add("StairsLandingsMinY", StairsLandingsMinY);
                    StairsLandingsBoundingXYZ_Dict[stairslanding].Add("StairsLandingsMaxY", StairsLandingsMaxY);
                    StairsLandingsBoundingXYZ_Dict[stairslanding].Add("StairsLandingsMinZ", StairsLandingsMinZ);
                    StairsLandingsBoundingXYZ_Dict[stairslanding].Add("StairsLandingsMaxZ", StairsLandingsMaxZ);
                    StairsLandingsBoundingXYZ_Dict[stairslanding].Add("StairsLandingsXLength", StairsLandingsXLength);
                    StairsLandingsBoundingXYZ_Dict[stairslanding].Add("StairsLandingsYLength", StairsLandingsYLength);

                    stringBuilder.AppendLine($"楼梯平台：{stairslanding.Name}，{stairslanding.Id}，X：({StairsLandingsMinX},{StairsLandingsMaxX})，Y：({StairsLandingsMinY},{StairsLandingsMaxY})，Z：({StairsLandingsMinZ},{StairsLandingsMaxZ})");
                    find_StairsLandings++;

                }
                catch (Exception)
                {
                    notfind_StairsLandings++;
                    //stringBuilder.AppendLine("找不到的楼梯平台：" + stairslanding.Id.ToString() + "，" + stairslanding.Name.ToString());
                    StairsLandingsBoundingXYZ_Dict.Remove(stairslanding);
                }
            }
            stringBuilder.AppendLine($"找得到的楼梯平台数量：{find_StairsLandings}，找不到的楼梯平台数量：{notfind_StairsLandings}");


            stringBuilder.AppendLine("==============================================");











            /*
            //进行判断
            stringBuilder.AppendLine("==============================================");
            int stairs_landins_num; //每个楼梯拥有的平台数
            bool exist_wall;
            bool exist_stairsrailing;
            //楼梯的坐标
            foreach (KeyValuePair<Element, Dictionary<string, double>> stairs in StairsBoundingXYZ_Dict)
            {
                exist_wall = false;
                exist_stairsrailing = false;
                //stringBuilder.AppendLine(stairs.Value.GetType().ToString());
                stairs_landins_num = 0;
                StairsMinX = stairs.Value["StairsMinX"];
                StairsMaxX = stairs.Value["StairsMaxX"];
                StairsMinY = stairs.Value["StairsMinY"];
                StairsMaxY = stairs.Value["StairsMaxY"];
                StairsMinZ = stairs.Value["StairsMinZ"];
                StairsMaxZ = stairs.Value["StairsMaxZ"];
                StairsXLength = Math.Round(Utils.FootToMeter(Math.Abs(StairsMinX - StairsMaxX)), 2);
                StairsYLength = Math.Round(Utils.FootToMeter(Math.Abs(StairsMinY - StairsMaxY)), 2);
                StairsWidth = Math.Min(StairsXLength, StairsYLength);
                //stringBuilder.AppendLine($"楼梯：{stairs.Key.Id}，{stairs.Key.Name}，X轴长度：{StairsXLength}，Y轴长度：{StairsYLength}，X：({StairsMinX},{StairsMaxX})，Y：({StairsMinY},{StairsMaxY})，Z：({StairsMinZ},{StairsMaxZ})");
                stringBuilder.AppendLine($"楼梯：{stairs.Key.Id}");

                //先判断楼梯位置框内是否存在墙体或栏杆，将其分为“转向侧为墙体”、“转向侧为栏杆”或“不发生转向”两种情况
                //判断重叠的条件：如果一个矩形的最左边界大于另一个矩形的最右边界，且一个矩形的最上边界大于另一个矩形的最下边界，则两个矩形不重叠。否则，它们重叠。
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
                    if (WallMinX > StairsMaxX && WallMinY > StairsMaxY) //转向侧存在墙体时，平台净宽应≥1.3m
                    {
                        exist_wall = true;
                        //楼梯平台的坐标
                        foreach (KeyValuePair<Element, Dictionary<string, double>> stairslandings in StairsLandingsBoundingXYZ_Dict)
                        {
                            StairsLandingsMinX = stairslandings.Value["StairsLandingsMinX"];
                            StairsLandingsMaxX = stairslandings.Value["StairsLandingsMaxX"];
                            StairsLandingsMinY = stairslandings.Value["StairsLandingsMinY"];
                            StairsLandingsMaxY = stairslandings.Value["StairsLandingsMaxY"];
                            StairsLandingsMinZ = stairslandings.Value["StairsLandingsMinZ"];
                            StairsLandingsMaxZ = stairslandings.Value["StairsLandingsMaxZ"];
                            StairsLandingsXLength = stairslandings.Value["StairsLandingsXLength"];
                            StairsLandingsYLength = stairslandings.Value["StairsLandingsYLength"];
                            if ((StairsLandingsMinX >= StairsMinX - 1) && (StairsLandingsMaxX <= StairsMaxX + 1) && (StairsLandingsMinY >= StairsMinY - 1) && (StairsLandingsMaxY <= StairsMaxY + 1) && (StairsLandingsMinZ >= StairsMinZ - 1) && (StairsLandingsMaxZ <= StairsMaxZ + 1))
                            {
                                //判断平台净宽
                                if (WallXY == 0)
                                {
                                    StairsLandingsWidth = StairsLandingsXLength;
                                }
                                else
                                {
                                    StairsLandingsWidth = StairsLandingsYLength;
                                }
                                stairs_landins_num++;
                                //stringBuilder.AppendLine($"  -->楼梯平台：{stairslandings.Key.Id}，{stairslandings.Key.Name}，X轴长度：{StairsLandingsXLength}，Y轴长度：{StairsLandingsYLength}，X：({StairsLandingsMinX},{StairsLandingsMaxX})，Y：({StairsLandingsMinY},{StairsLandingsMaxY})，Z：({StairsLandingsMinZ},{StairsLandingsMaxZ})");
                                stringBuilder.AppendLine($"  -->楼梯平台：{stairslandings.Key.Id}，X轴长度：{StairsLandingsXLength}，Y轴长度：{StairsLandingsYLength}");
                                if (StairsLandingsWidth < 1.30) //转向侧为墙体，平台净宽应≥1.3m
                                {
                                    stringBuilder_notpass.Append(stairs.Key.Id.ToString() + "  ");
                                    H00099 = false;
                                }
                                else
                                {
                                    stringBuilder_notpass.Append(stairs.Key.Id.ToString() + "  ");

                                }
                            }
                        }
                    }
                }
                if (exist_wall == false)  // 如果转向侧不存在墙体，就进行栏杆判断，若转向一侧仅为栏杆，无墙体时，平台净宽应≥1.2m
                {
                    foreach (KeyValuePair<Element, Dictionary<string, double>> stairsrailing in StairsRailingBoundingXYZ_Dict)
                    {
                        StairsRailingMinX = stairsrailing.Value["StairsRailingMinX"];
                        StairsRailingMaxX = stairsrailing.Value["StairsRailingMaxX"];
                        StairsRailingMinY = stairsrailing.Value["StairsRailingMinY"];
                        StairsRailingMaxY = stairsrailing.Value["StairsRailingMaxY"];
                        StairsRailingMinZ = stairsrailing.Value["StairsRailingMinZ"];
                        StairsRailingMaxZ = stairsrailing.Value["StairsRailingMaxZ"];
                        StairsRailingXLength = stairsrailing.Value["StairsRailingXLength"];
                        StairsRailingYLength = stairsrailing.Value["StairsRailingYLength"];
                        StairsRailingXY = stairsrailing.Value["StairsRailingXY"];
                        if (StairsRailingMinX > StairsMaxX && StairsRailingMinY > StairsMaxY) //转向侧存在墙体时，平台净宽应≥1.3m
                        {
                            exist_stairsrailing = true;
                            //楼梯平台的坐标
                            foreach (KeyValuePair<Element, Dictionary<string, double>> stairslandings in StairsLandingsBoundingXYZ_Dict)
                            {
                                StairsLandingsMinX = stairslandings.Value["StairsLandingsMinX"];
                                StairsLandingsMaxX = stairslandings.Value["StairsLandingsMaxX"];
                                StairsLandingsMinY = stairslandings.Value["StairsLandingsMinY"];
                                StairsLandingsMaxY = stairslandings.Value["StairsLandingsMaxY"];
                                StairsLandingsMinZ = stairslandings.Value["StairsLandingsMinZ"];
                                StairsLandingsMaxZ = stairslandings.Value["StairsLandingsMaxZ"];
                                StairsLandingsXLength = stairslandings.Value["StairsLandingsXLength"];
                                StairsLandingsYLength = stairslandings.Value["StairsLandingsYLength"];
                                if ((StairsLandingsMinX >= StairsMinX - 1) && (StairsLandingsMaxX <= StairsMaxX + 1) && (StairsLandingsMinY >= StairsMinY - 1) && (StairsLandingsMaxY <= StairsMaxY + 1) && (StairsLandingsMinZ >= StairsMinZ - 1) && (StairsLandingsMaxZ <= StairsMaxZ + 1))
                                {
                                    //判断平台净宽
                                    if (StairsRailingXY == 0)
                                    {
                                        StairsLandingsWidth = StairsLandingsXLength;
                                    }
                                    else
                                    {
                                        StairsLandingsWidth = StairsLandingsYLength;
                                    }
                                    stairs_landins_num++;
                                    //stringBuilder.AppendLine($"  -->楼梯平台：{stairslandings.Key.Id}，{stairslandings.Key.Name}，X轴长度：{StairsLandingsXLength}，Y轴长度：{StairsLandingsYLength}，X：({StairsLandingsMinX},{StairsLandingsMaxX})，Y：({StairsLandingsMinY},{StairsLandingsMaxY})，Z：({StairsLandingsMinZ},{StairsLandingsMaxZ})");
                                    stringBuilder.AppendLine($"  -->楼梯平台：{stairslandings.Key.Id}，X轴长度：{StairsLandingsXLength}，Y轴长度：{StairsLandingsYLength}");
                                    if (StairsLandingsWidth < 1.20) //转向侧为墙体，平台净宽应≥1.3m
                                    {
                                        stringBuilder_notpass.Append(stairs.Key.Id.ToString() + "  ");
                                        H00099 = false;
                                    }
                                    else
                                    {
                                        stringBuilder_notpass.Append(stairs.Key.Id.ToString() + "  ");

                                    }
                                }
                            }
                        }
                    }
                }
                //如果不存在栏杆，则认为楼梯中不存在实体，
                
                






                foreach (KeyValuePair<Element, Dictionary<string, double>> stairsrailing in StairsRailingBoundingXYZ_Dict)
                {
                    StairsRailingMinX = stairsrailing.Value["StairsRailingMinX"];
                    StairsRailingMaxX = stairsrailing.Value["StairsRailingMaxX"];
                    StairsRailingMinY = stairsrailing.Value["StairsRailingMinY"];
                    StairsRailingMaxY = stairsrailing.Value["StairsRailingMaxY"];
                    StairsRailingMinZ = stairsrailing.Value["StairsRailingMinZ"];
                    StairsRailingMaxZ = stairsrailing.Value["StairsRailingMaxZ"];
                    StairsRailingXLength = Math.Round(Utils.FootToMeter(Math.Abs(StairsRailingMinX - StairsRailingMaxX)), 2);
                    StairsRailingYLength = Math.Round(Utils.FootToMeter(Math.Abs(StairsRailingMinY - StairsRailingMaxY)), 2);
                    if (StairsRailingMinX > StairsMaxX && StairsRailingMinY > StairsMaxY) //判断重叠的条件
                    { 
                        
                    }
                        
                }
                
                

                if (stairs_landins_num == 0) //如果楼梯没有平台，就认为是直跑楼梯
                {
                    if (StairsWidth < 0.90) {
                        stringBuilder_notpass.Append(stairs.Key.Id.ToString() + "  ");
                    } else {
                        stringBuilder_pass.Append(stairs.Key.Id.ToString() + "  ");
                    }
                } else //如果楼梯有平台，就判断平面的净宽是否符合要求
                {
                    if (StairsWidth < 1.20) {
                        stringBuilder_notpass.Append(stairs.Key.Id.ToString() + "  ");
                    } else {
                        stringBuilder_pass.Append(stairs.Key.Id.ToString() + "  ");
                    }
                }
                stringBuilder.AppendLine($"楼梯净宽：{StairsWidth}，有{stairs_landins_num}个平面");

            }

            */









            stringBuilder.AppendLine("不符合的楼梯：\n" + stringBuilder_notpass);
            stringBuilder.AppendLine("符合的楼梯：\n" + stringBuilder_pass);
            if (H00099 == true) {
                stringBuilder_result.AppendLine("符合/不符合5.3.5民用建筑通用规范 GB55031-2022");
            } else {

                stringBuilder_result.AppendLine("符合/不符合5.3.5民用建筑通用规范 GB55031-2022");
            }


            Utils.PrintLog(stringBuilder.ToString(), "H00099", document);
            TaskDialog.Show("FilterStairs", stringBuilder_result + stringBuilder.ToString());
            return Result.Succeeded;
        }
    }
}
