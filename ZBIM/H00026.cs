using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.ComponentModel;
using System.Data;
using System.Text.RegularExpressions;


namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    public class H00026 : IExternalCommand
    {
        private double metre_to_inch(double metre)
        {
            double inch = 0;
            inch = metre * 3.28083989501;
            //return Math.Round(inch, 4)
            return inch;
        }

        private static double inch_to_metre(double inch)
        {
            double metre = 0;
            metre = inch / 3.28083989501;
            return metre;
        }

        private double area_inch_to_metre(double area_inch)
        {
            double area_metre = 0;
            area_metre = area_inch / 3.28083989501 / 3.28083989501;
            return area_metre;
        }

        // (截头)截取字符串，从指定位置startIdx开始，出现"结束字符"位置之间的字符串，是否包含开始字符,是否忽略大小写
        private static string StartSubString(string str, int startIdx, string endStr, bool isContains = false, bool isIgnoreCase = true)
        {
            if (string.IsNullOrEmpty(str) || startIdx > str.Length - 1 || startIdx < 0)
                return string.Empty;
            int idx = str.IndexOf(endStr, startIdx, isIgnoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
            if (idx < 0) //没找到
                return string.Empty;
            return str.Substring(0, isContains ? idx + endStr.Length : idx);
        }

        // 反向截末——截取字符串，根据开始字符，开始搜索位置(反向的即右到左)，是否忽略大小写，是否包含开始字符
        private string LastSubEndString(string str, string endStr, bool isContains = false, bool isIgnoreCase = true)
        {
            if (string.IsNullOrEmpty(str))
                return string.Empty;
            int idx = str.LastIndexOf(endStr, isIgnoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
            if (idx < 0) //没找到
                return string.Empty;
            return str.Substring(isContains ? idx : idx + endStr.Length);
        }

        // (截中)截取字符串，根据开始字符,结束字符,是否包含开始字符,结束字符(默认为不包括),大小写是否敏感（从0位置开始）
        public static string SubBetweenString(string str, string startStr, string endstr, bool isContainsStartStr = false, bool isContainsEndStr = false, bool isIgnoreCase = true)
        {
            if (string.IsNullOrEmpty(str))
                return string.Empty;
            int staridx = str.IndexOf(startStr, 0, isIgnoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
            if (staridx < 0) //没找到
                return string.Empty;
            int endidx = str.IndexOf(endstr, staridx + startStr.Length, isIgnoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
            if (endidx < 0) //没找到
                return string.Empty;
            var start = isContainsStartStr ? staridx : staridx + startStr.Length;
            var end = isContainsEndStr ? endidx + endstr.Length : endidx;
            if (end <= start)
                return string.Empty;
            return str.Substring(start, end - start);
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

        //根据轴网确定建筑楼栋
        private double[] get_Floor_Line_BBox(string bulidName, ViewPlan viewPlan, List<Element> elementsGrids)
        {
            double[] xyxy = new double[4] { 0, 0, 0, 0 }; //记录最小和最大的(x,y)坐标
            bool[] emptyFlag = new bool[2] { true, true }; //判断xyxy是否为空

            //楼号, 目前还是用int来表示楼号
            if (bulidName != "0")
            {
                bool matchFlag = false; //匹配flag
                foreach (var item in elementsGrids)
                {
                    /*------判断是否为buildingNum号楼------*/
                    string name = item.Name;
                    if (name.Contains('-'))
                    {
                        string[] preName = name.Split('-');
                        if (preName[0] == bulidName)
                            matchFlag = true;
                        else
                            matchFlag = false;
                    }
                    /*------------------------------------*/

                    /*------------------------计算极限位置------------------------*/
                    if (matchFlag == true)
                    {
                        BoundingBoxXYZ BBox = null;
                        BBox = item.get_BoundingBox(viewPlan);
                        if (BBox != null)
                        {
                            double width_X = Math.Abs(BBox.Max.X - BBox.Min.X);
                            double width_Y = Math.Abs(BBox.Max.Y - BBox.Min.Y);
                            int cmp_dis = 0;
                            double value;
                            if (width_X > width_Y)
                            {
                                value = (BBox.Max.Y + BBox.Min.Y) / 2;
                                cmp_dis = 1;
                            }
                            else
                            {
                                value = (BBox.Max.X + BBox.Min.X) / 2;
                                cmp_dis = 0;
                            }

                            if (emptyFlag[cmp_dis] == true)
                            {
                                xyxy[cmp_dis] = value;
                                xyxy[cmp_dis + 2] = value;
                                emptyFlag[cmp_dis] = false;
                            }
                            else
                            {
                                if (xyxy[cmp_dis] > value)
                                    xyxy[cmp_dis] = value;
                                if (xyxy[cmp_dis + 2] < value)
                                    xyxy[cmp_dis + 2] = value;
                            }
                        }
                    }
                    /*------------------------计算极限位置------------------------*/
                }
            }
            /*------------------------取轴网包络矩形的1.2宽高------------------------*/
            double deta_X = (xyxy[2] - xyxy[0]) * 0.1;
            double deta_Y = (xyxy[3] - xyxy[1]) * 0.1;
            xyxy[0] -= deta_X;
            xyxy[2] += deta_X;
            xyxy[1] -= deta_Y;
            xyxy[3] += deta_Y;
            /*------------------------取轴网包络矩形的1.2宽高------------------------*/
            return xyxy; //返回数组{x_min, y_min, x_max, y_max}
        }

        //获得某一种的面积分区的所有视图
        public static List<ViewPlan> GetAreaView(Autodesk.Revit.DB.Document document, string AreaTypeName, bool isContainBasement = true)//获得某种类型面积分区的视图View
        {
            List<ViewPlan> AreaViewPlan = new List<ViewPlan>();
            if (document.ActiveView.GetType().Equals(typeof(ViewPlan)))
            {
                FilteredElementCollector ViewPlanCollector = new FilteredElementCollector(document).OfClass(typeof(ViewPlan));
                foreach (var viewPlan in ViewPlanCollector)
                {
                    try
                    {
                        if (viewPlan.LookupParameter("类型").AsValueString().Contains(AreaTypeName))
                        {
                            if (isContainBasement == true)
                            {
                                //stringBuilder.AppendLine("视图名：" + viewPlan.Name + "，视图ID：" + viewPlan.Id + "，视图类型：" + viewPlan_category);
                                //viewPlan2D = viewPlan;
                                AreaViewPlan.Add((ViewPlan)viewPlan);
                            }
                            else
                            {
                                if (!viewPlan.Name.Contains("B"))
                                {
                                    //stringBuilder.AppendLine("视图名：" + viewPlan.Name + "，视图ID：" + viewPlan.Id + "，视图类型：" + viewPlan_category);
                                    //viewPlan2D = viewPlan;
                                    AreaViewPlan.Add((ViewPlan)viewPlan);
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        //有些视图类型不是通过AsValueString属性获得的，因此这一步需要处理异常
                        //stringBuilder.AppendLine(e.Message);
                    }
                }
            }
            return AreaViewPlan;
        }


        //根据过滤好的消防电梯数组，判断每栋建筑有哪些消防电梯
        public static Dictionary<string, List<Room>> GetBuildRoom(Autodesk.Revit.DB.Document document, List<Room> RoomList, List<Element> elementsGridsSort)
        {
            double buildMinX = 0; //防火分区最小的X坐标
            double buildMinY = 0; //建筑最小的Y坐标
            double buildMaxX = 0; //建筑最大的X坐标
            double buildMaxY = 0; //建筑最大的Y坐标
            double buildMinZ = 0; //建筑最大的Z坐标
            double buildMaxZ = 0; //建筑最大的Z坐标

            double RoomMaxX; //消防电梯的最大X坐标
            double RoomMinX; //消防电梯的最小X坐标
            double RoomMaxY; //消防电梯大Y坐标
            double RoomMinY; //消防电梯的最小Y坐标
            double RoomMaxZ; //消防电梯的最大Z坐标
            double RoomMinZ; //消防电梯的最小Z坐标
            double RoomXLength; //消防电梯的X轴长度
            double RoomYLength; //消防电梯的Y轴长度
            double RoomZLength; //消防电梯的Z轴长度

            Dictionary<string, List<Room>> Room_buildNum = new Dictionary<string, List<Room>>
            {
                { "1", new List<Room>() },
                { "2", new List<Room>() },
                { "3", new List<Room>() },
                { "4", new List<Room>() },
                { "5", new List<Room>() },
                { "6", new List<Room>() },
                { "7", new List<Room>() },
                { "8", new List<Room>() },
                { "9", new List<Room>() },
                { "10", new List<Room>() },
                { "11", new List<Room>() },
                { "12", new List<Room>() },
                { "13", new List<Room>() },
                { "14", new List<Room>() },
                { "15", new List<Room>() },
                { "16", new List<Room>() },
                { "17", new List<Room>() },
                { "18", new List<Room>() },
                { "19", new List<Room>() },
                { "20", new List<Room>() },
                { "21", new List<Room>() },
                { "22", new List<Room>() },
                { "23", new List<Room>() },
                { "24", new List<Room>() },
                { "25", new List<Room>() },
                { "26", new List<Room>() },
                { "27", new List<Room>() },
                { "28", new List<Room>() },
                { "29", new List<Room>() },
                { "30", new List<Room>() },
                { "31", new List<Room>() },
                { "32", new List<Room>() },
                { "33", new List<Room>() },
                { "34", new List<Room>() },
                { "B1", new List<Room>() },
                { "B2", new List<Room>() },
                { "B3", new List<Room>() },
                { "B4", new List<Room>() },
                { "B5", new List<Room>() },
                { "B6", new List<Room>() },
                { "B7", new List<Room>() },
                { "B8", new List<Room>() },
                { "B9", new List<Room>() },
                { "Y1", new List<Room>() },
                { "Y2", new List<Room>() },
                { "Y3", new List<Room>() },
                { "Y4", new List<Room>() },
                { "Y5", new List<Room>() },
                { "Y6", new List<Room>() },
                { "Y7", new List<Room>() },
                { "Y8", new List<Room>() },
                { "Y9", new List<Room>() },
                { "S1", new List<Room>() },
                { "S2", new List<Room>() },
                { "S3", new List<Room>() },
                { "S4", new List<Room>() },
                { "S5", new List<Room>() },
                { "S6", new List<Room>() },
                { "S7", new List<Room>() },
                { "S8", new List<Room>() },
                { "S9", new List<Room>() }
            };

            string item_name;
            string pre_item_name = "0";

            //Dictionary<Element, List<double>> FireRoomXYZ = new Dictionary<Element, List<double>>();
            //找到每栋楼的消防电梯
            foreach (var item in elementsGridsSort)
            {
                //stringBuilder.AppendLine("ele.Name：" + ele.Name + "，ele.Id：" + ele.Id + "，ele.LevelId：" + ele.LevelId);
                //stringBuilder.AppendLine(StartSubString(ele.Name, 0, "-"));
                if (item.Name.Contains("-") &&
                    !item.Name.Contains(".") &&
                    !item.Name.Contains("d-") &&
                    !item.Name.Contains("A") &&
                    //!item.Name.Contains("B") &&
                    !item.Name.Contains("C") &&
                    !item.Name.Contains("D") &&
                    !item.Name.Contains("1S") &&
                    !item.Name.Contains("公") &&
                    !item.Name.Contains("测"))
                {
                    item_name = StartSubString(item.Name, 0, "-");
                    if (item_name == pre_item_name)//如果轴网不发生变化，就继续求BoundingXYZ的最大最小XY值
                    {
                        BoundingBoxXYZ buildBoundingBoxXYZ = item.get_BoundingBox(document.ActiveView);//计算轴网建筑的占地范围
                        XYZ buildXYZMax = buildBoundingBoxXYZ.Max; //取到Max元组
                        XYZ buildXYZMin = buildBoundingBoxXYZ.Min; //取到Min元组
                                                                   //stringBuilder.AppendLine("轴网：" + item_name + "的方位：" + "X:(" + buildXYZMin.X + "," + buildXYZMax.X + ")" + "，Y:(" + buildXYZMin.Y + "," + buildXYZMax.Y + ")");
                        if (Math.Abs(buildXYZMax.X - buildXYZMin.X) < 5) //如果Max和Min的X变化不大，就认为是一条竖的轴网，就只更新Y
                        {
                            buildMaxY = buildXYZMax.Y;
                            buildMinY = buildXYZMin.Y;
                        }
                        if (Math.Abs(buildXYZMax.Y - buildXYZMin.Y) < 5) //如果Max和Min的Y变化不大，就认为是一条水平的轴网，就只更新X
                        {
                            buildMaxX = buildXYZMax.X;
                            buildMinX = buildXYZMin.X;
                        }
                        buildMinZ = buildXYZMin.Z;
                        buildMaxZ = buildXYZMax.Z;
                        //stringBuilder.AppendLine(item.Name); 
                        //stringBuilder.AppendLine($"更新后的：({Math.Round(buildMinX, 2)},{Math.Round(buildMaxX, 2)}),({Math.Round(buildMinY, 2)},{Math.Round(buildMaxY, 2)}),({Math.Round(buildMinZ, 2)},{Math.Round(buildMaxZ, 2)})");

                        //stringBuilder.AppendLine("轴网：" + item_name + "的方位：" + "X:(" + buildMinX + "," + buildMaxX + ")" + "，Y:(" + buildMinY + "," + buildMaxY + ")");
                    }
                    else //如果前后两个轴网发生变化，说明已经切换到下一栋楼了，打印输出一下这栋楼的占地方位
                    {
                        //轴网变化，新的建筑占地坐标重新更新，此时进行物品的检测最准确

                        //找每个楼栋的消防电梯
                        foreach (var Room in RoomList)
                        {
                            BoundingBoxXYZ FireRoomBoundingBoxXYZ = Room.get_BoundingBox(document.ActiveView);
                            XYZ FireRoomXYZMax = FireRoomBoundingBoxXYZ.Max; //取到Max元组
                            XYZ FireRoomXYZMin = FireRoomBoundingBoxXYZ.Min; //取到Min元组
                            RoomMinX = FireRoomXYZMin.X;
                            RoomMaxX = FireRoomXYZMax.X;
                            RoomMinY = FireRoomXYZMin.Y;
                            RoomMaxY = FireRoomXYZMax.Y;
                            RoomMinZ = FireRoomXYZMin.Z;
                            RoomMaxZ = FireRoomXYZMax.Z;
                            RoomXLength = Math.Round(inch_to_metre(Math.Abs(RoomMaxX - RoomMinX)), 2);
                            RoomYLength = Math.Round(inch_to_metre(Math.Abs(RoomMaxY - RoomMinY)), 2);
                            RoomZLength = Math.Round(inch_to_metre(Math.Abs(RoomMaxZ - RoomMinZ)), 2);

                            if (RoomMinX > buildMinX && RoomMinY > buildMinY && RoomMaxX < buildMaxX && RoomMaxY < buildMaxY && !Room_buildNum[item_name].Contains(Room))
                            {
                                Room_buildNum[pre_item_name].Add(Room);
                            }
                        }

                        BoundingBoxXYZ buildBoundingBoxXYZ = item.get_BoundingBox(document.ActiveView);//计算轴网建筑的占地范围
                        XYZ buildXYZMax = buildBoundingBoxXYZ.Max; //取到Max元组
                        XYZ buildXYZMin = buildBoundingBoxXYZ.Min; //取到Min元组
                        buildMinX = buildXYZMin.X; //建筑最小的X坐标
                        buildMaxX = buildXYZMax.X; //建筑最大的X坐标
                        buildMinY = buildXYZMin.Y; //建筑最小的Y坐标
                        buildMaxY = buildXYZMax.Y; //建筑最大的Y坐标
                    }
                    pre_item_name = item_name;
                    //给最后一栋楼匹配管道
                    foreach (var FireRoom in RoomList)
                    {
                        BoundingBoxXYZ FireRoomBoundingBoxXYZ = FireRoom.get_BoundingBox(document.ActiveView);
                        XYZ FireRoomXYZMax = FireRoomBoundingBoxXYZ.Max; //取到Max元组
                        XYZ FireRoomXYZMin = FireRoomBoundingBoxXYZ.Min; //取到Min元组
                        RoomMaxX = FireRoomXYZMax.X;
                        RoomMinX = FireRoomXYZMin.X;
                        RoomMaxY = FireRoomXYZMax.Y;
                        RoomMinY = FireRoomXYZMin.Y;
                        RoomMaxZ = FireRoomXYZMax.Z;
                        RoomMinZ = FireRoomXYZMin.Z;
                        RoomXLength = Math.Round(inch_to_metre(Math.Abs(RoomMaxX - RoomMinX)), 2);
                        RoomYLength = Math.Round(inch_to_metre(Math.Abs(RoomMaxY - RoomMinY)), 2);
                        RoomZLength = Math.Round(inch_to_metre(Math.Abs(RoomMaxZ - RoomMinZ)), 2);

                        if (RoomMinX > buildMinX && RoomMinY > buildMinY && RoomMaxX < buildMaxX && RoomMaxY < buildMaxY && !Room_buildNum[item_name].Contains(FireRoom))
                        {
                            Room_buildNum[pre_item_name].Add(FireRoom);
                        }
                    }
                }
            }
            return Room_buildNum;
        }

        //关闭当前视图
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

        //获得排序后所有建筑的轴网
        public static List<Element> GetBuildGrid(Autodesk.Revit.DB.Document document)
        {
            // 轴网收集器
            FilteredElementCollector collectorGrid = new FilteredElementCollector(document);
            //过滤出每栋楼的轴网
            List<Element> elementsGrids = collectorGrid.OfCategory(BuiltInCategory.OST_Grids).ToList<Element>();
            //排序后每栋楼的轴网
            List<Element> elementsGridsSort = new List<Element>();
            //过滤掉不需要的轴网
            foreach (var ele in elementsGrids)
            {
                if (ele.Name.Contains("-") &&
                    !ele.Name.Contains(".") &&
                    !ele.Name.Contains("d-") &&
                    !ele.Name.Contains("A-") &&
                    !ele.Name.Contains("B-") &&
                    !ele.Name.Contains("C-") &&
                    !ele.Name.Contains("D-") &&
                    !ele.Name.Contains("B1-") &&
                    !ele.Name.Contains("B2-") &&
                    !ele.Name.Contains("B3-") &&
                    !ele.Name.Contains("S1-") &&
                    !ele.Name.Contains("S2-") &&
                    !ele.Name.Contains("S3-") &&
                    !ele.Name.Contains("Y1-") &&
                    !ele.Name.Contains("Y2-") &&
                    !ele.Name.Contains("Y3-") &&
                    !ele.Name.Contains("1S") &&
                    !ele.Name.Contains("公") &&
                    !ele.Name.Contains("测"))
                {
                    //stringBuilder.AppendLine("ele.Name：" + ele.Name + "，ele.Id：" + ele.Id + "，ele.LevelId：" + ele.LevelId);
                    //stringBuilder.AppendLine(StartSubString(ele.Name, 0, "-"));
                    elementsGridsSort.Add(ele);
                }
            }
            //对轴网进行冒泡操作
            for (int ii = elementsGridsSort.ToArray().Length - 1; ii > 0; ii--)
            {
                for (int jj = 0; jj < ii; jj++)
                {
                    if (!"null".Equals(elementsGridsSort[jj].Name) && !"".Equals(elementsGridsSort[jj].Name) && !"null".Equals(elementsGridsSort[jj + 1].Name) && !"".Equals(elementsGridsSort[jj + 1].Name))
                    {
                        //获取列表中第jj个轴网的开头第一个字符，一般就是数字，字母之前已经过滤掉了
                        int j1 = Convert.ToInt16(StartSubString(elementsGridsSort[jj].Name, 0, "-")); //将数字字符转为int类型比较
                        int j2 = Convert.ToInt16(StartSubString(elementsGridsSort[jj + 1].Name, 0, "-"));
                        if (j1 > j2)
                        {
                            Element temp_elementsGrids = elementsGridsSort[jj];
                            elementsGridsSort[jj] = elementsGridsSort[jj + 1];
                            elementsGridsSort[jj + 1] = temp_elementsGrids;
                        }
                    }
                }
            }
            //最后单独把B1、B2、B3、S1、S2、S3、Y1、Y2、Y3的轴网加进排序好的轴网里
            foreach (var ele in elementsGrids)
            {
                if (ele.Name.Contains("B1-") ||
                    ele.Name.Contains("B2-") ||
                    ele.Name.Contains("B3-") ||
                    ele.Name.Contains("B4-") ||
                    ele.Name.Contains("B5-") ||
                    ele.Name.Contains("B6-") ||
                    ele.Name.Contains("B7-") ||
                    ele.Name.Contains("B8-") ||
                    ele.Name.Contains("B9-") ||
                    ele.Name.Contains("S1-") ||
                    ele.Name.Contains("S2-") ||
                    ele.Name.Contains("S3-") ||
                    ele.Name.Contains("S4-") ||
                    ele.Name.Contains("S5-") ||
                    ele.Name.Contains("S6-") ||
                    ele.Name.Contains("S7-") ||
                    ele.Name.Contains("S8-") ||
                    ele.Name.Contains("S9-") ||
                    ele.Name.Contains("Y1-") ||
                    ele.Name.Contains("Y2-") ||
                    ele.Name.Contains("Y3-") ||
                    ele.Name.Contains("Y4-") ||
                    ele.Name.Contains("Y5-") ||
                    ele.Name.Contains("Y6-") ||
                    ele.Name.Contains("Y7-") ||
                    ele.Name.Contains("Y8-") ||
                    ele.Name.Contains("Y9-"))
                {
                    elementsGridsSort.Add(ele);
                }
            }
            //返回过滤和排序完的建筑轴网
            return elementsGridsSort;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //界面交互的document
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            //实际内容的document
            Document document = commandData.Application.ActiveUIDocument.Document;

            StringBuilder stringBuilder = new StringBuilder();
            StringBuilder stringBuilder_notpass = new StringBuilder();

            //存储符合与不符合国标的输出
            //List<string> fit = new List<string>();
            List<string> not_fit = new List<string>();

            List<Element> elementGroupsList = new List<Element>();
            List<Element> elementGridsList = new List<Element>();
            List<Element> elementFloorsList = new List<Element>();

            ElementId viewid = document.ActiveView.Id;

            FilteredElementCollector collectorGroups = new FilteredElementCollector(document, viewid); // 梁结构收集器
            FilteredElementCollector collectorGrid = new FilteredElementCollector(document); // 轴网收集器
            FilteredElementCollector collectorFloors = new FilteredElementCollector(document, viewid);
            FilteredElementCollector collectorLevels = new FilteredElementCollector(document, viewid);
            FilteredElementCollector collectorAreas = new FilteredElementCollector(document, viewid);
            FilteredElementCollector collectorRooms = new FilteredElementCollector(document, viewid);
            FilteredElementCollector collectorViews = new FilteredElementCollector(document, viewid);

            FilteredElementCollector collectorViewSchedules = new FilteredElementCollector(document);
            FilteredElementCollector collectorViewSections = new FilteredElementCollector(document);
            FilteredElementCollector collectorViewPlans = new FilteredElementCollector(document);

            List<Element> elementsGroups = collectorGroups.OfCategory(BuiltInCategory.OST_IOSModelGroups).ToList<Element>();
            List<Element> elementsGrids = collectorGrid.OfCategory(BuiltInCategory.OST_Grids).ToList<Element>();
            List<Element> elementsFloors = collectorFloors.OfCategory(BuiltInCategory.OST_Floors).ToList<Element>();
            List<Element> elementsLevels = collectorLevels.OfCategory(BuiltInCategory.OST_Levels).ToList<Element>();
            List<Element> elementsAreas = collectorAreas.OfCategory(BuiltInCategory.OST_Areas).ToList<Element>();
            List<Element> elementsRooms = collectorRooms.OfCategory(BuiltInCategory.OST_Rooms).ToList<Element>();
            List<Element> elementsViews = collectorViews.OfCategory(BuiltInCategory.OST_Views).ToList<Element>();


            IList<Element> viewSchedules = collectorViewSchedules.OfClass(typeof(ViewSchedule)).ToElements();
            IList<Element> viewSections = collectorViewSections.OfClass(typeof(ViewSection)).ToElements();
            IList<Element> viewPlans = collectorViewPlans.OfClass(typeof(ViewPlan)).ToElements();


            IList<ElementId> elementGroupsListId = new List<ElementId>();
            IList<ElementId> elementGridsListId = new List<ElementId>();
            //IList<ElementId> elementFloorsListId = new List<ElementId>();
            //IList<ElementId> elementLevelsListId = new List<ElementId>();

            IList<int> buildNumList = new List<int>();

            List<Room> ok_RoomList = new List<Room>();

            foreach (ViewSchedule view in viewSchedules)
            {
                if (view.Name.Contains("总建筑面积"))
                {
                    collectorAreas = new FilteredElementCollector(document, view.Id);
                    elementsAreas = collectorAreas.OfCategory(BuiltInCategory.OST_Areas).ToList<Element>();
                }
            }

            foreach (ViewSection viewSection in viewSections)
            {
                if (viewSection.Name.Contains("南"))
                {
                    collectorLevels = new FilteredElementCollector(document, viewSection.Id);
                    elementsLevels = collectorLevels.OfCategory(BuiltInCategory.OST_Levels).ToList<Element>();
                }
            }


            //========存放每栋楼的防火分区Area=========
            Dictionary<string, List<Area>> FireArea_buildNum = new Dictionary<string, List<Area>>
            {
                { "1", new List<Area>() },
                { "2", new List<Area>() },
                { "3", new List<Area>() },
                { "4", new List<Area>() },
                { "5", new List<Area>() },
                { "6", new List<Area>() },
                { "7", new List<Area>() },
                { "8", new List<Area>() },
                { "9", new List<Area>() },
                { "10", new List<Area>() },
                { "11", new List<Area>() },
                { "12", new List<Area>() },
                { "13", new List<Area>() },
                { "14", new List<Area>() },
                { "15", new List<Area>() },
                { "16", new List<Area>() },
                { "17", new List<Area>() },
                { "18", new List<Area>() },
                { "19", new List<Area>() },
                { "20", new List<Area>() },
                { "21", new List<Area>() },
                { "22", new List<Area>() },
                { "23", new List<Area>() },
                { "24", new List<Area>() },
                { "25", new List<Area>() },
                { "26", new List<Area>() },
                { "27", new List<Area>() },
                { "28", new List<Area>() },
                { "29", new List<Area>() },
                { "30", new List<Area>() },
                { "B1", new List<Area>() },
                { "B2", new List<Area>() },
                { "B3", new List<Area>() },
                { "Y1", new List<Area>() },
                { "Y2", new List<Area>() },
                { "Y3", new List<Area>() },
                { "S1", new List<Area>() },
                { "S2", new List<Area>() },
                { "S3", new List<Area>() }
            };

            //过滤得到所有面积分区Area，然后按照名字过滤，得到防火分区
            FilteredElementCollector AreaCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Areas);
            List<Element> areaLevels = new List<Element>();
            foreach (Area area in AreaCollector)
            {
                string up_levelname = null;
                int before = 0;
                int after = 0;
                if (area.AreaScheme.Name.Contains("防火分区"))
                {
                    areaLevels.Add(area.Level);
                    foreach (Level levelitem in areaLevels)//elementsLevels是所有楼层，我们需要按照buildFlag楼号来分类到对应的楼层
                    {
                        if (levelitem.Name.Contains("#-"))//处理楼层的名字
                        {
                            string levelitem_Name = levelitem.Name;
                            IEnumerable<int> levelitem_Name_indexList = FindIndexList(levelitem_Name, "#-");
                            foreach (var location in levelitem_Name_indexList)
                            {
                                if (levelitem_Name.Substring(location, 2) == "#-")
                                {
                                    up_levelname = StartSubString(levelitem_Name, 0, "#-", false, false);
                                    if (up_levelname.Contains("、"))
                                    {
                                        up_levelname = LastSubEndString(up_levelname, "、", false, false);
                                    }
                                    before = int.Parse(up_levelname);

                                    up_levelname = SubBetweenString(levelitem_Name, "#-", "#", false, false, false);
                                    after = int.Parse(up_levelname);

                                    if (after - before <= 0)
                                    {
                                        stringBuilder.AppendLine("标高命名有误");
                                    }
                                    else
                                    {
                                        for (int f = 0; f < after - before - 1; f++)
                                        {
                                            int temp = before + f + 1;
                                            levelitem_Name = levelitem_Name + "、" + temp + "#";
                                        }
                                    }

                                    levelitem_Name = levelitem_Name.Remove(location + 1, 1);
                                    levelitem_Name = levelitem_Name.Insert(location + 1, "、");
                                }
                                else
                                {
                                    stringBuilder.AppendLine("标高命名有误");
                                }
                            }
                            up_levelname = levelitem_Name;
                        }
                        else
                        {
                            up_levelname = levelitem.Name;
                        }
                    }
                    string area_build;
                    string[] area_builds = Regex.Split(up_levelname, "、");
                    if (area.Level.Name.Contains("B"))
                    {
                        area_build = StartSubString(area.Level.Name, 0, "-");
                        FireArea_buildNum[area_build].Add(area);
                        //stringBuilder.AppendLine("防火分区Id：" + area.Id + "，名称：" + area.Name + "，原标高：" + area.Level.Name + "，处理后的标高：" + up_levelname + "每栋楼号：" + area_build);
                    }
                    else
                    {
                        //stringBuilder.Append("防火分区Id：" + area.Id + "，名称：" + area.Name + "，原标高：" + area.Level.Name + "，处理后的标高：" + up_levelname + "，每栋楼号：");
                        foreach (string temp_area_build in area_builds)
                        {
                            if (temp_area_build.Contains("F"))
                            {
                                area_build = StartSubString(temp_area_build, 0, "#", true);
                            }
                            else
                            {
                                area_build = temp_area_build;
                            }
                            //去掉area_build里的“#”
                            area_build = StartSubString(area_build, 0, "#");
                            //stringBuilder.Append(area_build + " ");
                            FireArea_buildNum[area_build].Add(area);
                        }
                        //stringBuilder.AppendLine();
                    }
                }
            }


            Element grid_element;
            //获取建筑栋数以及每栋轴网大小
            int buildNum;

            Dictionary<string, List<int>> Room_buildNum_int = new Dictionary<string, List<int>>
            {
                { "1", new List<int>() },
                { "2", new List<int>() },
                { "3", new List<int>() },
                { "4", new List<int>() },
                { "5", new List<int>() },
                { "6", new List<int>() },
                { "7", new List<int>() },
                { "8", new List<int>() },
                { "9", new List<int>() },
                { "10", new List<int>() },
                { "11", new List<int>() },
                { "12", new List<int>() },
                { "13", new List<int>() },
                { "14", new List<int>() },
                { "15", new List<int>() },
                { "16", new List<int>() },
                { "17", new List<int>() },
                { "18", new List<int>() },
                { "19", new List<int>() },
                { "20", new List<int>() },
                { "21", new List<int>() },
                { "22", new List<int>() },
                { "23", new List<int>() },
                { "24", new List<int>() },
                { "25", new List<int>() },
                { "26", new List<int>() },
                { "27", new List<int>() },
                { "28", new List<int>() },
                { "29", new List<int>() },
                { "30", new List<int>() },
                { "31", new List<int>() },
                { "32", new List<int>() },
                { "33", new List<int>() },
                { "34", new List<int>() },
                { "B1", new List<int>() },
                { "B2", new List<int>() },
                { "B3", new List<int>() },
                { "B4", new List<int>() },
                { "B5", new List<int>() },
                { "B6", new List<int>() },
                { "B7", new List<int>() },
                { "B8", new List<int>() },
                { "B9", new List<int>() },
                { "Y1", new List<int>() },
                { "Y2", new List<int>() },
                { "Y3", new List<int>() },
                { "Y4", new List<int>() },
                { "Y5", new List<int>() },
                { "Y6", new List<int>() },
                { "Y7", new List<int>() },
                { "Y8", new List<int>() },
                { "Y9", new List<int>() },
                { "S1", new List<int>() },
                { "S2", new List<int>() },
                { "S3", new List<int>() },
                { "S4", new List<int>() },
                { "S5", new List<int>() },
                { "S6", new List<int>() },
                { "S7", new List<int>() },
                { "S8", new List<int>() },
                { "S9", new List<int>() }
            };

            //stringBuilder.AppendLine("过滤模型组轴网结构");

            foreach (var item in elementsGrids)
            {
                //stringBuilder.AppendLine(item.Name);
                if (item.Name.Contains("-") &&
                    !item.Name.Contains("Y") &&
                    !item.Name.Contains("S") &&
                    !item.Name.Contains(".") &&
                    !item.Name.Contains("d") &&
                    !item.Name.Contains("D") &&
                    !item.Name.Contains("公") &&
                    !item.Name.Contains("测"))
                {
                    //stringBuilder.AppendLine(item.Name);
                    buildNum = int.Parse(StartSubString(item.Name, 0, "-", false, false));
                    if (!buildNumList.Contains(buildNum))
                    {
                        buildNumList.Add(buildNum);

                        elementGridsList.Add(item);
                        grid_element = document.GetElement(new ElementId(item.Id.IntegerValue));
                        elementGridsListId.Add(grid_element.Id);

                        //stringBuilder.AppendLine(item.Id + " " + item.Name + " " + item.LevelId + " " + item.GetType());

                        //uiDocument.Selection.SetElementIds(elementGroupsListId);
                    }
                }
            }

            //==========轴网过滤、排序==========
            List<Element> elementsGridsSort = GetBuildGrid(document);

            //计算楼栋和楼层模块
            int GroupsNumber = buildNumList.Count;
            if (GroupsNumber != 0)
            {
                //stringBuilder.AppendLine("住宅建筑数量：" + GroupsNumber);
            }
            else
            {
                //stringBuilder.AppendLine("住宅建筑数量：" + 0);
            }
            //stringBuilder.AppendLine();

            //stringBuilder.AppendLine("过滤标高线");

            int i = -1;

            Level outdoor_floor = null;

            List<string> Area_level_pass = new List<string>();
            List<Area> Area_pass = new List<Area>();

            //地上建筑
            foreach (int buildFlag in buildNumList)
            {
                //stringBuilder.AppendLine();
                //stringBuilder.AppendLine("楼栋编号：" + buildFlag + "#");
                i++;
                //group_element = elementGroupsList[i];
                grid_element = elementGridsList[i];

                List<Element> build_upfloorsNumList = new List<Element>();
                List<Element> build_upfloorsNumList_reply = new List<Element>();

                outdoor_floor = null;
                Level parapet = null;
                Level roof = null;
                Level machine_room_roof = null;

                string up_levelname = null;

                int before = 0;
                int after = 0;

                double build_height = 0;

                foreach (Level levelitem in elementsLevels)
                {
                    if (levelitem.Name.Contains("#-"))
                    {
                        string levelitem_Name = levelitem.Name;
                        IEnumerable<int> levelitem_Name_indexList = FindIndexList(levelitem_Name, "#-");
                        foreach (var location in levelitem_Name_indexList)
                        {
                            if (levelitem_Name.Substring(location, 2) == "#-")
                            {
                                up_levelname = StartSubString(levelitem_Name, 0, "#-", false, false);
                                if (up_levelname.Contains("、"))
                                {
                                    up_levelname = LastSubEndString(up_levelname, "、", false, false);
                                }
                                before = int.Parse(up_levelname);

                                up_levelname = SubBetweenString(levelitem_Name, "#-", "#", false, false, false);
                                after = int.Parse(up_levelname);

                                if (after - before <= 0)
                                {
                                    stringBuilder.AppendLine("标高命名有误");
                                }
                                else
                                {
                                    for (int f = 0; f < after - before - 1; f++)
                                    {
                                        int temp = before + f + 1;
                                        levelitem_Name = levelitem_Name + "、" + temp + "#";
                                    }
                                }

                                levelitem_Name = levelitem_Name.Remove(location + 1, 1);
                                levelitem_Name = levelitem_Name.Insert(location + 1, "、");
                            }
                            else
                            {
                                stringBuilder.AppendLine("标高命名有误");
                            }
                        }
                        up_levelname = levelitem_Name;
                        //stringBuilder.AppendLine("标高处理" + up_levelname);
                    }
                    else
                    {
                        up_levelname = levelitem.Name;
                    }

                    //地上建筑
                    if (up_levelname.Contains("地坪"))
                    {
                        outdoor_floor = levelitem;
                    }

                    IEnumerable<int> indexList = FindIndexList(up_levelname, (Convert.ToString(buildFlag) + "#"));

                    foreach (var location in indexList)
                    {
                        if (location == 0)
                        {
                            if (up_levelname.Contains("女儿墙") &&
                                up_levelname.Contains(Convert.ToString(buildFlag) + "#"))
                            {
                                parapet = levelitem;
                            }
                            if (up_levelname.Contains("屋面") &&
                                up_levelname.Contains(Convert.ToString(buildFlag) + "#") &&
                                !up_levelname.Contains("机房") &&
                                !up_levelname.Contains("S" + Convert.ToString(buildFlag)))
                            {
                                roof = levelitem;
                            }
                            if (up_levelname.Contains("机房屋面") &&
                                up_levelname.Contains(Convert.ToString(buildFlag) + "#"))
                            {
                                machine_room_roof = levelitem;
                            }

                            if (up_levelname.Contains(Convert.ToString(buildFlag) + "#") &&
                                up_levelname.Contains("F") &&
                                (up_levelname.Contains("A") || up_levelname.Contains("建")) &&
                                !up_levelname.Contains("S") &&
                                !up_levelname.Contains("B") &&
                                !up_levelname.Contains("铺") &&
                                !up_levelname.Contains("结"))
                            {
                                build_upfloorsNumList.Add(levelitem);
                            }
                        }
                        else if (location > 0)
                        {
                            if (up_levelname.Substring(location - 1, 1) != "1" &&
                                up_levelname.Substring(location - 1, 1) != "2" &&
                                up_levelname.Substring(location - 1, 1) != "3" &&
                                up_levelname.Substring(location - 1, 1) != "4" &&
                                up_levelname.Substring(location - 1, 1) != "5" &&
                                up_levelname.Substring(location - 1, 1) != "6" &&
                                up_levelname.Substring(location - 1, 1) != "7" &&
                                up_levelname.Substring(location - 1, 1) != "8" &&
                                up_levelname.Substring(location - 1, 1) != "9" &&
                                up_levelname.Substring(location - 1, 1) != "0")
                            {
                                if (up_levelname.Contains("女儿墙") &&
                                up_levelname.Contains(Convert.ToString(buildFlag) + "#"))
                                {
                                    parapet = levelitem;
                                }
                                if (up_levelname.Contains("屋面") &&
                                    up_levelname.Contains(Convert.ToString(buildFlag) + "#") &&
                                    !up_levelname.Contains("机房") &&
                                    !up_levelname.Contains("S" + Convert.ToString(buildFlag)))
                                {
                                    roof = levelitem;
                                }
                                if (up_levelname.Contains("机房屋面") &&
                                    up_levelname.Contains(Convert.ToString(buildFlag) + "#"))
                                {
                                    machine_room_roof = levelitem;
                                }

                                if (up_levelname.Contains(Convert.ToString(buildFlag) + "#") &&
                                    up_levelname.Contains("F") &&
                                    (up_levelname.Contains("A") || up_levelname.Contains("建")) &&
                                    !up_levelname.Contains("S") &&
                                    !up_levelname.Contains("B") &&
                                    !up_levelname.Contains("铺") &&
                                    !up_levelname.Contains("结") &&
                                    !up_levelname.Contains("Y") &&
                                    !up_levelname.Contains("Y") &&
                                    !up_levelname.Contains("面积"))
                                {
                                    build_upfloorsNumList.Add(levelitem);
                                }
                            }
                        }
                    }

                    //if (up_levelname.Contains("F0.00") &&
                    //    up_levelname.Contains("建"))
                    //if (up_levelname.Contains("1F0.00") &&
                    //    up_levelname.Contains("A") &&
                    //    !up_levelname.Contains("S"))
                    //{
                    //    build_upfloorsNumList.Add(levelitem);
                    //}
                }

                //个别视图中会加入含有B的
                build_upfloorsNumList_reply = build_upfloorsNumList.ToList();
                foreach (var litem in build_upfloorsNumList)
                {
                    if (litem.Name.Contains("B"))
                    {
                        build_upfloorsNumList_reply.Remove(litem);
                    }
                }
                build_upfloorsNumList.Clear();
                build_upfloorsNumList = build_upfloorsNumList_reply.ToList();

                //stringBuilder.AppendLine("地上楼层数" + build_upfloorsNumList.Count);
                //foreach (var litem in build_upfloorsNumList)
                //{
                //    stringBuilder.AppendLine("楼层" + litem.Name);
                //}

                //stringBuilder.AppendLine("地下楼层数" + build_downfloorsNumList.Count);
                //foreach (var litem in build_downfloorsNumList)
                //{
                //    stringBuilder.AppendLine("楼层" + litem.Name);
                //}

                //地上建筑
                if (outdoor_floor == null)
                {
                    //stringBuilder.AppendLine("找不到室外地坪，无法判断");

                    not_fit.Add("楼栋编号：" + buildFlag + "#");
                    not_fit.Add("找不到室外地坪，无法判断");
                    not_fit.Add("");
                    continue;
                }
                else if (parapet == null)
                {
                    //stringBuilder.AppendLine("找不到女儿墙，无法判断");

                    not_fit.Add("楼栋编号：" + buildFlag + "#");
                    not_fit.Add("找不到女儿墙，无法判断");
                    not_fit.Add("");
                    continue;
                }
                else if (roof == null)
                {
                    //stringBuilder.AppendLine("找不到屋面，无法判断");

                    not_fit.Add("楼栋编号：" + buildFlag + "#");
                    not_fit.Add("找不到屋面，无法判断");
                    not_fit.Add("");
                    continue;
                }
                //else if (machine_room_roof == null)
                //{
                //    stringBuilder.AppendLine("找不机房屋面，无法判断");
                //    continue;
                //}
                else
                {
                    //stringBuilder.AppendLine("室外地坪标高：" + inch_to_metre(outdoor_floor.Elevation));
                    //stringBuilder.AppendLine("女儿墙标高：" + inch_to_metre(parapet.Elevation));
                    //stringBuilder.AppendLine("规划高度为" + Math.Round(Convert.ToDecimal(Math.Abs(inch_to_metre(outdoor_floor.Elevation) - inch_to_metre(parapet.Elevation))), 2, MidpointRounding.AwayFromZero));


                    //stringBuilder.AppendLine("屋面" + roof.Elevation);
                    //stringBuilder.AppendLine("消防高度为" + Math.Round(Convert.ToDecimal(Math.Abs(inch_to_metre(outdoor_floor.Elevation) - inch_to_metre(roof.Elevation))), 2, MidpointRounding.AwayFromZero));

                    //有机房屋面这个标高才判断
                    if (!(machine_room_roof == null))
                    {
                        //stringBuilder.AppendLine("机房高度为" + Math.Round(Convert.ToDecimal(Math.Abs(inch_to_metre(outdoor_floor.Elevation) - inch_to_metre(machine_room_roof.Elevation))), 2, MidpointRounding.AwayFromZero));

                        double area = 0;
                        double roof_area = 0;
                        double machine_room_roof_area = 0;

                        //stringBuilder.AppendLine("过滤模型楼板计算面积");

                        foreach (var flooritem in elementsFloors)
                        {
                            if (flooritem.GetType() == typeof(Floor) &&
                                flooritem.Name.Contains("结构"))
                            {
                                if (flooritem.LevelId == roof.Id)
                                {
                                    //创建几何选项
                                    Options opt = new Options();
                                    opt.ComputeReferences = true;
                                    //Options.ComputeReferences必须为true，否是拿到的几何体的Reference都将是null
                                    opt.DetailLevel = ViewDetailLevel.Fine;
                                    GeometryElement geometryElement = flooritem.get_Geometry(opt);//转换为几何元素

                                    double FaceArea = 0;
                                    //double EdgeLength = 0;

                                    foreach (GeometryObject geomObj in geometryElement)//获取到几何元素的边和面
                                    {
                                        Solid geomSolid = geomObj as Solid;
                                        if (null != geomSolid)
                                        {
                                            foreach (Face geoFace in geomSolid.Faces)
                                            {
                                                //得到元素面
                                                if (geoFace is PlanarFace)
                                                {
                                                    FaceArea += geoFace.Area;
                                                }
                                            }

                                            //foreach (Edge geoEdge in geomSolid.Edges)
                                            //{
                                            //    EdgeLength += geoEdge.ApproximateLength;
                                            //    //得到边
                                            //}
                                        }
                                    }

                                    area = FaceArea / 2;
                                    //stringBuilder.AppendLine("-：" + area);
                                    roof_area = roof_area + area;
                                }

                                if (flooritem.LevelId == machine_room_roof.Id)
                                {
                                    //创建几何选项
                                    Options opt = new Options();
                                    opt.ComputeReferences = true;
                                    //Options.ComputeReferences必须为true，否是拿到的几何体的Reference都将是null
                                    opt.DetailLevel = ViewDetailLevel.Fine;
                                    GeometryElement geometryElement = flooritem.get_Geometry(opt);//转换为几何元素

                                    double FaceArea = 0;
                                    //double EdgeLength = 0;

                                    foreach (GeometryObject geomObj in geometryElement)//获取到几何元素的边和面
                                    {
                                        Solid geomSolid = geomObj as Solid;
                                        if (null != geomSolid)
                                        {
                                            foreach (Face geoFace in geomSolid.Faces)
                                            {
                                                //得到元素面
                                                if (geoFace is PlanarFace)
                                                {
                                                    FaceArea += geoFace.Area;
                                                }
                                            }
                                        }
                                    }
                                    area = FaceArea / 2;
                                    //stringBuilder.AppendLine("+：" + area);
                                    machine_room_roof_area = machine_room_roof_area + area;
                                }
                            }
                        }
                        //stringBuilder.AppendLine("机房屋面面积："+ area_inch_to_metre(machine_room_roof_area));
                        //stringBuilder.AppendLine("屋面面积：" + area_inch_to_metre(roof_area));

                        if (roof_area >= machine_room_roof_area * 4)
                        {
                            //stringBuilder.AppendLine("机房屋面不超过总屋面1/4，不算建筑高度");
                            build_height = Math.Abs(inch_to_metre(outdoor_floor.Elevation) - inch_to_metre(parapet.Elevation));
                        }
                        else
                        {
                            //stringBuilder.AppendLine("机房屋面超过总屋面1/4，要算建筑高度");
                            build_height = Math.Abs(inch_to_metre(outdoor_floor.Elevation) - inch_to_metre(machine_room_roof.Elevation));
                        }
                    }
                    else
                    {
                        build_height = Math.Abs(inch_to_metre(outdoor_floor.Elevation) - inch_to_metre(parapet.Elevation));
                    }
                    //stringBuilder.AppendLine("建筑高度为" + Math.Round(Convert.ToDecimal(build_height), 2, MidpointRounding.AwayFromZero));
                }

                if (build_height != 0)
                {
                    if (build_height > 33)//仅地上建筑
                    {
                        //stringBuilder.AppendLine("!!!该建筑地上部分应该设置消防电梯");
                        //stringBuilder.AppendLine("建筑高度>33m");
                        //stringBuilder.AppendLine("地下室深度≤10m");
                        //stringBuilder.AppendLine("判断地上建筑");

                        foreach (var litem in build_upfloorsNumList)
                        {
                            foreach (ViewPlan viewPlan in viewPlans)
                            {
                                //stringBuilder.AppendLine("所有楼层：" + viewPlan.GetType() + " "+viewPlan.Id + " " + viewPlan.Name);
                                if (viewPlan.Name.Equals(litem.Name))
                                {
                                    collectorRooms = new FilteredElementCollector(document, viewPlan.Id);
                                    elementsRooms = collectorRooms.OfCategory(BuiltInCategory.OST_Rooms).ToList<Element>();

                                    double[] buildboundingBoxXY = get_Floor_Line_BBox(buildFlag.ToString(), viewPlan, elementsGrids);

                                    //stringBuilder.AppendLine("轴网坐标：" +buildFlag.ToString() + buildboundingBoxXY[0] + " " + buildboundingBoxXY[1] + " " + buildboundingBoxXY[2] + " " + buildboundingBoxXY[3]);

                                    foreach (Room roomitem in elementsRooms)
                                    {
                                        if (roomitem.LevelId == litem.Id &&
                                            roomitem.Name.Contains("消防电梯") &&
                                            !roomitem.Name.Contains("前室"))
                                        {
                                            // 得到编辑框
                                            BoundingBoxXYZ roomboundingBoxXYZ = roomitem.get_BoundingBox(viewPlan);
                                            //BoundingBoxXYZ roomboundingBoxXYZ = roomitem.get_BoundingBox(document.ActiveView);
                                            //通过边界框创建一个OutLine
                                            XYZ roommin = roomboundingBoxXYZ.Min;
                                            XYZ roommax = roomboundingBoxXYZ.Max;
                                            XYZ roompoint = new XYZ((roommax.X + roommin.X) / 2, (roommax.Y + roommin.Y) / 2, (roommax.Z + roommin.Z) / 2);
                                            //stringBuilder.AppendLine("房间中心点坐标：" + roompoint.X + " " + roompoint.Y);

                                            if (buildboundingBoxXY[0] <= roompoint.X &&
                                                buildboundingBoxXY[1] <= roompoint.Y &&
                                                buildboundingBoxXY[2] >= roompoint.X &&
                                                buildboundingBoxXY[3] >= roompoint.Y)
                                            {
                                                ok_RoomList.Add(roomitem);
                                                
                                                CloseCurrentView(uiDocument);
                                                string buildFlag_str = buildFlag.ToString();
                                                foreach (Area area in FireArea_buildNum[buildFlag_str])
                                                {
                                                    string area_level = SubBetweenString(area.Name, "-", "F");
                                                    string level_name = litem.Name; //仅用于记录楼层的名字，1#、2#、3#
                                                    while (level_name.Contains("#")) { level_name = SubBetweenString(level_name, "#", "F", false, true, true); }
                                                    string level_name_num;
                                                    level_name_num = StartSubString(level_name, 0, "F");
                                                    
                                                    if (area_level == level_name_num)
                                                    {
                                                        //==========防火分区的ViewPlan==========
                                                        List<ViewPlan> FireAreaViewPlan = GetAreaView(document, "防火分区");

                                                        foreach (View view in FireAreaViewPlan)
                                                        {
                                                            try
                                                            {
                                                                //拿到防火分区的坐标
                                                                uiDocument.ActiveView = view;
                                                                BoundingBoxXYZ FireAreaBoundingBoxXYZ = area.get_BoundingBox(document.ActiveView);
                                                                stringBuilder.AppendLine(FireAreaBoundingBoxXYZ.Min.X.ToString() + " " + FireAreaBoundingBoxXYZ.Max.X.ToString());

                                                                if (FireAreaBoundingBoxXYZ.Min.X <= roompoint.X &&
                                                                    FireAreaBoundingBoxXYZ.Min.Y <= roompoint.Y &&
                                                                    FireAreaBoundingBoxXYZ.Max.X >= roompoint.X &&
                                                                    FireAreaBoundingBoxXYZ.Max.Y >= roompoint.Y)
                                                                {
                                                                    Room_buildNum_int[buildFlag_str][FireArea_buildNum[buildFlag_str].IndexOf(area)] = 1;
                                                                    break;
                                                                }
                                                                CloseCurrentView(uiDocument);
                                                            }
                                                            catch { }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        //stringBuilder.AppendLine("已设置消防电梯的楼层：" + ok);
                    }
                    else
                    {
                        //stringBuilder.AppendLine("---该建筑不用设置消防电梯");
                        //stringBuilder.AppendLine("建筑高度≤33m");
                        //stringBuilder.AppendLine("地下室深度≤10m");
                        //fit.Add("楼栋编号：" + buildFlag + "#");
                        //fit.Add("建筑高度≤33m");
                        //fit.Add("消防电梯不作检测");
                        //fit.Add("");
                        string buildFlag_str = buildFlag.ToString();
                        for (int flag_i = 0; flag_i < Room_buildNum_int[buildFlag_str].Count(); flag_i++)
                        {
                            Room_buildNum_int[buildFlag_str][flag_i] = 0;
                        }
                    }
                }
            }


            //地下建筑
            List<Element> build_downfloorsNumList = new List<Element>();

            Level basement = null;

            string down_levelname = null;

            int down = 0;
            int down_flag = 0;

            double build_depth = 0;

            double build_basement_area = 0;

            foreach (Level levelitem in elementsLevels)
            {
                if (levelitem.Name.Contains("B") &&
                    levelitem.Name.Contains("-") &&
                    (levelitem.Name.Contains("A") || levelitem.Name.Contains("建")) &&
                    !levelitem.Name.Contains("S") &&
                    !levelitem.Name.Contains("Y"))
                {
                    build_downfloorsNumList.Add(levelitem);
                    down_levelname = SubBetweenString(levelitem.Name, "B", "-", false, false, false);
                    down = int.Parse(down_levelname);
                }

                //地下建筑 找到最底层
                if (down > down_flag)
                {
                    basement = levelitem;
                    down_flag = down;
                }
            }

            //stringBuilder.AppendLine("地下楼层数" + build_downfloorsNumList.Count);
            //foreach (var litem in build_downfloorsNumList)
            //{
            //    stringBuilder.AppendLine("楼层" + litem.Name);
            //}

            if (basement == null)
            {
                stringBuilder.AppendLine("没有地下室");
            }
            else if (!(outdoor_floor == null))
            {
                build_depth = Math.Abs(inch_to_metre(outdoor_floor.Elevation) - inch_to_metre(basement.Elevation));
                //stringBuilder.AppendLine("地下室深度为" + Math.Round(Convert.ToDecimal(build_depth), 2, MidpointRounding.AwayFromZero));

                if (elementsAreas.Count == 0)
                {
                    build_basement_area = -1;
                }
                else
                {
                    foreach (Area areaitem in elementsAreas)
                    {
                        //stringBuilder.AppendLine("面积" + areaitem.Area + " " + areaitem.Name);
                        if (areaitem.Name.Contains("面积"))
                        {
                            build_basement_area = build_basement_area + areaitem.Area;
                        }
                    }
                    //stringBuilder.AppendLine("地下室面积为" + Math.Round(Convert.ToDecimal(area_inch_to_metre(build_basement_area)), 2, MidpointRounding.AwayFromZero));
                }
            }

            if (build_depth != 0 && build_basement_area == -1)
            {
                stringBuilder.AppendLine("！！！没有地下室的面积明细表或该表为空，无法判断");
            }
            else
            {
                down = 0;

                if (buildNumList != null)//仅地下建筑
                {
                    //stringBuilder.AppendLine("!!!该建筑地下部分应该设置消防电梯");
                    //stringBuilder.AppendLine("地下室深度>10m且面积>3000平方米");
                    //stringBuilder.AppendLine("需要判断地下建筑");

                    foreach (var litem in build_downfloorsNumList)
                    {
                        foreach (ViewPlan viewPlan in viewPlans)
                        {
                            //stringBuilder.AppendLine("所有楼层：" + viewPlan.GetType() + " "+viewPlan.Id + " " + viewPlan.Name);
                            if (viewPlan.Name.Equals(litem.Name))
                            {
                                collectorRooms = new FilteredElementCollector(document, viewPlan.Id);
                                elementsRooms = collectorRooms.OfCategory(BuiltInCategory.OST_Rooms).ToList<Element>();

                                double[] buildboundingBoxXY = get_Floor_Line_BBox("d", viewPlan, elementsGrids);

                                foreach (Room roomitem in elementsRooms)
                                {
                                    if (roomitem.LevelId == litem.Id &&
                                        roomitem.Name.Contains("消防电梯") &&
                                        !roomitem.Name.Contains("前室"))
                                    {
                                        // 得到编辑框
                                        BoundingBoxXYZ roomboundingBoxXYZ = roomitem.get_BoundingBox(viewPlan);
                                        //BoundingBoxXYZ roomboundingBoxXYZ = roomitem.get_BoundingBox(document.ActiveView);
                                        //通过边界框创建一个OutLine
                                        XYZ roommin = roomboundingBoxXYZ.Min;
                                        XYZ roommax = roomboundingBoxXYZ.Max;
                                        XYZ roompoint = new XYZ((roommax.X + roommin.X) / 2, (roommax.Y + roommin.Y) / 2, (roommax.Z + roommin.Z) / 2);

                                        if (buildboundingBoxXY[0] <= roompoint.X &&
                                            buildboundingBoxXY[1] <= roompoint.Y &&
                                            buildboundingBoxXY[2] >= roompoint.X &&
                                            buildboundingBoxXY[3] >= roompoint.Y)
                                        {
                                            //stringBuilder.AppendLine("已设置消防电梯的楼层：" + litem.Name + " " + roomitem.Id + " " + roomitem.Name);
                                            ok_RoomList.Add(roomitem);

                                            View CurrnetView = null;
                                            CloseCurrentView(uiDocument);
                                            string buildFlag_str = null;
                                            if (litem.Name.Contains("B1"))
                                            {
                                                buildFlag_str = "B1";
                                            }
                                            else if (litem.Name.Contains("B2"))
                                            {
                                                buildFlag_str = "B2";
                                            }
                                            else if (litem.Name.Contains("B3"))
                                            {
                                                buildFlag_str = "B3";
                                            }
                                            else if (litem.Name.Contains("B4"))
                                            {
                                                buildFlag_str = "B4";
                                            }
                                            else
                                            {
                                                buildFlag_str = "B5";
                                            }
                                            foreach (Area area in FireArea_buildNum[buildFlag_str])
                                            {
                                                double area_firearea = Convert.ToDouble(area.LookupParameter("面积").AsValueString()); //得到Area指向的防火分区的面积，类似这种字段：“114514”
                                                string area_level = SubBetweenString(area.Name, "-", "F");
                                                string level_name = litem.Name; //仅用于记录楼层的名字，1#、2#、3#
                                                while (level_name.Contains("#")) { level_name = SubBetweenString(level_name, "#", "F", false, true, true); }
                                                string level_name_num;
                                                level_name_num = StartSubString(level_name, 0, "F");
                                                if (area_level == level_name_num)
                                                {
                                                    //==========防火分区的ViewPlan==========
                                                    List<ViewPlan> FireAreaViewPlan = GetAreaView(document, "防火分区");

                                                    foreach (View view in FireAreaViewPlan)
                                                    {
                                                        try
                                                        {
                                                            //拿到防火分区的坐标
                                                            uiDocument.ActiveView = view;
                                                            CurrnetView = view;
                                                            BoundingBoxXYZ FireAreaBoundingBoxXYZ = area.get_BoundingBox(document.ActiveView);

                                                            if (FireAreaBoundingBoxXYZ.Min.X <= roompoint.X &&
                                                                FireAreaBoundingBoxXYZ.Min.Y <= roompoint.Y &&
                                                                FireAreaBoundingBoxXYZ.Max.X >= roompoint.X &&
                                                                FireAreaBoundingBoxXYZ.Max.Y >= roompoint.Y)
                                                            {
                                                                Room_buildNum_int[buildFlag_str][FireArea_buildNum[buildFlag_str].IndexOf(area)] = 1;
                                                            }
                                                            CloseCurrentView(uiDocument);
                                                        }
                                                        catch { }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (build_depth > 10 && build_basement_area > 3000)//仅地下建筑
                {
                    //stringBuilder.AppendLine("!!!该建筑地下部分应该设置消防电梯");
                    //stringBuilder.AppendLine("地下室深度>10m且面积>3000平方米");
                    //stringBuilder.AppendLine("需要判断地下建筑");

                    foreach (var litem in build_downfloorsNumList)
                    {
                        foreach (ViewPlan viewPlan in viewPlans)
                        {
                            //stringBuilder.AppendLine("所有楼层：" + viewPlan.GetType() + " "+viewPlan.Id + " " + viewPlan.Name);
                            if (viewPlan.Name.Equals(litem.Name))
                            {
                                collectorRooms = new FilteredElementCollector(document, viewPlan.Id);
                                elementsRooms = collectorRooms.OfCategory(BuiltInCategory.OST_Rooms).ToList<Element>();

                                double[] buildboundingBoxXY = get_Floor_Line_BBox("d", viewPlan, elementsGrids);

                                foreach (Room roomitem in elementsRooms)
                                {
                                    if (roomitem.LevelId == litem.Id &&
                                        roomitem.Name.Contains("消防电梯") &&
                                        !roomitem.Name.Contains("前室"))
                                    {
                                        // 得到编辑框
                                        BoundingBoxXYZ roomboundingBoxXYZ = roomitem.get_BoundingBox(viewPlan);
                                        //BoundingBoxXYZ roomboundingBoxXYZ = roomitem.get_BoundingBox(document.ActiveView);
                                        //通过边界框创建一个OutLine
                                        XYZ roommin = roomboundingBoxXYZ.Min;
                                        XYZ roommax = roomboundingBoxXYZ.Max;
                                        XYZ roompoint = new XYZ((roommax.X + roommin.X) / 2, (roommax.Y + roommin.Y) / 2, (roommax.Z + roommin.Z) / 2);

                                        if (buildboundingBoxXY[0] <= roompoint.X &&
                                            buildboundingBoxXY[1] <= roompoint.Y &&
                                            buildboundingBoxXY[2] >= roompoint.X &&
                                            buildboundingBoxXY[3] >= roompoint.Y)
                                        {
                                            //stringBuilder.AppendLine("已设置消防电梯的楼层：" + litem.Name + " " + roomitem.Id + " " + roomitem.Name);
                                            ok_RoomList.Add(roomitem);

                                            View CurrnetView = null;
                                            CloseCurrentView(uiDocument);
                                            string buildFlag_str = null;
                                            if (litem.Name.Contains("B1"))
                                            {
                                                buildFlag_str = "B1";
                                            }
                                            else if (litem.Name.Contains("B2"))
                                            {
                                                buildFlag_str = "B2";
                                            }
                                            else if (litem.Name.Contains("B3"))
                                            {
                                                buildFlag_str = "B3";
                                            }
                                            else if (litem.Name.Contains("B4"))
                                            {
                                                buildFlag_str = "B4";
                                            }
                                            else
                                            {
                                                buildFlag_str = "B5";
                                            }
                                            foreach (Area area in FireArea_buildNum[buildFlag_str])
                                            {
                                                double area_firearea = Convert.ToDouble(area.LookupParameter("面积").AsValueString()); //得到Area指向的防火分区的面积，类似这种字段：“114514”
                                                string area_level = SubBetweenString(area.Name, "-", "F");
                                                string level_name = litem.Name; //仅用于记录楼层的名字，1#、2#、3#
                                                while (level_name.Contains("#")) { level_name = SubBetweenString(level_name, "#", "F", false, true, true); }
                                                string level_name_num;
                                                level_name_num = StartSubString(level_name, 0, "F");
                                                if (area_level == level_name_num)
                                                {
                                                    //==========防火分区的ViewPlan==========
                                                    List<ViewPlan> FireAreaViewPlan = GetAreaView(document, "防火分区");

                                                    foreach (View view in FireAreaViewPlan)
                                                    {
                                                        try
                                                        {
                                                            //拿到防火分区的坐标
                                                            uiDocument.ActiveView = view;
                                                            CurrnetView = view;
                                                            BoundingBoxXYZ FireAreaBoundingBoxXYZ = area.get_BoundingBox(document.ActiveView);

                                                            if (FireAreaBoundingBoxXYZ.Min.X <= roompoint.X &&
                                                                FireAreaBoundingBoxXYZ.Min.Y <= roompoint.Y &&
                                                                FireAreaBoundingBoxXYZ.Max.X >= roompoint.X &&
                                                                FireAreaBoundingBoxXYZ.Max.Y >= roompoint.Y)
                                                            {
                                                                Room_buildNum_int[buildFlag_str][FireArea_buildNum[buildFlag_str].IndexOf(area)] = 1;
                                                            }
                                                            CloseCurrentView(uiDocument);
                                                        }
                                                        catch { }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    //stringBuilder.AppendLine("---该建筑不用设置消防电梯");
                    //stringBuilder.AppendLine("建筑高度≤33m");
                    //stringBuilder.AppendLine("地下室深度≤10m");
                    //fit.Add("楼栋编号：" + "地下室");
                    ////fit.Add("深埋≤10m或总建筑面积≤3000平方米");
                    //fit.Add("无需设置消防电梯");
                    //fit.Add("");

                    foreach (var litem in build_downfloorsNumList)
                    {
                        string buildFlag_str = null;
                        if (litem.Name.Contains("B1"))
                        {
                            buildFlag_str = "B1";
                        }
                        else if (litem.Name.Contains("B2"))
                        {
                            buildFlag_str = "B2";
                        }
                        else if (litem.Name.Contains("B3"))
                        {
                            buildFlag_str = "B3";
                        }
                        else if (litem.Name.Contains("B4"))
                        {
                            buildFlag_str = "B4";
                        }
                        else
                        {
                            buildFlag_str = "B5";
                        }
                        for (int flag_i = 0; flag_i < Room_buildNum_int[buildFlag_str].Count(); flag_i++)
                        {
                            Room_buildNum_int[buildFlag_str][flag_i] = 0;
                        }
                    }
                }
            }

            bool flag_fire = true;

            foreach (var key in Room_buildNum_int.Keys)
            {
                foreach (int Room_buildNum_int_area_item in Room_buildNum_int[key])
                {
                    if (!(Room_buildNum_int_area_item == 0 ||
                        Room_buildNum_int_area_item == 1))
                    {
                        flag_fire = false;
                        stringBuilder_notpass.AppendLine("防火分区Id：" + FireArea_buildNum[key][Room_buildNum_int[key].IndexOf(Room_buildNum_int_area_item)].Id + "，名称：" + FireArea_buildNum[key][Room_buildNum_int[key].IndexOf(Room_buildNum_int_area_item)].Name);
                    }
                }
            }
            if (flag_fire == true)
            {
                stringBuilder.AppendLine("符合建筑设计防火规范GB50016-2014（2018年版）");
                stringBuilder.AppendLine();
            }
            else
            {
                stringBuilder.AppendLine("不符合建筑设计防火规范 GB50016-2014（2018年版）");
                stringBuilder.AppendLine();
                stringBuilder.AppendLine("消防电梯少于1台的防火分区有：");
                stringBuilder.Append(stringBuilder_notpass);
            }

            TaskDialog.Show("提示", stringBuilder.ToString());

            return Result.Succeeded;
        }
    }
}
