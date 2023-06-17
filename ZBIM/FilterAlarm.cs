using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    public class FilterAlarm : IExternalCommand
    {
        private double metre_to_inch(double metre)
        {
            double inch = 0;
            inch = metre * 3.28083989501;
            //return Math.Round(inch, 4)
            return inch;
        }
        private double inch_to_metre(double inch)
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
        private string StartSubString(string str, int startIdx, string endStr, bool isContains = false, bool isIgnoreCase = true)
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
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //界面交互的document
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            //实际内容的document
            Document document = commandData.Application.ActiveUIDocument.Document;

            StringBuilder stringBuilder = new StringBuilder();

            //存储符合与不符合国标的输出
            List<string> fit = new List<string>();
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


            //==========消火栓箱的==========
            FilteredElementCollector FireAlarmCollector = new FilteredElementCollector(document); //消火栓箱的过滤器
            //FireAlarmCollector.OfCategory(BuiltInCategory.OST_FireAlarmDevices).OfClass(typeof(FamilyInstance));  //过滤获得所有消火栓箱
            FireAlarmCollector.OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_FireAlarmDevices);
            List<Element> FireAlarmList = new List<Element>();
            foreach (var item in FireAlarmCollector)
            {
                stringBuilder.AppendLine(item.Name);
                //String name = item.LookupParameter("类型").AsValueString();//消防立管的缩写
                //if (name.Contains("消火栓箱"))
                //{
                //FireAlarmList.Add(item);
                //}
            }
            Dictionary<String, List<Element>> FireAlarm_buildNum = new Dictionary<string, List<Element>>();
            FireAlarm_buildNum.Add("1", new List<Element>());
            FireAlarm_buildNum.Add("2", new List<Element>());
            FireAlarm_buildNum.Add("3", new List<Element>());
            FireAlarm_buildNum.Add("4", new List<Element>());
            FireAlarm_buildNum.Add("5", new List<Element>());
            FireAlarm_buildNum.Add("6", new List<Element>());
            FireAlarm_buildNum.Add("7", new List<Element>());
            FireAlarm_buildNum.Add("8", new List<Element>());
            FireAlarm_buildNum.Add("9", new List<Element>());
            FireAlarm_buildNum.Add("10", new List<Element>());
            FireAlarm_buildNum.Add("11", new List<Element>());
            FireAlarm_buildNum.Add("12", new List<Element>());
            FireAlarm_buildNum.Add("13", new List<Element>());
            FireAlarm_buildNum.Add("14", new List<Element>());
            FireAlarm_buildNum.Add("15", new List<Element>());
            FireAlarm_buildNum.Add("16", new List<Element>());
            FireAlarm_buildNum.Add("17", new List<Element>());
            FireAlarm_buildNum.Add("18", new List<Element>());
            FireAlarm_buildNum.Add("19", new List<Element>());
            FireAlarm_buildNum.Add("20", new List<Element>());
            FireAlarm_buildNum.Add("21", new List<Element>());
            FireAlarm_buildNum.Add("22", new List<Element>());
            FireAlarm_buildNum.Add("23", new List<Element>());
            FireAlarm_buildNum.Add("24", new List<Element>());
            FireAlarm_buildNum.Add("25", new List<Element>());
            FireAlarm_buildNum.Add("26", new List<Element>());
            FireAlarm_buildNum.Add("27", new List<Element>());
            FireAlarm_buildNum.Add("28", new List<Element>());
            FireAlarm_buildNum.Add("29", new List<Element>());
            FireAlarm_buildNum.Add("30", new List<Element>());
            FireAlarm_buildNum.Add("Y1", new List<Element>());
            FireAlarm_buildNum.Add("S1", new List<Element>());
            FireAlarm_buildNum.Add("S2", new List<Element>());
            FireAlarm_buildNum.Add("S3", new List<Element>());

            //==========轴网的==========
            List<Element> elementsGroups = collectorGroups.OfCategory(BuiltInCategory.OST_IOSModelGroups).ToList<Element>();
            List<Element> elementsGrids = collectorGrid.OfCategory(BuiltInCategory.OST_Grids).ToList<Element>();//过滤出每栋楼的轴网


            List<Element> elementsGridsSort = new List<Element>();//排序后每栋楼的轴网

            //对轴网进行排序
            //清洗一下，过滤掉不需要的轴网
            foreach (var ele in elementsGrids)
            {
                if (ele.Name.Contains("-") &&
                    !ele.Name.Contains(".") &&
                    !ele.Name.Contains("d-") &&
                    !ele.Name.Contains("A-") &&
                    !ele.Name.Contains("B-") &&
                    !ele.Name.Contains("C-") &&
                    !ele.Name.Contains("D-") &&
                    !ele.Name.Contains("S1-") &&
                    !ele.Name.Contains("S2-") &&
                    !ele.Name.Contains("Y1-") &&
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
            Element temp_elementsGrids;
            int ii, jj;
            int length = elementsGridsSort.ToArray().Length;
            for (ii = length - 1; ii > 0; ii--)
            {
                //stringBuilder.AppendLine(StartSubString(elementsGrids[i].Name, 0, "-"));
                for (jj = 0; jj < ii; jj++)
                {
                    if (!"null".Equals(elementsGridsSort[jj].Name) && !"".Equals(elementsGridsSort[jj].Name) && !"null".Equals(elementsGridsSort[jj + 1].Name) && !"".Equals(elementsGridsSort[jj + 1].Name))
                    {
                        try
                        {
                            String j1_str = StartSubString(elementsGridsSort[jj].Name, 0, "-");
                            String j2_str = StartSubString(elementsGridsSort[jj + 1].Name, 0, "-");
                            int j1 = Convert.ToInt16(j1_str);
                            int j2 = Convert.ToInt16(j2_str);
                            if (j1 >= j2)
                            {
                                //stringBuilder.AppendLine("j1：" + j1 + "，j2：" + j2+"，j1>j2");
                                temp_elementsGrids = elementsGridsSort[jj];
                                elementsGridsSort[jj] = elementsGridsSort[jj + 1];
                                elementsGridsSort[jj + 1] = temp_elementsGrids;
                            }
                        }
                        catch (Exception e)//这里会输出异常的轴网，也就是无法排序的轴网
                        {
                            //stringBuilder.AppendLine("异常的楼层：" + elementsGridsSort[j].Name + "，" + elementsGridsSort[j + 1].Name);
                        }
                    }
                }
            }

            //最后单独把S1、S2、Y1的轴网加进排序好的轴网里
            foreach (var ele in elementsGrids)
            {
                if (ele.Name.Contains("S1-") ||
                    ele.Name.Contains("S2-") ||
                    ele.Name.Contains("Y1-")
                    )
                {
                    elementsGridsSort.Add(ele);
                }
            }

            double buildMinX = 0; //建筑最小的X坐标
            double buildMinY = 0; //建筑最小的Y坐标
            double buildMaxX = 0; //建筑最大的X坐标
            double buildMaxY = 0; //建筑最大的Y坐标


            double FireAlarmMaxX = 0; //消火栓箱的最大X坐标
            double FireAlarmMinX = 0; //消火栓箱的最小X坐标
            double FireAlarmMaxY = 0; //消火栓箱的最大Y坐标
            double FireAlarmMinY = 0; //消火栓箱的最小Y坐标
            double FireAlarmMaxZ = 0; //消火栓箱的最大Z坐标
            double FireAlarmMinZ = 0; //消火栓箱的最小Z坐标

            String item_name = null;
            String pre_item_name = null;
            BoundingBoxXYZ buildXYZ = new BoundingBoxXYZ();


            foreach (var item in elementsGridsSort)
            {
                if (item.Name.Contains("-") &&
                    !item.Name.Contains(".") &&
                    !item.Name.Contains("d-") &&
                    !item.Name.Contains("A") &&
                    !item.Name.Contains("B") &&
                    !item.Name.Contains("C") &&
                    !item.Name.Contains("D") &&
                    !item.Name.Contains("1S") &&
                    !item.Name.Contains("公") &&
                    !item.Name.Contains("测"))
                {
                    item_name = StartSubString(item.Name, 0, "-"); //取轴网从第一个字符开始到“-”结束的字符作为轴网的唯一名字
                    if (item_name == pre_item_name)//如果轴网不发生变化，就继续求BoundingXYZ的最大最小XY值
                    {
                        //计算轴网建筑的占地范围
                        buildXYZ = item.get_BoundingBox(document.ActiveView);
                        if (buildXYZ != null)
                        {
                            XYZ buildMax = buildXYZ.Max; //取到Max元组
                            XYZ buildMin = buildXYZ.Min; //取到Min元组
                            if (Math.Abs(buildMax.Y - buildMin.Y) < 10)//如果Y范围在误差内，就认为是一条横的轴网，计算X的最大最小值
                            {
                                if (buildMax.X > buildMin.X)
                                {
                                    buildMaxX = Math.Round(buildMax.X, 1);
                                    buildMinX = Math.Round(buildMin.X, 1);
                                }
                                else
                                {
                                    buildMaxX = Math.Round(buildMin.X, 1);
                                    buildMinX = Math.Round(buildMax.X, 1);
                                }
                            }
                            if (Math.Abs(buildMax.X - buildMin.X) < 10)//如果X范围在误差内，就认为是一条竖的轴网，计算Y的最大最小值
                            {
                                if (buildMax.Y > buildMin.Y)
                                {
                                    buildMaxY = Math.Round(buildMax.Y, 1);
                                    buildMinY = Math.Round(buildMin.Y, 1);
                                }
                                else
                                {
                                    buildMaxY = Math.Round(buildMin.Y, 1);
                                    buildMinY = Math.Round(buildMax.Y, 1);
                                }
                            }

                            //找每个楼栋的消火栓箱
                            foreach (var FireAlarm in FireAlarmList)
                            {
                                BoundingBoxXYZ FireAlarmBoundingBoxXYZ = FireAlarm.get_BoundingBox(document.ActiveView);
                                XYZ FireAlarmXYZMax = FireAlarmBoundingBoxXYZ.Max; //取到Max元组
                                XYZ FireAlarmXYZMin = FireAlarmBoundingBoxXYZ.Min; //取到Min元组
                                FireAlarmMaxX = FireAlarmXYZMax.X;
                                FireAlarmMinX = FireAlarmXYZMin.X;
                                FireAlarmMaxY = FireAlarmXYZMax.Y;
                                FireAlarmMinY = FireAlarmXYZMin.Y;
                                FireAlarmMaxZ = FireAlarmXYZMax.Z;
                                FireAlarmMinZ = FireAlarmXYZMin.Z;
                                if (FireAlarmMinX > buildMinX && FireAlarmMinY > buildMinY && FireAlarmMaxX < buildMaxX && FireAlarmMaxY < buildMaxY && !FireAlarm_buildNum[item_name].Contains(FireAlarm))
                                {
                                    stringBuilder.AppendLine("item.id" + item.Id + "，FireAlarm.LevelId：" + FireAlarm.LevelId + "，FireAlarm.Name：" + FireAlarm.Name + "，FireAlarm.Id" + FireAlarm.Id + "，方位：" + "X:(" + FireAlarmMinX + "," + FireAlarmMaxX + ")" + "，Y:(" + FireAlarmMinY + "," + FireAlarmMaxY + ")" + "，Z:(" + FireAlarmMinZ + "," + FireAlarmMaxZ + ")");
                                    //FireAlarm_buildNum[item_name].Add(FireAlarm);
                                }
                            }

                        }
                    }
                    else //如果前后两个轴网发生变化，说明已经切换到下一栋楼了，打印输出一下这栋楼的占地方位
                    {
                        //stringBuilder.AppendLine("以上是属于" + pre_item_name + "的物品");
                        //stringBuilder.AppendLine("轴网发生变化" + "pre_item_name：" + pre_item_name + "，item_name：" + item_name + "，" + pre_item_name + "的方位：" + "X:(" + buildMinX + "," + buildMaxX + ")" + "，Y:(" + buildMinY + "," + buildMaxY + ")");
                    }
                    //stringBuilder.AppendLine(item.Name);
                    //stringBuilder.AppendLine(item_name);
                    pre_item_name = item_name;
                }
            }
            //stringBuilder.AppendLine("轴网发生变化" + "pre_item_name：" + pre_item_name + "，item_name：" + item_name + "，" + pre_item_name + "的方位：" + "X:(" + buildMinX + "," + buildMaxX + ")" + "，Y:(" + buildMinY + "," + buildMaxY + ")");


            Boolean H00018_pass = true;
            StringBuilder H00018_result = new StringBuilder();


            if (H00018_pass == false)
            {
                H00018_result.AppendLine("检测结论：不符合8.2.1建筑设计防火规范 GB50016-2014（2018年版）");
            }
            else
            {
                H00018_result.AppendLine("检测结论：符合8.2.1建筑设计防火规范 GB50016-2014（2018年版）");
            }
            stringBuilder = H00018_result.AppendLine(stringBuilder.ToString());
            TaskDialog.Show("H00018强条检测", stringBuilder.ToString());
            return Result.Succeeded;
        }
    }
}
