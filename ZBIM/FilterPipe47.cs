using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;


namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    public class FilterPipe47 : IExternalCommand
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

            ElementId viewid = document.ActiveView.Id;

            FilteredElementCollector collectorGroups = new FilteredElementCollector(document, viewid); // 梁结构收集器
            FilteredElementCollector collectorGrid = new FilteredElementCollector(document); // 轴网收集器


            //==========自动喷火灭火系统==========
            double FirePipeMaxX; //消防立管的最大X坐标
            double FirePipeMinX; //消防立管的最小X坐标
            double FirePipeMaxY; //消火的最大Y坐标
            double FirePipeMinY; //消防立管的最小Y坐标
            double FirePipeMaxZ; //消防立管的最大Z坐标
            double FirePipeMinZ; //消防立管的最小Z坐标
            double FirePipeXLength; //消防立管的X轴长度
            double FirePipeYLength; //消防立管的Y轴长度
            double FirePipeZLength; //消防立管的Z轴长度

            FilteredElementCollector FirePipeCollector = new FilteredElementCollector(document);  //管道的过滤器
            FirePipeCollector.OfClass(typeof(Autodesk.Revit.DB.Plumbing.Pipe)).OfCategory(BuiltInCategory.OST_PipeCurves); //获得所有的管道
            List<Element> FirePipeList = new List<Element>();   //存放“ZP/PL消防立管”
            int pipe_num = 0;
            foreach (var FirePipe in FirePipeCollector)//从管道中筛选出缩写是“ZP”、“PL”的消防立管，并放入FirePipeList中
            {
                string abbr = FirePipe.LookupParameter("系统缩写").AsString();//消防立管的缩写
                if (abbr.Contains("ZP") || abbr.Contains("PL"))
                {
                    BoundingBoxXYZ FirePipeBoundingBoxXYZ = FirePipe.get_BoundingBox(document.ActiveView);
                    XYZ FirePipeXYZMax = FirePipeBoundingBoxXYZ.Max; //取到Max元组
                    XYZ FirePipeXYZMin = FirePipeBoundingBoxXYZ.Min; //取到Min元组
                    FirePipeMaxX = FirePipeXYZMax.X;
                    FirePipeMinX = FirePipeXYZMin.X;
                    FirePipeMaxY = FirePipeXYZMax.Y;
                    FirePipeMinY = FirePipeXYZMin.Y;
                    FirePipeMaxZ = Math.Round(inch_to_metre(FirePipeXYZMax.Z), 2);
                    FirePipeMinZ = Math.Round(inch_to_metre(FirePipeXYZMin.Z), 2);
                    FirePipeXLength = FirePipeMaxX - FirePipeMinX;
                    FirePipeYLength = FirePipeMaxY - FirePipeMinY;
                    FirePipeZLength = FirePipeMaxZ - FirePipeMinZ;
                    if (FirePipeMaxZ >= 0 && FirePipeMinZ >= 0)//首先在地坪上
                    {
                        if (FirePipeZLength > 2.0 || FirePipeYLength > 0.5 || FirePipeXLength > 0.5)//考虑竖管和横管
                        {
                            FirePipeList.Add(FirePipe);
                            pipe_num++;
                            stringBuilder.AppendLine("找到管道：" + FirePipe.Id + "，" + FirePipe.LevelId + "，" + abbr + "，X：（" + FirePipeMinX + "，" + FirePipeMaxX + "），Y：（" + FirePipeMinY + "，" + FirePipeMaxY + "），Z：(" + FirePipeMinZ + "，" + FirePipeMaxZ + ")");
                        }
                    }
                }
            }
            stringBuilder.AppendLine("一共有" + pipe_num + "根管道");
            Dictionary<string, List<Element>> FirePipe_buildNum = new Dictionary<string, List<Element>>
            {
                { "1", new List<Element>() },
                { "2", new List<Element>() },
                { "3", new List<Element>() },
                { "4", new List<Element>() },
                { "5", new List<Element>() },
                { "6", new List<Element>() },
                { "7", new List<Element>() },
                { "8", new List<Element>() },
                { "9", new List<Element>() },
                { "10", new List<Element>() },
                { "11", new List<Element>() },
                { "12", new List<Element>() },
                { "13", new List<Element>() },
                { "14", new List<Element>() },
                { "15", new List<Element>() },
                { "16", new List<Element>() },
                { "17", new List<Element>() },
                { "18", new List<Element>() },
                { "19", new List<Element>() },
                { "20", new List<Element>() },
                { "21", new List<Element>() },
                { "22", new List<Element>() },
                { "23", new List<Element>() },
                { "24", new List<Element>() },
                { "25", new List<Element>() },
                { "26", new List<Element>() },
                { "27", new List<Element>() },
                { "28", new List<Element>() },
                { "29", new List<Element>() },
                { "30", new List<Element>() },
                { "B1", new List<Element>() },
                { "B2", new List<Element>() },
                { "B3", new List<Element>() },
                { "Y1", new List<Element>() },
                { "Y2", new List<Element>() },
                { "Y3", new List<Element>() },
                { "S1", new List<Element>() },
                { "S2", new List<Element>() },
                { "S3", new List<Element>() }
            };



            //==========轴网过滤、排序==========
            List<Element> elementsGroups = collectorGroups.OfCategory(BuiltInCategory.OST_IOSModelGroups).ToList<Element>();
            List<Element> elementsGrids = collectorGrid.OfCategory(BuiltInCategory.OST_Grids).ToList<Element>();//过滤出每栋楼的轴网
            /*
            foreach (var ele in elementsGrids)
            { 
                stringBuilder.AppendLine(ele.Name);
            }
            */

            List<Element> elementsGridsSort = new List<Element>();//排序后每栋楼的轴网

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
            /*
            stringBuilder.AppendLine("排序前的轴网");
            foreach (var ele in elementsGridsSort)
            {
                stringBuilder.AppendLine(ele.Name);
            }
            */

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
                    ele.Name.Contains("S1-") ||
                    ele.Name.Contains("S2-") ||
                    ele.Name.Contains("S3-") ||
                    ele.Name.Contains("Y1-") ||
                    ele.Name.Contains("Y2-") ||
                    ele.Name.Contains("Y3-"))
                {
                    elementsGridsSort.Add(ele);
                }
            }

            //查看排序后的轴网
            /*
            foreach (var ele in elementsGridsSort)
            {
                stringBuilder.AppendLine(ele.Name);
            }
            */





            //==========找到每个轴网范围内的喷淋管道==========
            double buildMinX = 0; //建筑最小的X坐标
            double buildMinY = 0; //建筑最小的Y坐标
            double buildMaxX = 0; //建筑最大的X坐标
            double buildMaxY = 0; //建筑最大的Y坐标
            double buildMinZ = 0; //建筑最大的Z坐标
            double buildMaxZ = 0; //建筑最大的Z坐标

            string item_name;
            string pre_item_name = "0";

            Dictionary<Element, List<double>> FirePipeXYZ = new Dictionary<Element, List<double>>();
            //找到每栋楼的喷淋管
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

                        //找每个楼栋的喷淋管道
                        foreach (var FirePipe in FirePipeList)
                        {
                            BoundingBoxXYZ FirePipeBoundingBoxXYZ = FirePipe.get_BoundingBox(document.ActiveView);
                            XYZ FirePipeXYZMax = FirePipeBoundingBoxXYZ.Max; //取到Max元组
                            XYZ FirePipeXYZMin = FirePipeBoundingBoxXYZ.Min; //取到Min元组
                            FirePipeMaxX = FirePipeXYZMax.X;
                            FirePipeMinX = FirePipeXYZMin.X;
                            FirePipeMaxY = FirePipeXYZMax.Y;
                            FirePipeMinY = FirePipeXYZMin.Y;
                            FirePipeMaxZ = FirePipeXYZMax.Z;
                            FirePipeMinZ = FirePipeXYZMin.Z;
                            FirePipeXLength = Math.Round(inch_to_metre(Math.Abs(FirePipeMaxX - FirePipeMinX)), 2);
                            FirePipeYLength = Math.Round(inch_to_metre(Math.Abs(FirePipeMaxY - FirePipeMinY)), 2);
                            FirePipeZLength = Math.Round(inch_to_metre(Math.Abs(FirePipeMaxZ - FirePipeMinZ)), 2);
                            //stringBuilder.AppendLine("FirePipe.LevelId：" + FirePipe.LevelId + "，FirePipe.Name：" + FirePipe.Name + "，FirePipeAbbr：" + FirePipe.LookupParameter("系统缩写").AsString() + "，FirePipe.Id：" + FirePipe.Id + "，方位：" + "X:(" + FirePipeMinX + "," + FirePipeMaxX + ")" + "，Y:(" + FirePipeMinY + "," + FirePipeMaxY + ")" + "，Z:(" + FirePipeMinZ + "," + FirePipeMaxZ + ")" + "，Z轴长度：" + FirePipeZLength + "m");
                            //stringBuilder.AppendLine($"{FirePipeXLength},{FirePipeYLength},{FirePipeZLength}");
                            //stringBuilder.AppendLine($"管道：({Math.Round(FirePipeMinX,2)},{Math.Round(FirePipeMaxX,2)}),({Math.Round(FirePipeMinY,2)},{Math.Round(FirePipeMaxY,2)}),({Math.Round(FirePipeMinZ,2)},{Math.Round(FirePipeMaxZ,2)})");
                            //stringBuilder.AppendLine("现在item.Name是：" + item.Name + "，pre_item_name是：" + pre_item_name);
                            //stringBuilder.AppendLine($"楼栋：({Math.Round(buildMinX,2)},{Math.Round(buildMaxX, 2)}),({Math.Round(buildMinY, 2)},{Math.Round(buildMaxY, 2)}),({Math.Round(buildMinZ, 2)},{Math.Round(buildMaxZ,2)})");
                            if (FirePipeMinX > buildMinX && FirePipeMinY > buildMinY && FirePipeMaxX < buildMaxX && FirePipeMaxY < buildMaxY && !FirePipe_buildNum[item_name].Contains(FirePipe))
                            {
                                if ((FirePipeXLength < 1 && FirePipeYLength < 1 && FirePipeZLength >= 0.1) || FirePipeYLength > 0.1 || FirePipeXLength > 0.1) //筛选出垂直的管道（X,Y偏移量不超过1m，Z轴高度大于0.5m），去掉水平的管道
                                {
                                    FirePipe_buildNum[pre_item_name].Add(FirePipe);
                                    if (!FirePipeXYZ.ContainsKey(FirePipe))
                                    {
                                        FirePipeXYZ[FirePipe] = new List<double>();
                                        FirePipeXYZ[FirePipe].Add(FirePipeMinX);
                                        FirePipeXYZ[FirePipe].Add(FirePipeMaxX);
                                        FirePipeXYZ[FirePipe].Add(FirePipeMinY);
                                        FirePipeXYZ[FirePipe].Add(FirePipeMaxY);
                                        FirePipeXYZ[FirePipe].Add(FirePipeMinZ);
                                        FirePipeXYZ[FirePipe].Add(FirePipeMaxZ);
                                    }
                                    else
                                    {
                                        FirePipeXYZ[FirePipe].Add(FirePipeMinX);
                                        FirePipeXYZ[FirePipe].Add(FirePipeMaxX);
                                        FirePipeXYZ[FirePipe].Add(FirePipeMinY);
                                        FirePipeXYZ[FirePipe].Add(FirePipeMaxY);
                                        FirePipeXYZ[FirePipe].Add(FirePipeMinZ);
                                        FirePipeXYZ[FirePipe].Add(FirePipeMaxZ);
                                    }
                                    stringBuilder.AppendLine(pre_item_name + "#：" + FirePipe.Id + "，" + FirePipe.LookupParameter("系统缩写").AsString() + "，X：（" + FirePipeMaxX + "，" + FirePipeMinX + "），Y：（" + FirePipeMaxY + "，" + FirePipeMinY + "），Z：(" + FirePipeMinZ + "，" + FirePipeMaxZ + ")");
                                }
                            }
                        }
                        //stringBuilder.AppendLine("以上是属于" + pre_item_name + "的物品");
                        //stringBuilder.AppendLine("轴网变化：" + "前一个轴网：" + pre_item_name + "，现在轴网：" + item_name + "，" + pre_item_name + "的方位：" + "X:(" + buildMinX + "," + buildMaxX + ")" + "，Y:(" + buildMinY + "," + buildMaxY + ")");
                        //stringBuilder.AppendLine("轴网变化：" + "前一个轴网：" + pre_item_name + "，现在轴网：" + item_name);
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
                    foreach (var FirePipe in FirePipeList)
                    {
                        BoundingBoxXYZ FirePipeBoundingBoxXYZ = FirePipe.get_BoundingBox(document.ActiveView);
                        XYZ FirePipeXYZMax = FirePipeBoundingBoxXYZ.Max; //取到Max元组
                        XYZ FirePipeXYZMin = FirePipeBoundingBoxXYZ.Min; //取到Min元组
                        FirePipeMaxX = FirePipeXYZMax.X;
                        FirePipeMinX = FirePipeXYZMin.X;
                        FirePipeMaxY = FirePipeXYZMax.Y;
                        FirePipeMinY = FirePipeXYZMin.Y;
                        FirePipeMaxZ = FirePipeXYZMax.Z;
                        FirePipeMinZ = FirePipeXYZMin.Z;
                        FirePipeXLength = Math.Round(inch_to_metre(Math.Abs(FirePipeMaxX - FirePipeMinX)), 2);
                        FirePipeYLength = Math.Round(inch_to_metre(Math.Abs(FirePipeMaxY - FirePipeMinY)), 2);
                        FirePipeZLength = Math.Round(inch_to_metre(Math.Abs(FirePipeMaxZ - FirePipeMinZ)), 2);
                        //stringBuilder.AppendLine("FirePipe.LevelId：" + FirePipe.LevelId + "，FirePipe.Name：" + FirePipe.Name + "，FirePipeAbbr：" + FirePipe.LookupParameter("系统缩写").AsString() + "，FirePipe.Id：" + FirePipe.Id + "，方位：" + "X:(" + FirePipeMinX + "," + FirePipeMaxX + ")" + "，Y:(" + FirePipeMinY + "," + FirePipeMaxY + ")" + "，Z:(" + FirePipeMinZ + "," + FirePipeMaxZ + ")" + "，Z轴长度：" + FirePipeZLength + "m");
                        //stringBuilder.AppendLine($"{FirePipeXLength},{FirePipeYLength},{FirePipeZLength}");
                        //stringBuilder.AppendLine($"管道：({Math.Round(FirePipeMinX, 2)},{Math.Round(FirePipeMaxX, 2)}),({Math.Round(FirePipeMinY, 2)},{Math.Round(FirePipeMaxY, 2)}),({Math.Round(FirePipeMinZ, 2)},{Math.Round(FirePipeMaxZ, 2)})");
                        //stringBuilder.AppendLine("现在item.Name是：" + item.Name + "，pre_item_name是：" + pre_item_name);
                        //stringBuilder.AppendLine($"楼栋：({Math.Round(buildMinX, 2)},{Math.Round(buildMaxX, 2)}),({Math.Round(buildMinY, 2)},{Math.Round(buildMaxY, 2)}),({Math.Round(buildMinZ, 2)},{Math.Round(buildMaxZ, 2)})");
                        if (FirePipeMinX > buildMinX && FirePipeMinY > buildMinY && FirePipeMaxX < buildMaxX && FirePipeMaxY < buildMaxY && !FirePipe_buildNum[item_name].Contains(FirePipe))
                        {
                            if ((FirePipeXLength < 1 && FirePipeYLength < 1 && FirePipeZLength >= 0.1) || FirePipeYLength > 0.1 || FirePipeXLength > 0.1) //筛选出垂直的管道（X,Y偏移量不超过1m，Z轴高度大于0.5m），去掉水平的管道
                            {
                                FirePipe_buildNum[pre_item_name].Add(FirePipe);
                                stringBuilder.AppendLine(pre_item_name + "#：" + FirePipe.Id + "，" + FirePipe.LookupParameter("系统缩写").AsString() + "，X：（" + FirePipeMaxX + "，" + FirePipeMinX + "），Y：（" + FirePipeMaxY + "，" + FirePipeMinY + "），Z：(" + FirePipeMinZ + "，" + FirePipeMaxZ + ")");
                            }
                        }
                    }
                }
            }



            //查看每栋楼的喷淋管
            foreach (var item in FirePipe_buildNum)
            {
                stringBuilder.AppendLine($"楼号：{item.Key}#");
                foreach (var FirePipe in item.Value)
                {
                    BoundingBoxXYZ FirePipeBoundingBoxXYZ = FirePipe.get_BoundingBox(document.ActiveView);
                    XYZ FirePipeXYZMax = FirePipeBoundingBoxXYZ.Max; //取到Max元组
                    XYZ FirePipeXYZMin = FirePipeBoundingBoxXYZ.Min; //取到Min元组
                    FirePipeMaxX = FirePipeXYZMax.X;
                    FirePipeMinX = FirePipeXYZMin.X;
                    FirePipeMaxY = FirePipeXYZMax.Y;
                    FirePipeMinY = FirePipeXYZMin.Y;
                    FirePipeMaxZ = FirePipeXYZMax.Z;
                    FirePipeMinZ = FirePipeXYZMin.Z;
                    stringBuilder.AppendLine("找到管道：" + FirePipe.Id + "，" + FirePipe.LevelId + "，" + FirePipe.LookupParameter("系统缩写").AsString() + "，X：（" + FirePipeMaxX + "，" + FirePipeMinX + "），Y：（" + FirePipeMaxY + "，" + FirePipeMinY + "），Z：(" + FirePipeMinZ + "，" + FirePipeMaxZ + ")");

                }
            }


            //查看FirePipeXYZ的喷淋管
            /*
            stringBuilder.AppendLine("查看FirePipeXYZ的喷淋管");
            foreach (var FirePipe in FirePipeXYZ.Keys)
            {
                stringBuilder.AppendLine($"喷淋管Id：{FirePipe.Id}，名称：{FirePipe.Name}，X：({Math.Round(FirePipeXYZ[FirePipe][0], 2)}，{Math.Round(FirePipeXYZ[FirePipe][1], 2)})，Y：({Math.Round(FirePipeXYZ[FirePipe][2], 2)}，{Math.Round(FirePipeXYZ[FirePipe][3], 2)})，Z：({Math.Round(FirePipeXYZ[FirePipe][4], 2)}，{Math.Round(FirePipeXYZ[FirePipe][5], 2)})，缩写：{FirePipe.LookupParameter("系统缩写").AsString()}");
            }
            */


            TaskDialog.Show("H00047强条检测", stringBuilder.ToString());

            return Result.Succeeded;
        }
    }
}
