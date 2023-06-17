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


namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    public class H00047 : IExternalCommand
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

        public static void PrintLog(string info, string itemNum, Document doc, string path = null)
        {
            // 获取模型编号M0000x
            Regex regex = new Regex(@"(?<=\\)M[0-9]{5}");
            //string modelNum = "M00001";
            string docPath = doc.PathName;
            if (regex.IsMatch(docPath))
                //modelNum = regex.Match(docPath).ToString();

                // 默认路径为当前用户的桌面
                if (path == null)
                {
                    string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string date = DateTime.Now.ToString("yyyyMMdd-HHmmss");
                    //path = $"{desktop}\\{itemNum}-{modelNum}-{date}.txt";
                    path = $"{desktop}\\{itemNum}-{date}.txt";
                }

            // 打印文本信息至指定路径
            FileStream fs = new FileStream(path, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            //开始写入           
            sw.Write(info);
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //界面交互的document
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            //实际内容的document
            Document document = commandData.Application.ActiveUIDocument.Document;

            /*
            //////////////////////切换到2D视图////////////////////////////
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
            /////////////////////////////////////////////////////////
            */

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







            int area_num = 0;
            //过滤得到所有面积分区Area，然后按照名字过滤，得到防火分区
            FilteredElementCollector AreaCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Areas);
            foreach (Area area in AreaCollector)
            {
                string area_firearea = area.LookupParameter("面积").AsValueString();
                if (area.Name.Contains("防火"))
                {
                    area_num++;
                    Level area_level = area.Level;
                    string area_level_name = area_level.Name;
                    stringBuilder.AppendLine(area_num + "、Id：" + area.Id + "，名字：" + area.Name + "，面积：" + area_firearea + "，所在标高：" + area_level_name);

                }
            }
            stringBuilder.AppendLine("一共有" + area_num + "个防火分区\n\n");


            //========存放每栋楼的防火分区Tag=========，一栋楼会存在多个防火分区，将多个防火分区Tag存入List中
            Dictionary<string, List<SpatialElementTag>> FireArea_buildNum = new Dictionary<string, List<SpatialElementTag>>
            {
                { "1", new List<SpatialElementTag>() },
                { "2", new List<SpatialElementTag>() },
                { "3", new List<SpatialElementTag>() },
                { "4", new List<SpatialElementTag>() },
                { "5", new List<SpatialElementTag>() },
                { "6", new List<SpatialElementTag>() },
                { "7", new List<SpatialElementTag>() },
                { "8", new List<SpatialElementTag>() },
                { "9", new List<SpatialElementTag>() },
                { "10", new List<SpatialElementTag>() },
                { "11", new List<SpatialElementTag>() },
                { "12", new List<SpatialElementTag>() },
                { "13", new List<SpatialElementTag>() },
                { "14", new List<SpatialElementTag>() },
                { "15", new List<SpatialElementTag>() },
                { "16", new List<SpatialElementTag>() },
                { "17", new List<SpatialElementTag>() },
                { "18", new List<SpatialElementTag>() },
                { "19", new List<SpatialElementTag>() },
                { "20", new List<SpatialElementTag>() },
                { "21", new List<SpatialElementTag>() },
                { "22", new List<SpatialElementTag>() },
                { "23", new List<SpatialElementTag>() },
                { "24", new List<SpatialElementTag>() },
                { "25", new List<SpatialElementTag>() },
                { "26", new List<SpatialElementTag>() },
                { "27", new List<SpatialElementTag>() },
                { "28", new List<SpatialElementTag>() },
                { "29", new List<SpatialElementTag>() },
                { "30", new List<SpatialElementTag>() },
                { "B1", new List<SpatialElementTag>() },
                { "B2", new List<SpatialElementTag>() },
                { "B3", new List<SpatialElementTag>() },
                { "Y1", new List<SpatialElementTag>() },
                { "Y2", new List<SpatialElementTag>() },
                { "Y3", new List<SpatialElementTag>() },
                { "S1", new List<SpatialElementTag>() },
                { "S2", new List<SpatialElementTag>() },
                { "S3", new List<SpatialElementTag>() }
            };
            FilteredElementCollector SpatialElementTagCollector = new FilteredElementCollector(document).OfClass(typeof(Autodesk.Revit.DB.SpatialElementTag));
            List<SpatialElementTag> SpatialElementTagsList = new List<SpatialElementTag>();
            int tag_num = 0;
            foreach (SpatialElementTag Tag in SpatialElementTagCollector)
            {
                if (Tag.TagText.Contains("防火"))
                {
                    tag_num++;
                    string view_name = Tag.View.Name; //得到Tag所在标高，类似这种字段：“B2-9.300（建）”
                    string firearea = StartSubString(Tag.TagText, 0, "m²", true);// 得到Tag的面积，类似这种字段：“114514m²”
                    //stringBuilder.Append(tag_num + "、标签名Id：" + Tag.Id.ToString() + "，面积：" + firearea + "，标签TagText：" + Tag.TagText + "，标签Name：" + Tag.Name + "，所属视图：" + view_name);
                    stringBuilder.Append(tag_num + "、标签Id：" + Tag.Id.ToString() + "，标签TagText：" + Tag.TagText + "，面积：" + firearea + "，所属视图：" + view_name);


                    if (Tag.TagText.Contains("B") || Tag.TagText.Contains("Y") || Tag.TagText.Contains("S"))
                    {
                        string tag_level = SubBetweenString(Tag.TagText, "防火分区", "-");// 类似字段：“B1”，地下的Tag只有层数信息，没有楼号
                        stringBuilder.AppendLine("，层号：" + tag_level);
                        FireArea_buildNum[tag_level].Add(Tag);
                    }
                    else
                    {
                        string temp_tag_level = SubBetweenString(Tag.TagText, "防火分区", "F-");// 类似字段：“5#-3”
                        string tag_level = LastSubEndString(temp_tag_level, "-"); //进一步，提取层号，类似字段：“3”
                        string tag_build_num = StartSubString(temp_tag_level, 0, "#"); // 提取楼号，类似字段：“5”
                        stringBuilder.AppendLine("，楼号：" + tag_build_num + "#" + "，层号：" + tag_level + "F");
                        FireArea_buildNum[tag_build_num].Add(Tag);
                    }
                }
            }
            stringBuilder.AppendLine("一共有" + tag_num + "个防火分区标签\n\n");


            stringBuilder.AppendLine("查看每栋楼每层的防火分区：");
            foreach (var item in FireArea_buildNum)
            {
                stringBuilder.AppendLine($"楼号：{item.Key}#");
                foreach (SpatialElementTag Tag in item.Value)
                {
                    string view_name = Tag.View.Name; //得到Tag所在标高，类似这种字段：“B2-9.300（建）”
                    string firearea = StartSubString(Tag.TagText, 0, "m²", true);// 得到Tag的面积，类似这种字段：“114514m²”
                    if (Tag.TagText.Contains("B") || Tag.TagText.Contains("Y") || Tag.TagText.Contains("S"))
                    {
                        string tag_level = LastSubEndString(Tag.TagText, "-");
                        stringBuilder.AppendLine("第" + tag_level + "个：标签名Id：" + Tag.Id.ToString() + "，面积：" + firearea + "，标签TagText：" + Tag.TagText + "，所属视图：" + view_name);
                    }
                    else
                    {
                        string tag_level = SubBetweenString(Tag.TagText, "-", "F");
                        stringBuilder.AppendLine("第" + tag_level + "层：标签名Id：" + Tag.Id.ToString() + "，面积：" + firearea + "，标签TagText：" + Tag.TagText + "，所属视图：" + view_name);
                    }
                }
            }




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
                    double FirePipeLength = FirePipeMaxZ - FirePipeMinZ;
                    if ((FirePipeMaxZ >= 0 && FirePipeMinZ >= 0) && FirePipeLength > 2.0)
                    {
                        FirePipeList.Add(FirePipe);
                        pipe_num++;
                        //stringBuilder.AppendLine("找到管道：" + FirePipe.Id + "，" + FirePipe.LevelId + "，" + abbr + "，X：（" + FirePipeMaxX + "，" + FirePipeMinX + "），Y：（" + FirePipeMaxY + "，" + FirePipeMinY + "），Z：(" + FirePipeMinZ + "，" + FirePipeMaxZ + ")");
                    }
                }
            }
            //stringBuilder.AppendLine("一共有" + pipe_num + "根管道");
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

            /*
            stringBuilder.AppendLine("排序后的轴网");
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

            string item_name;
            string pre_item_name = "0";


            //查看排序号的轴网
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
                            if (FirePipeMinX > buildMinX && FirePipeMinY > buildMinY && FirePipeMaxX < buildMaxX && FirePipeMaxY < buildMaxY && !FirePipe_buildNum[item_name].Contains(FirePipe))
                            {
                                //stringBuilder.AppendLine("FirePipe.LevelId：" + FirePipe.LevelId + "，FirePipe.Name：" + FirePipe.Name + "，FirePipeAbbr：" + FirePipe.LookupParameter("系统缩写").AsString() + "，FirePipe.Id：" + FirePipe.Id + "，方位：" + "X:(" + FirePipeMinX + "," + FirePipeMaxX + ")" + "，Y:(" + FirePipeMinY + "," + FirePipeMaxY + ")" + "，Z:(" + FirePipeMinZ + "," + FirePipeMaxZ + ")" + "，Z轴长度：" + FirePipeZLength + "m");
                                if (FirePipeXLength < 1 && FirePipeYLength < 1 && FirePipeZLength >= 0.5) //筛选出垂直的管道（X,Y偏移量不超过1m，Z轴高度大于0.5m），去掉水平的管道
                                {
                                    FirePipe_buildNum[pre_item_name].Add(FirePipe);
                                    //stringBuilder.AppendLine(pre_item_name + "#：" + FirePipe.Id + "，" + FirePipe.LookupParameter("系统缩写").AsString() + "，X：（" + FirePipeMaxX + "，" + FirePipeMinX + "），Y：（" + FirePipeMaxY + "，" + FirePipeMinY + "），Z：(" + FirePipeMinZ + "，" + FirePipeMaxZ + ")");

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
                }
            }









            List<Element> elementsLevels = collectorLevels.OfCategory(BuiltInCategory.OST_Levels).ToList<Element>();
            List<Element> elementsAreas = collectorAreas.OfCategory(BuiltInCategory.OST_Areas).ToList<Element>();
            List<Element> elementsRooms = collectorRooms.OfCategory(BuiltInCategory.OST_Rooms).ToList<Element>();
            List<Element> elementsViews = collectorViews.OfCategory(BuiltInCategory.OST_Views).ToList<Element>();

            IList<Element> viewSchedules = collectorViewSchedules.OfClass(typeof(ViewSchedule)).ToElements();
            IList<Element> viewSections = collectorViewSections.OfClass(typeof(ViewSection)).ToElements();
            IList<Element> viewPlans = collectorViewPlans.OfClass(typeof(ViewPlan)).ToElements();

            IList<ElementId> elementGroupsListId = new List<ElementId>();
            IList<ElementId> elementGridsListId = new List<ElementId>();

            FilteredElementCollector levelid = new FilteredElementCollector(document);   //楼层的过滤器
            ICollection<ElementId> levelIds = levelid.OfClass(typeof(Level)).ToElementIds();  //过滤获得所有楼层的id
            List<ElementId> LevelIds = new List<ElementId>();   //存放每一个楼层的id 

            Boolean H00047_pass = true;
            StringBuilder H00047_result = new StringBuilder();


            
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
            
            Element grid_element;
            //获取建筑栋数以及每栋轴网大小
            string buildNum;//存放楼号，例如1#、2#
            IList<string> buildNumList = new List<string>();//存放楼号的列表，buildNumList里存的是1、2、3、……、S1、Y1、S2、……、30这些楼栋的编号
            //去除不需要的标高
            foreach (var item in elementsGridsSort)
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
                    buildNum = StartSubString(item.Name, 0, "-", false, false);
                    //stringBuilder.AppendLine(Convert.ToString(buildNum));
                    if (!buildNumList.Contains(buildNum))
                    {
                        buildNumList.Add(buildNum);
                        elementGridsList.Add(item);
                        grid_element = document.GetElement(new ElementId(item.Id.IntegerValue));//编号不重复的首个轴网，1-1、2-1、3-1
                        //stringBuilder.AppendLine(grid_element.Name);
                        elementGridsListId.Add(grid_element.Id);
                    }
                }
                else if (item.Name.Substring(0, 1) == "Y" || item.Name.Substring(0, 1) == "S")
                {
                    buildNum = StartSubString(item.Name, 0, "-", false, false);
                    //stringBuilder.AppendLine(Convert.ToString(item.Name));
                    if (!buildNumList.Contains(buildNum))
                    {
                        buildNumList.Add(buildNum);
                        //stringBuilder.AppendLine(buildNum);//打印看一下，应该能找到S1、Y1、S2这三栋楼
                        elementGridsList.Add(item);
                        grid_element = document.GetElement(new ElementId(item.Id.IntegerValue));
                        //stringBuilder.AppendLine(grid_element.Name);
                        elementGridsListId.Add(grid_element.Id);
                    }
                }
            }
            //输出查看楼栋号
            /*
            foreach (string str in buildNumList)
            {
                stringBuilder.AppendLine(str);
            }
            */

            //计算楼栋和楼层模块
            //int GroupsNumber = buildNumList.Count;
            //stringBuilder.AppendLine("检测到的建筑数量：" + GroupsNumber);//没有9#、26#，所以最后应该是3+28=31栋建筑
            //stringBuilder.AppendLine("过滤标高线");
            int i = -1;
            Level outdoor_floor = null;

            //地上建筑，从这里开始遍历楼栋
            foreach (string buildFlag in buildNumList)
            {
                //buildFireAlarms.Add(buildFlag,该栋楼中的所有消火栓箱)
                //stringBuilder.AppendLine("--\n楼栋编号：" + buildFlag + "#");
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
                //从这里开始遍历楼层
                foreach (Level levelitem in elementsLevels)//elementsLevels是所有楼层，我们需要按照buildFlag楼号来分类到对应的楼层
                {
                    //stringBuilder.AppendLine("levelitem.Name:" + levelitem.Name);
                    //此处会遍历每一个所有楼层，因此在此处添加所有楼层id，并去除重复的楼层id
                    if (!LevelIds.Contains(levelitem.Id))
                    {
                        LevelIds.Add(levelitem.Id);//LevelIds存放了所有的楼层的id，是不重复的
                        //stringBuilder.AppendLine(levelitem.Name);
                    }

                    //处理楼层的名字
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
                        /*
                        if (up_levelname.Contains("Y") && buildFlag.Contains("Y"))
                        {
                            stringBuilder.AppendLine(up_levelname);
                            //build_upfloorsNumList.Add(levelitem);
                        }
                        if (up_levelname.Contains("S") && buildFlag.Contains("S"))
                        {
                            stringBuilder.AppendLine(up_levelname);
                            //build_upfloorsNumList.Add(levelitem);
                        }
                        */

                        if (location == 0)
                        {
                            if (up_levelname.Contains("女儿墙") &&
                                up_levelname.Contains(Convert.ToString(buildFlag) + "#"))
                            {
                                parapet = levelitem;
                                //stringBuilder.AppendLine("parapet.Name："+parapet.Name);
                            }
                            if ((up_levelname.Contains("屋面") || up_levelname.Contains("屋顶")) &&
                                up_levelname.Contains(Convert.ToString(buildFlag) + "#") &&
                                !up_levelname.Contains("机房") &&
                                !up_levelname.Contains("S" + Convert.ToString(buildFlag))
                                )
                            {
                                roof = levelitem;
                                //stringBuilder.AppendLine("roof.Name：" + roof.Name);
                            }
                            if (up_levelname.Contains("机房屋面") &&
                                up_levelname.Contains(Convert.ToString(buildFlag) + "#"))
                            {
                                machine_room_roof = levelitem;
                                //stringBuilder.AppendLine("machine_room_roof.Name：" + machine_room_roof.Name);
                            }

                            if (up_levelname.Contains(Convert.ToString(buildFlag) + "#") &&
                                up_levelname.Contains("F") &&
                                (up_levelname.Contains("A") || up_levelname.Contains("建")) &&
                                //!up_levelname.Contains("S") &&  //要检测S1、S2建筑，所以这行要注释掉
                                !up_levelname.Contains("B") &&
                                !up_levelname.Contains("铺") &&
                                !up_levelname.Contains("结")
                                )
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
                                up_levelname.Substring(location - 1, 1) != "0"
                                )
                            {
                                if (up_levelname.Contains("女儿墙") &&
                                up_levelname.Contains(Convert.ToString(buildFlag) + "#"))
                                {
                                    parapet = levelitem;
                                    //stringBuilder.AppendLine("parapet.Name："+parapet.Name);
                                }
                                if ((up_levelname.Contains("屋面") || up_levelname.Contains("屋顶")) &&
                                    up_levelname.Contains(Convert.ToString(buildFlag) + "#") &&
                                    !up_levelname.Contains("机房") &&
                                    !up_levelname.Contains("S" + Convert.ToString(buildFlag))
                                    )
                                {
                                    roof = levelitem;
                                    //stringBuilder.AppendLine("roof.Name：" + roof.Name);
                                }
                                if (up_levelname.Contains("机房屋面") &&
                                    up_levelname.Contains(Convert.ToString(buildFlag) + "#"))
                                {
                                    machine_room_roof = levelitem;
                                    //stringBuilder.AppendLine("machine_room_roof.Name：" + machine_room_roof.Name);
                                }

                                if (up_levelname.Contains(Convert.ToString(buildFlag) + "#") &&
                                    up_levelname.Contains("F") &&
                                    (up_levelname.Contains("A") || up_levelname.Contains("建")) &&
                                    !up_levelname.Contains("S") &&
                                    !up_levelname.Contains("Y") &&
                                    !up_levelname.Contains("B") &&
                                    !up_levelname.Contains("铺") &&
                                    !up_levelname.Contains("结")
                                    )
                                {
                                    build_upfloorsNumList.Add(levelitem);
                                }
                            }
                        }
                        else
                        {
                            // 我也不知道这里是干嘛的？
                        }
                    }
                }

                //个别视图中会加入含有B的，因此要去掉
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


                //对build_upfloorsNumList进行排序，数组里面的是楼平面的名字
                //为什么要排序呢？因为在输出Level的时候，有些Level放在了后面，相应地也放在了build_upfloorsNumList后面，这会导致有些楼地楼栋顺序不对
                for (int k = build_upfloorsNumList.Count - 1; k > 0; k--)
                {

                    for (int kk = 0; kk < k; kk++)
                    {
                        Level level = (Level)build_upfloorsNumList[kk];
                        Level level_next = (Level)build_upfloorsNumList[kk + 1];
                        string levelName_str = level.Name;
                        string levelName_str_next = level_next.Name;
                        while (levelName_str.Contains("#"))
                        {
                            levelName_str = SubBetweenString(levelName_str, "#", "F", false, true, true);
                        }
                        levelName_str = StartSubString(levelName_str, 0, "F", false, true);
                        int levelName_int = Convert.ToInt16(levelName_str);

                        while (levelName_str_next.Contains("#"))
                        {
                            levelName_str_next = SubBetweenString(levelName_str_next, "#", "F", false, true, true);
                        }
                        levelName_str_next = StartSubString(levelName_str_next, 0, "F", false, true);
                        int levelName_int_next = Convert.ToInt16(levelName_str_next);
                        if (levelName_int > levelName_int_next)
                        {
                            //stringBuilder.AppendLine("发现顺序错误");
                            Level temp = (Level)build_upfloorsNumList[kk];
                            build_upfloorsNumList[kk] = build_upfloorsNumList[kk + 1];
                            build_upfloorsNumList[kk + 1] = temp;
                        }
                    }
                    //stringBuilder.AppendLine(buildFlag.ToString() + "，" + levelName_str + "，" + levelName_str_next + "，" + level.Name);
                }
                //下面计算楼栋高度
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
                else
                {
                    build_height = Math.Abs(inch_to_metre(outdoor_floor.Elevation) - inch_to_metre(roof.Elevation));
                }

                //输出建筑的高度
                stringBuilder.AppendLine("");
                //stringBuilder.Append(buildFlag + "建筑高度");
                //stringBuilder.Append("楼栋编号：" + buildFlag + "#，建筑高度：" + Math.Round(build_height, 2) + "m");
                stringBuilder.Append(buildFlag + "# 建筑高度：" + Math.Round(build_height, 2) + "m");

                if (build_height <= 27)
                {
                    stringBuilder.AppendLine("（≤27m），单、多层民用建筑");
                }
                else if (build_height > 27 && build_height <= 54)
                {
                    stringBuilder.AppendLine("（27-54m），二类高层住宅");
                }
                else
                {
                    stringBuilder.AppendLine("（＞54m），一类高层住宅");
                }


                StringBuilder stringBuilder_pass = new StringBuilder();
                StringBuilder stringBuilder_notpass = new StringBuilder();


                //没排序前的楼层号存放数组
                List<int> string_pass = new List<int>();
                List<int> string_notpass = new List<int>();
                //没排序前每层楼的检测结果内容
                List<string> string_pass_result = new List<string>();
                List<string> string_notpass_result = new List<string>();

                //排序后，将每层楼和楼的检测结果内容一一结合后，再输出
                string pass_FirePipe;   //消防立管通过的id号
                string level_name = "";  //仅用于记录楼层的名字，1#、2#、3#
                int level_name_num;

                double level_height = 0;
                //排序后，将每层楼和楼的检测结果内容一一结合后，再输出






                if (build_height > 54)
                {
                    foreach (Level litem in build_upfloorsNumList) //从这里开始遍历建筑的楼层
                    {
                        //levelNum++;
                        //stringBuilder.AppendLine("litem.Name:" + litem.Name + "，litem.Id:" + litem.Id);
                        pass_FirePipe = "";
                        level_name = litem.Name;
                        while (level_name.Contains("#"))
                        {
                            level_name = SubBetweenString(level_name, "#", "F", false, true, true);
                        }
                        level_name_num = Convert.ToInt16(StartSubString(level_name, 0, "F"));
                        //判断消防立管
                        bool exist_FirePipe = false;
                        //判断消防立管
                        if (litem.Name.Contains("避难层"))
                        {
                            level_height = 56.4;
                        }
                        else
                        {
                            //利用楼层的“名称”属性获得高度，然后将String转成Double，就是真正的楼层对地高度
                            //获得小数点"."后面的一个字符
                            int point_idx = litem.Name.IndexOf(".");
                            string nextpoint = litem.Name.Substring(point_idx + 3, 2);//获得小数点后三个起始的字符，作为SubBetweenString的截断字符，长度只需取2
                            level_height = Convert.ToDouble(SubBetweenString(litem.Name, "F", nextpoint, false, false));
                        }
                        foreach (Element FirePipe in FirePipe_buildNum[buildFlag])
                        {
                            BoundingBoxXYZ FirePipeBoundingBoxXYZ = FirePipe.get_BoundingBox(document.ActiveView);
                            XYZ FirePipeXYZMax = FirePipeBoundingBoxXYZ.Max;
                            XYZ FirePipeXYZMin = FirePipeBoundingBoxXYZ.Min;
                            FirePipeMaxX = FirePipeXYZMax.X;
                            FirePipeMinX = FirePipeXYZMin.X;
                            FirePipeMaxY = FirePipeXYZMax.Y;
                            FirePipeMinY = FirePipeXYZMin.Y;
                            FirePipeMaxZ = FirePipeXYZMax.Z;
                            FirePipeMinZ = FirePipeXYZMin.Z;
                            FirePipeZLength = Math.Round(inch_to_metre(Math.Abs(FirePipeMaxZ - FirePipeMinZ)), 2);

                            if ((FirePipeMinZ <= level_height && level_height <= FirePipeMaxZ) || (FirePipeMinZ <= level_height && FirePipeMaxZ <= level_height)) //这样就认为这一层里有喷淋立管
                            {
                                pass_FirePipe += FirePipe.Id + "，";
                                exist_FirePipe = true;
                            }
                        }
                        //上面判断该层是否有喷淋系统

                        //下面读取该层的防火分区面积
                        //判断该楼的该层是否有防火分区，buildFlag代表楼号，level_name代表层号
                        foreach (SpatialElementTag Tag in FireArea_buildNum[buildFlag])
                        {
                            string firearea = StartSubString(Tag.TagText, 0, "m²", true);// 得到Tag指向的防火分区的面积，类似这种字段：“114514m²”
                            stringBuilder.AppendLine("第" + level_name + "层：标签名Id：" + Tag.Id + "，面积：" + firearea + "，标签TagText：" + Tag.TagText);






                            /*
                            string view_name = Tag.View.Name;
                            string fire_area = StartSubString(Tag.TagText, 0, "m²", false);// 得到这种字段：“114514m²”
                            //stringBuilder.AppendLine(Tag.Id.ToString() + "，" + fire_area + "，标签名：" + Tag.TagText + "\n所属视图：" + view_name);
                            //获取该防火分区在哪些楼
                            string[] build_num_part = view_name.Split(new char[] { '、' }); //用“、”分割长名字
                            string level_num = "未知楼层？";
                            //获取该防火分区在哪一层，获取该层的楼层号
                            for (int ii = 0; ii < build_num_part.Length; ii++) //使用for循环输出
                            {
                                if (build_num_part[ii].Contains("B") || build_num_part[ii].Contains("Y") || build_num_part[ii].Contains("S"))
                                {
                                    level_num = StartSubString(build_num_part[ii], 0, "-");
                                }
                                else
                                {
                                    level_num = StartSubString(LastSubEndString(build_num_part[ii], "#"), 0, "F");
                                }
                            }
                            //stringBuilder.AppendLine("所属楼层：" + level_num + "，当前楼层：" + StartSubString(level_name, 0, "F"));
                            if (level_num.Equals(StartSubString(level_name, 0, "F")))//如果该防火分区的所属楼层等于当前检测到的楼层，则
                            {
                                //stringBuilder.AppendLine(level_name + "，" + litem.Id + "，防火分区Id：" + Tag.Id + "，" +fire_area);
                                if ((exist_FirePipe == true && Convert.ToDouble(fire_area) > 3000) || (exist_FirePipe == false && Convert.ToDouble(fire_area) > 1500)) //高度大于27m，有喷淋，面积不超过3000m^2，无喷淋，面积不得超过1500m^2
                                {
                                    //运用面积比较法，将Tag上标注的面积和所有防火分区面积对比，如果两者相同则说明是Tag对应的防火分区，将该不符合的防火分区的Id和名字输出
                                    foreach (Area area in AreaCollector)
                                    {
                                        string area_firearea = area.LookupParameter("面积").AsValueString();
                                        //stringBuilder.AppendLine("Id：" + area.Id + "，名字：" + area.Name + "，面积：" + area_firearea);
                                        if (Convert.ToDouble(area_firearea).Equals(Math.Round(Convert.ToDouble(fire_area), 2)))
                                        {
                                            stringBuilder_notpass.AppendLine(level_name + "，防火分区Id：" + area.Id + "，名字：" + area.Name);// 添加不符合的防火分区
                                            break;
                                        }
                                    }
                                    H00047_pass = false;
                                }
                                else
                                {
                                    //运用面积比较法，将Tag上标注的面积和所有防火分区面积对比，如果两者相同则说明是Tag对应的防火分区，将该不符合的防火分区的Id和名字输出
                                    foreach (Area area in AreaCollector)
                                    {
                                        string area_firearea = area.LookupParameter("面积").AsValueString();
                                        //stringBuilder.AppendLine("Id：" + area.Id + "，名字：" + area.Name + "，面积：" + area_firearea);
                                        if (Convert.ToDouble(area_firearea).Equals(Math.Round(Convert.ToDouble(fire_area), 2)))
                                        {
                                            stringBuilder_pass.AppendLine(level_name + "，防火分区Id：" + area.Id + "，名字：" + area.Name);// 添加不符合的防火分区
                                            break;
                                        }
                                    }
                                }
                            }
                            */
                        }

                        /*
                        if (exist_FirePipe == true)
                        {
                            stringBuilder_pass.Append(level_name + "，有喷淋系统：" + pass_FirePipe + "最大防火分区面积：3000㎡\n");
                        }
                        else
                        {
                            stringBuilder_notpass.Append(level_name + "，无喷淋系统" + "，最大防火分区面积：1500㎡\n");
                            H00047_pass = false;
                        }
                        */
                    }
                    stringBuilder.Append("不符合的防火分区：\n" + stringBuilder_notpass);
                    stringBuilder.Append("符合的防火分区：\n" + stringBuilder_pass);

                    //stringBuilder.Append("有喷淋的楼层：\n" + stringBuilder_pass);
                    //stringBuilder.Append("无喷淋的楼层：\n" + stringBuilder_notpass);
                }


                
                if (build_height <= 27) //低于27m的建筑，不用判断
                {
                    foreach (Level litem in build_upfloorsNumList) //从这里开始遍历建筑的楼层
                    {
                        //levelNum++;
                        //stringBuilder.AppendLine("litem.Name:" + litem.Name + "，litem.Id:" + litem.Id);
                        pass_FirePipe = "";
                        level_name = litem.Name;
                        while (level_name.Contains("#"))
                        {
                            level_name = SubBetweenString(level_name, "#", "F", false, true, true);
                        }
                        level_name_num = Convert.ToInt16(StartSubString(level_name, 0, "F"));
                        //判断消防立管
                        bool exist_FirePipe = false;
                        //判断消防立管
                        if (litem.Name.Contains("避难层"))
                        {
                            level_height = 56.4;
                        }
                        else
                        {
                            //利用楼层的“名称”属性获得高度，然后将String转成Double，就是真正的楼层对地高度
                            //获得小数点"."后面的一个字符
                            int point_idx = litem.Name.IndexOf(".");
                            string nextpoint = litem.Name.Substring(point_idx + 3, 2);//获得小数点后三个起始的字符，作为SubBetweenString的截断字符，长度只需取2
                            level_height = Convert.ToDouble(SubBetweenString(litem.Name, "F", nextpoint, false, false));
                        }
                        foreach (Element FirePipe in FirePipe_buildNum[buildFlag])
                        {
                            BoundingBoxXYZ FirePipeBoundingBoxXYZ = FirePipe.get_BoundingBox(document.ActiveView);
                            XYZ FirePipeXYZMax = FirePipeBoundingBoxXYZ.Max;
                            XYZ FirePipeXYZMin = FirePipeBoundingBoxXYZ.Min;
                            FirePipeMaxX = FirePipeXYZMax.X;
                            FirePipeMinX = FirePipeXYZMin.X;
                            FirePipeMaxY = FirePipeXYZMax.Y;
                            FirePipeMinY = FirePipeXYZMin.Y;
                            FirePipeMaxZ = FirePipeXYZMax.Z;
                            FirePipeMinZ = FirePipeXYZMin.Z;
                            FirePipeZLength = Math.Round(inch_to_metre(Math.Abs(FirePipeMaxZ - FirePipeMinZ)), 2);
                            if ((FirePipeMinZ <= level_height && FirePipeMaxZ >= level_height) || (FirePipeMinZ > level_height && FirePipeMaxZ < level_height))//这样就认为这一层里有消防立管
                            {
                                pass_FirePipe += FirePipe.Id + "，";
                                exist_FirePipe = true;
                            }
                        }
                        //判断防火分区
                        foreach (SpatialElementTag Tag in FireArea_buildNum[buildFlag])//遍历当前检测到这栋楼有的所有防火分区，其中一层可能包含多个防火分区
                        {
                            string view_name = Tag.View.Name;
                            string fire_area = StartSubString(Tag.TagText, 0, "m²", false);// 得到这种字段：“114514m²”
                            //stringBuilder.AppendLine(Tag.Id.ToString() + "，" + fire_area + "，标签名：" + Tag.TagText + "\n所属视图：" + view_name);
                            //获取该防火分区在哪些楼
                            string[] build_num_part = view_name.Split(new char[] { '、' }); //用“、”分割长名字
                            string level_num = "未知楼层？";
                            //获取该防火分区在哪一层，获取该层的楼层号
                            for (int ii = 0; ii < build_num_part.Length; ii++) //使用for循环输出
                            {
                                if (build_num_part[ii].Contains("B") || build_num_part[ii].Contains("Y") || build_num_part[ii].Contains("S"))
                                {
                                    level_num = StartSubString(build_num_part[ii], 0, "-");
                                }
                                else
                                {
                                    level_num = StartSubString(LastSubEndString(build_num_part[ii], "#"), 0, "F");
                                }
                            }
                            //stringBuilder.AppendLine("所属楼层：" + level_num + "，当前楼层：" + StartSubString(level_name, 0, "F"));
                            if (level_num.Equals(StartSubString(level_name, 0, "F")))//如果该防火分区的所属楼层等于当前检测到的楼层，则
                            {
                                //stringBuilder.AppendLine(level_name + "，" + litem.Id + "，防火分区Id：" + Tag.Id + "，" + fire_area);
                                if ((exist_FirePipe == true && Convert.ToDouble(fire_area) > 5000) || (exist_FirePipe == false && Convert.ToDouble(fire_area) > 2500)) //高度小于等于27m，有喷淋，面积不超过5000m^2，无喷淋，面积不得超过2500m^2
                                {
                                    //运用面积比较法，将Tag上标注的面积和所有防火分区面积对比，如果两者相同则说明是Tag对应的防火分区，将该不符合的防火分区的Id和名字输出
                                    foreach (Area area in AreaCollector)
                                    {
                                        string area_firearea = area.LookupParameter("面积").AsValueString();
                                        //stringBuilder.AppendLine("Id：" + area.Id + "，名字：" + area.Name + "，面积：" + area_firearea);
                                        if (Convert.ToDouble(area_firearea).Equals(Math.Round(Convert.ToDouble(fire_area), 2)))
                                        {
                                            stringBuilder_notpass.AppendLine(level_name + "，防火分区Id：" + area.Id + "，名字：" + area.Name);// 添加不符合的防火分区
                                            break;
                                        }
                                    }
                                    H00047_pass = false;
                                }
                                else
                                {
                                    //运用面积比较法，将Tag上标注的面积和所有防火分区面积对比，如果两者相同则说明是Tag对应的防火分区，将该不符合的防火分区的Id和名字输出
                                    foreach (Area area in AreaCollector)
                                    {
                                        string area_firearea = area.LookupParameter("面积").AsValueString();
                                        //stringBuilder.AppendLine("Id：" + area.Id + "，名字：" + area.Name + "，面积：" + area_firearea);
                                        if (Convert.ToDouble(area_firearea).Equals(Math.Round(Convert.ToDouble(fire_area), 2)))
                                        {
                                            stringBuilder_pass.AppendLine(level_name + "，防火分区Id：" + area.Id + "，名字：" + area.Name);// 添加不符合的防火分区
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        /*
                        if (exist_FirePipe == true)
                        {
                            stringBuilder_pass.Append(level_name + "，有喷淋系统：" + pass_FirePipe + "最大防火分区面积：5000㎡\n");
                        }
                        else
                        {
                            stringBuilder_notpass.Append(level_name + "，无喷淋系统" + "，最大防火分区面积：2500㎡\n");
                            
                        }
                        */
                    }

                    stringBuilder.Append("不符合的防火分区：\n" + stringBuilder_notpass);
                    stringBuilder.Append("符合的防火分区：\n" + stringBuilder_pass);
                    //stringBuilder.Append("符合要求的楼层：\n" + stringBuilder_pass);
                    //stringBuilder.Append("不符合要求的楼层：\n" + stringBuilder_notpass);
                    
                }
                else if (build_height > 27 && build_height <= 54) //27-54m的建筑
                {
                    foreach (Level litem in build_upfloorsNumList) //从这里开始遍历建筑的楼层
                    {
                        //levelNum++;
                        //stringBuilder.AppendLine("litem.Name:" + litem.Name + "，litem.Id:" + litem.Id);
                        pass_FirePipe = "";
                        level_name = litem.Name;
                        while (level_name.Contains("#"))
                        {
                            level_name = SubBetweenString(level_name, "#", "F", false, true, true);
                        }
                        level_name_num = Convert.ToInt16(StartSubString(level_name, 0, "F"));
                        //判断消防立管
                        bool exist_FirePipe = false;
                        //判断消防立管
                        if (litem.Name.Contains("避难层"))
                        {
                            level_height = 56.4;
                        }
                        else
                        {
                            //利用楼层的“名称”属性获得高度，然后将String转成Double，就是真正的楼层对地高度
                            //获得小数点"."后面的一个字符
                            int point_idx = litem.Name.IndexOf(".");
                            string nextpoint = litem.Name.Substring(point_idx + 3, 2);//获得小数点后三个起始的字符，作为SubBetweenString的截断字符，长度只需取2
                            level_height = Convert.ToDouble(SubBetweenString(litem.Name, "F", nextpoint, false, false));
                        }
                        foreach (Element FirePipe in FirePipe_buildNum[buildFlag])
                        {
                            BoundingBoxXYZ FirePipeBoundingBoxXYZ = FirePipe.get_BoundingBox(document.ActiveView);
                            XYZ FirePipeXYZMax = FirePipeBoundingBoxXYZ.Max;
                            XYZ FirePipeXYZMin = FirePipeBoundingBoxXYZ.Min;
                            FirePipeMaxX = FirePipeXYZMax.X;
                            FirePipeMinX = FirePipeXYZMin.X;
                            FirePipeMaxY = FirePipeXYZMax.Y;
                            FirePipeMinY = FirePipeXYZMin.Y;
                            FirePipeMaxZ = FirePipeXYZMax.Z;
                            FirePipeMinZ = FirePipeXYZMin.Z;
                            FirePipeZLength = Math.Round(inch_to_metre(Math.Abs(FirePipeMaxZ - FirePipeMinZ)), 2);
                            if ((FirePipeMinZ <= level_height && FirePipeMaxZ >= level_height) || (FirePipeMinZ > level_height && FirePipeMaxZ < level_height))//这样就认为这一层里有消防立管
                            {
                                pass_FirePipe += FirePipe.Id + "，";
                                exist_FirePipe = true;
                            }
                        }


                        //判断防火分区
                        foreach (SpatialElementTag Tag in FireArea_buildNum[buildFlag])//遍历当前检测到这栋楼有的所有防火分区，其中一层可能包含多个防火分区
                        {
                            string view_name = Tag.View.Name;
                            string fire_area = StartSubString(Tag.TagText, 0, "m²", false);// 得到这种字段：“114514m²”
                            //stringBuilder.AppendLine(Tag.Id.ToString() + "，" + fire_area + "，标签名：" + Tag.TagText + "\n所属视图：" + view_name);
                            //获取该防火分区在哪些楼
                            string[] build_num_part = view_name.Split(new char[] { '、' }); //用“、”分割长名字
                            string level_num = "未知楼层？";
                            //获取该防火分区在哪一层，获取该层的楼层号
                            for (int ii = 0; ii < build_num_part.Length; ii++) //使用for循环输出
                            {
                                if (build_num_part[ii].Contains("B") || build_num_part[ii].Contains("Y") || build_num_part[ii].Contains("S"))
                                {
                                    level_num = StartSubString(build_num_part[ii], 0, "-");
                                }
                                else
                                {
                                    level_num = StartSubString(LastSubEndString(build_num_part[ii], "#"), 0, "F");
                                }
                            }
                            //stringBuilder.AppendLine("所属楼层：" + level_num + "，当前楼层：" + StartSubString(level_name, 0, "F"));
                            if (level_num.Equals(StartSubString(level_name, 0, "F")))//如果该防火分区的所属楼层等于当前检测到的楼层，则
                            {
                                //stringBuilder.AppendLine(level_name + "，" + litem.Id + "，防火分区Id：" + Tag.Id + "，" + fire_area);
                                if ((exist_FirePipe == true && Convert.ToDouble(fire_area) > 3000) || (exist_FirePipe == false && Convert.ToDouble(fire_area) > 1500)) //高度大于27m，有喷淋，面积不超过3000m^2，无喷淋，面积不得超过1500m^2
                                {
                                    //运用面积比较法，将Tag上标注的面积和所有防火分区面积对比，如果两者相同则说明是Tag对应的防火分区，将该不符合的防火分区的Id和名字输出
                                    foreach (Area area in AreaCollector)
                                    {
                                        string area_firearea = area.LookupParameter("面积").AsValueString();
                                        //stringBuilder.AppendLine("Id：" + area.Id + "，名字：" + area.Name + "，面积：" + area_firearea);
                                        if (Convert.ToDouble(area_firearea).Equals(Math.Round(Convert.ToDouble(fire_area), 2)))
                                        {
                                            stringBuilder_notpass.AppendLine(level_name + "，防火分区Id：" + area.Id + "，名字：" + area.Name);// 添加不符合的防火分区
                                            break;
                                        }
                                    }
                                    H00047_pass = false;
                                }
                            }
                            else
                            {
                                //运用面积比较法，将Tag上标注的面积和所有防火分区面积对比，如果两者相同则说明是Tag对应的防火分区，将该不符合的防火分区的Id和名字输出
                                foreach (Area area in AreaCollector)
                                {
                                    string area_firearea = area.LookupParameter("面积").AsValueString();
                                    //stringBuilder.AppendLine("Id：" + area.Id + "，名字：" + area.Name + "，面积：" + area_firearea);
                                    if (Convert.ToDouble(area_firearea).Equals(Math.Round(Convert.ToDouble(fire_area), 2)))
                                    {
                                        stringBuilder_pass.AppendLine(level_name + "，防火分区Id：" + area.Id + "，名字：" + area.Name);// 添加不符合的防火分区
                                        break;
                                    }
                                }
                            }
                        }
                        /*
                        if (exist_FirePipe == true)
                        {
                            stringBuilder_pass.Append(level_name + "，有喷淋系统：" + pass_FirePipe +"最大防火分区面积：3000㎡\n");
                        }
                        else
                        {
                            stringBuilder_notpass.Append(level_name + "，无喷淋系统" + "，最大防火分区面积：1500㎡\n");
                        }
                        */
                    }

                    stringBuilder.Append("不符合的防火分区：\n" + stringBuilder_notpass);
                    stringBuilder.Append("符合的防火分区：\n" + stringBuilder_pass);
                    //stringBuilder.Append("符合的楼层：\n" + stringBuilder_pass);
                    //stringBuilder.Append("不符合的楼层：\n" + stringBuilder_notpass);
                }
                else if (build_height > 54) //大于54m的建筑
                {
                    foreach (Level litem in build_upfloorsNumList) //从这里开始遍历建筑的楼层
                    {
                        //levelNum++;
                        //stringBuilder.AppendLine("litem.Name:" + litem.Name + "，litem.Id:" + litem.Id);
                        pass_FirePipe = "";
                        level_name = litem.Name;
                        while (level_name.Contains("#"))
                        {
                            level_name = SubBetweenString(level_name, "#", "F", false, true, true);
                        }
                        level_name_num = Convert.ToInt16(StartSubString(level_name, 0, "F"));
                        //判断消防立管
                        bool exist_FirePipe = false;
                        //判断消防立管
                        if (litem.Name.Contains("避难层"))
                        {
                            level_height = 56.4;
                        }
                        else
                        {
                            //利用楼层的“名称”属性获得高度，然后将String转成Double，就是真正的楼层对地高度
                            //获得小数点"."后面的一个字符
                            int point_idx = litem.Name.IndexOf(".");
                            string nextpoint = litem.Name.Substring(point_idx + 3, 2);//获得小数点后三个起始的字符，作为SubBetweenString的截断字符，长度只需取2
                            level_height = Convert.ToDouble(SubBetweenString(litem.Name, "F", nextpoint, false, false));
                        }
                        foreach (Element FirePipe in FirePipe_buildNum[buildFlag])
                        {
                            BoundingBoxXYZ FirePipeBoundingBoxXYZ = FirePipe.get_BoundingBox(document.ActiveView);
                            XYZ FirePipeXYZMax = FirePipeBoundingBoxXYZ.Max;
                            XYZ FirePipeXYZMin = FirePipeBoundingBoxXYZ.Min;
                            FirePipeMaxX = FirePipeXYZMax.X;
                            FirePipeMinX = FirePipeXYZMin.X;
                            FirePipeMaxY = FirePipeXYZMax.Y;
                            FirePipeMinY = FirePipeXYZMin.Y;
                            FirePipeMaxZ = FirePipeXYZMax.Z;
                            FirePipeMinZ = FirePipeXYZMin.Z;
                            FirePipeZLength = Math.Round(inch_to_metre(Math.Abs(FirePipeMaxZ - FirePipeMinZ)), 2);

                            if ((FirePipeMinZ <= level_height && FirePipeMaxZ >= level_height)||(FirePipeMinZ > level_height && FirePipeMaxZ < level_height))//这样就认为这一层里有消防立管
                            {
                                pass_FirePipe += FirePipe.Id + "，";
                                exist_FirePipe = true;
                            }
                        }

                        //判断防火分区
                        foreach (SpatialElementTag Tag in FireArea_buildNum[buildFlag])//遍历当前检测到这栋楼有的所有防火分区，其中一层可能包含多个防火分区
                        {
                            string view_name = Tag.View.Name;
                            string fire_area = StartSubString(Tag.TagText, 0, "m²", false);// 得到这种字段：“114514m²”
                            //stringBuilder.AppendLine(Tag.Id.ToString() + "，" + fire_area + "，标签名：" + Tag.TagText + "\n所属视图：" + view_name);
                            //获取该防火分区在哪些楼
                            string[] build_num_part = view_name.Split(new char[] { '、' }); //用“、”分割长名字
                            string level_num = "未知楼层？";
                            //获取该防火分区在哪一层，获取该层的楼层号
                            for (int ii = 0; ii < build_num_part.Length; ii++) //使用for循环输出
                            {
                                if (build_num_part[ii].Contains("B") || build_num_part[ii].Contains("Y") || build_num_part[ii].Contains("S"))
                                {
                                    level_num = StartSubString(build_num_part[ii], 0, "-");
                                }
                                else
                                {
                                    level_num = StartSubString(LastSubEndString(build_num_part[ii], "#"), 0, "F");
                                }
                            }
                            //stringBuilder.AppendLine("所属楼层：" + level_num + "，当前楼层：" + StartSubString(level_name, 0, "F"));
                            if (level_num.Equals(StartSubString(level_name, 0, "F")))//如果该防火分区的所属楼层等于当前检测到的楼层，则
                            {
                                //stringBuilder.AppendLine(level_name + "，" + litem.Id + "，防火分区Id：" + Tag.Id + "，" +fire_area);
                                if ((exist_FirePipe == true && Convert.ToDouble(fire_area) > 3000) || (exist_FirePipe == false && Convert.ToDouble(fire_area) > 1500)) //高度大于27m，有喷淋，面积不超过3000m^2，无喷淋，面积不得超过1500m^2
                                {
                                    //运用面积比较法，将Tag上标注的面积和所有防火分区面积对比，如果两者相同则说明是Tag对应的防火分区，将该不符合的防火分区的Id和名字输出
                                    foreach (Area area in AreaCollector)
                                    {
                                        string area_firearea = area.LookupParameter("面积").AsValueString();
                                        //stringBuilder.AppendLine("Id：" + area.Id + "，名字：" + area.Name + "，面积：" + area_firearea);
                                        if (Convert.ToDouble(area_firearea).Equals(Math.Round(Convert.ToDouble(fire_area), 2)))
                                        {
                                            stringBuilder_notpass.AppendLine(level_name + "，防火分区Id：" + area.Id + "，名字：" + area.Name);// 添加不符合的防火分区
                                            break;
                                        }
                                    }
                                    H00047_pass = false;
                                }
                                else
                                {
                                    //运用面积比较法，将Tag上标注的面积和所有防火分区面积对比，如果两者相同则说明是Tag对应的防火分区，将该不符合的防火分区的Id和名字输出
                                    foreach (Area area in AreaCollector)
                                    {
                                        string area_firearea = area.LookupParameter("面积").AsValueString();
                                        //stringBuilder.AppendLine("Id：" + area.Id + "，名字：" + area.Name + "，面积：" + area_firearea);
                                        if (Convert.ToDouble(area_firearea).Equals(Math.Round(Convert.ToDouble(fire_area), 2)))
                                        {
                                            stringBuilder_pass.AppendLine(level_name + "，防火分区Id：" + area.Id + "，名字：" + area.Name);// 添加不符合的防火分区
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        
                        /*
                        if (exist_FirePipe == true)
                        {
                            stringBuilder_pass.Append(level_name + "，有喷淋系统：" + pass_FirePipe + "最大防火分区面积：3000㎡\n");
                        }
                        else
                        {
                            stringBuilder_notpass.Append(level_name + "，无喷淋系统" + "，最大防火分区面积：1500㎡\n");
                            H00047_pass = false;
                        }
                        */
                    }
                    stringBuilder.Append("不符合的防火分区：\n" + stringBuilder_notpass);
                    stringBuilder.Append("符合的防火分区：\n" + stringBuilder_pass);

                    //stringBuilder.Append("有喷淋的楼层：\n" + stringBuilder_pass);
                    //stringBuilder.Append("无喷淋的楼层：\n" + stringBuilder_notpass);
                }
            }




            if (H00047_pass == false)
            {
                H00047_result.AppendLine("不符合5.3.1建筑设计防火规范 GB50016-2014（2018年版）");
            }
            else
            {
                H00047_result.AppendLine("符合5.3.1建筑设计防火规范 GB50016-2014（2018年版）");
            }
            stringBuilder = H00047_result.AppendLine(stringBuilder.ToString());



            /*
            //////////////////////切换到3D视图////////////////////////////
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
            /////////////////////////////////////////////////////////
            */

            //PrintLog(stringBuilder.ToString(), "H00047", document, null);
            TaskDialog.Show("H00047强条检测", stringBuilder.ToString());

            return Result.Succeeded;
        }
    }
}
