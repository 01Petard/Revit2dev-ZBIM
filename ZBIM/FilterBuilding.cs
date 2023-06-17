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
    public class FilterBuilding : IExternalCommand
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




















            ////////////////////////////////这一段是对轴网进行过滤、排序//////////////////////////////
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

            //对轴网进行冒泡排序
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
                            string j1_str = StartSubString(elementsGridsSort[jj].Name, 0, "-");
                            string j2_str = StartSubString(elementsGridsSort[jj + 1].Name, 0, "-");
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
                        catch (Exception)//这里会输出异常的轴网，也就是无法排序的轴网
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
            ////////////////////////////////过滤完成//////////////////////////////














            ////////////////////////////////这一步是获得每一个轴网的BoundingXYZ属性，通过对比获得每栋建筑的X、Y坐标//////////////////////////////
            double buildMinX = 0; //建筑最小的X坐标
            double buildMinY = 0; //建筑最小的Y坐标
            double buildMaxX = 0; //建筑最大的X坐标
            double buildMaxY = 0; //建筑最大的Y坐标
            //double buildMaxZ = 0; //建筑最大的Y坐标

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
                    !item.Name.Contains("B") &&
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
                        stringBuilder.AppendLine("轴网：" + item_name + "的方位：" + "X:(" + buildMinX + "," + buildMaxX + ")" + "，Y:(" + buildMinY + "," + buildMaxY + ")");
                    }
                    else //如果前后两个轴网发生变化，说明已经切换到下一栋楼了，打印输出一下这栋楼的占地方位
                    {
                        //轴网变化，新的建筑占地坐标重新更新，此时进行物品的检测最准确
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
            ////////////////////////////////以上是过滤轴网，得到每栋楼的X、Y坐标//////////////////////////////


























            ////////////////////////////////以下是遍历每栋楼的每一层//////////////////////////////
            List<Element> elementsLevels = collectorLevels.OfCategory(BuiltInCategory.OST_Levels).ToList<Element>();
            List<Element> elementsAreas = collectorAreas.OfCategory(BuiltInCategory.OST_Areas).ToList<Element>();
            List<Element> elementsRooms = collectorRooms.OfCategory(BuiltInCategory.OST_Rooms).ToList<Element>();
            List<Element> elementsViews = collectorViews.OfCategory(BuiltInCategory.OST_Views).ToList<Element>();

            IList<Element> viewSchedules = collectorViewSchedules.OfClass(typeof(ViewSchedule)).ToElements();
            IList<Element> viewSections = collectorViewSections.OfClass(typeof(ViewSection)).ToElements();
            IList<Element> viewPlans = collectorViewPlans.OfClass(typeof(ViewPlan)).ToElements();

            IList<ElementId> elementGroupsListId = new List<ElementId>();
            IList<ElementId> elementGridsListId = new List<ElementId>();

            //FilteredElementCollector levelid = new FilteredElementCollector(document);   //楼层的过滤器
            //ICollection<ElementId> levelIds = levelid.OfClass(typeof(Level)).ToElementIds();  //过滤获得所有楼层的id
            List<ElementId> LevelIds = new List<ElementId>();   //存放每一个楼层的id 

            Boolean H00018_pass = true;
            StringBuilder H00018_result = new StringBuilder();



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
            IList<String> buildNumList = new List<String>();//存放楼号的列表

            //stringBuilder.AppendLine("过滤模型组轴网结构");
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

            int i = -1;

            Level outdoor_floor = null;

            //地上建筑，从这里开始遍历楼栋
            foreach (string buildFlag in buildNumList)
            {
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




                //////////////////////////////这一步我对build_upfloorsNumList进行了排序//////////////////////////////
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


                //下面在计算楼栋高度
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
                //stringBuilder.Append("楼栋编号：" + buildFlag + "#，建筑高度：" + Math.Round(build_height, 2) + "m");
                stringBuilder.Append(buildFlag + "# 建筑高度：" + Math.Round(build_height, 2) + "m");
                //stringBuilder.Append(buildFlag + "建筑高度");




                //////////////////////////////下面可以针对不同的建筑高度的楼，对他们的每一层进行操作//////////////////////////////
                if (build_height < 21)
                {
                    stringBuilder.AppendLine("（≤21m）");
                }
                else if (build_height >= 21 && build_height < 27)
                {
                    //stringBuilder.AppendLine("（＞21m且≤27m）");
                    stringBuilder.AppendLine("（21-27m）");
                }
                else
                {
                    stringBuilder.AppendLine("（＞27m）");
                }



                //对高度不同的楼，遍历楼的每一层
                if (build_height < 21) //低于21m的建筑
                {
                    foreach (Level litem in build_upfloorsNumList) //从这里开始遍历建筑的楼层
                    {

                    }
                }
                else if (build_height > 21 && build_height <= 27) //21-27m的建筑
                {
                    foreach (Level litem in build_upfloorsNumList) //从这里开始遍历建筑的楼层
                    {
                        
                    }
                }
                else if (build_height > 27) //大于27m的建筑
                {
                    foreach (Level litem in build_upfloorsNumList)//从这里开始遍历建筑的楼层
                    {

                    }
                }



            }

            //////////////////////////////完成//////////////////////////////






            if (H00018_pass == false)
            {
                H00018_result.AppendLine("检测结论：不符合8.2.1建筑设计防火规范 GB50016-2014（2018年版）");
            }
            else
            {
                H00018_result.AppendLine("检测结论：符合8.2.1建筑设计防火规范 GB50016-2014（2018年版）");
            }
            stringBuilder = H00018_result.AppendLine(stringBuilder.ToString());



            PrintLog(stringBuilder.ToString(), "H00018", document, null);
            TaskDialog.Show("H00018强条检测", stringBuilder.ToString());


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


            return Result.Succeeded;
        }
    }
}
