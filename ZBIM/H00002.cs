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
using Point = System.Drawing.Point;


namespace H00002
{
    [Transaction(TransactionMode.Manual)]
    public class H00002 : IExternalCommand
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
            int buildNum;

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

            //Element group_element;
            //int buildNum;

            //stringBuilder.AppendLine("过滤模型组梁结构");

            //foreach (var item in elementsGroups)
            //{
            //    if (item.Name.Contains("#") &&
            //        item.Name.Contains("梁"))
            //    {
            //        buildNum = int.Parse(StartSubString(item.Name, 0, "#", false, false));
            //        if (!buildNumList.Contains(buildNum))
            //        {
            //            buildNumList.Add(buildNum);

            //            elementGroupsList.Add(item);
            //            group_element = document.GetElement(new ElementId(item.Id.IntegerValue));
            //            elementGroupsListId.Add(group_element.Id);

            //            //stringBuilder.AppendLine(item.Id + " " + item.Name + " " + item.LevelId + " " + item.GetType());

            //            //uiDocument.Selection.SetElementIds(elementGroupsListId);
            //        }
            //    }
            //}

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
                                    !up_levelname.Contains("Y"))
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

                                            //foreach (Edge geoEdge in geomSolid.Edges)
                                            //{
                                            //    EdgeLength += geoEdge.ApproximateLength;
                                            //    //得到墙边
                                            //}
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
                    ////得到编辑框
                    //BoundingBoxXYZ buildboundingBoxXYZ = group_element.get_BoundingBox(document.ActiveView);
                    ////通过边界框创建一个OutLine
                    //XYZ buildmin = buildboundingBoxXYZ.Min;
                    //XYZ buildmax = buildboundingBoxXYZ.Max;

                    string ok = null;
                    string not_ok = null;
                    string no_name = null;

                    if (build_height > 33)//仅地上建筑
                    {
                        //stringBuilder.AppendLine("!!!该建筑地上部分应该设置消防电梯");
                        //stringBuilder.AppendLine("建筑高度>33m");
                        //stringBuilder.AppendLine("地下室深度≤10m");
                        //stringBuilder.AppendLine("判断地上建筑");

                        ok = null;
                        not_ok = null;
                        no_name = null;

                        bool up_room_flag = false;

                        foreach (var litem in build_upfloorsNumList)
                        {
                            bool viewPlan_flag = false;

                            foreach (ViewPlan viewPlan in viewPlans)
                            {
                                //stringBuilder.AppendLine("所有楼层：" + viewPlan.GetType() + " "+viewPlan.Id + " " + viewPlan.Name);
                                if (viewPlan.Name.Equals(litem.Name) &&
                                    !up_room_flag)
                                {
                                    viewPlan_flag = true;
                                    collectorRooms = new FilteredElementCollector(document, viewPlan.Id);
                                    elementsRooms = collectorRooms.OfCategory(BuiltInCategory.OST_Rooms).ToList<Element>();

                                    double[] buildboundingBoxXY = get_Floor_Line_BBox(buildFlag.ToString(), viewPlan, elementsGrids);

                                    //stringBuilder.AppendLine("轴网坐标：" +buildFlag.ToString() + buildboundingBoxXY[0] + " " + buildboundingBoxXY[1] + " " + buildboundingBoxXY[2] + " " + buildboundingBoxXY[3]);

                                    foreach (var roomitem in elementsRooms)
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
                                                up_room_flag = true;
                                                //stringBuilder.AppendLine("已设置消防电梯的楼层：" + litem.Name + " " + roomitem.Id + " " + roomitem.Name);
                                                string temp_string = litem.Name;
                                                while (temp_string.Contains("#"))
                                                {
                                                    temp_string = SubBetweenString(temp_string, "#", "F", false, true, true) + "  ";
                                                }
                                                ok = ok + temp_string;
                                            }
                                        }
                                    }
                                }
                            }
                            if (viewPlan_flag == true &&
                                up_room_flag == false)
                            {
                                //stringBuilder.AppendLine("未设置消防电梯的楼层：" + litem.Name);
                                string temp_string = litem.Name;
                                while (temp_string.Contains("#"))
                                {
                                    temp_string = SubBetweenString(temp_string, "#", "F", false, true, true) + "  ";
                                }
                                not_ok = not_ok + temp_string;
                            }

                            if (viewPlan_flag == false)
                            {
                                //stringBuilder.AppendLine("找不到该楼层平面图：" + litem.Name);
                                string temp_string = litem.Name;
                                while (temp_string.Contains("#"))
                                {
                                    temp_string = SubBetweenString(temp_string, "#", "F", false, true, true) + "  ";
                                }
                                no_name = no_name + temp_string;
                            }

                            up_room_flag = false;
                        }

                        //stringBuilder.AppendLine("未设置消防电梯的楼层：" + not_ok);
                        //stringBuilder.AppendLine("已设置消防电梯的楼层：" + ok);
                        //stringBuilder.AppendLine("找不到其平面图的楼层：" + no_name);

                        if (ok != null && not_ok == null && no_name == null)
                        {
                            fit.Add("楼栋编号：" + buildFlag + "#");
                            fit.Add("建筑高度>33m");
                            fit.Add("未设置消防电梯的楼层：" + not_ok);
                            fit.Add("已设置消防电梯的楼层：" + ok);
                            fit.Add("找不到其平面图的楼层：" + no_name);
                            fit.Add("");
                            //stringBuilder.AppendLine("！！！符合建筑设计防火规范GB50016-2014（2018年版）");
                        }
                        else
                        {
                            not_fit.Add("楼栋编号：" + buildFlag + "#");
                            not_fit.Add("建筑高度>33m");
                            not_fit.Add("未设置消防电梯的楼层：" + not_ok);
                            not_fit.Add("已设置消防电梯的楼层：" + ok);
                            not_fit.Add("找不到其平面图的楼层：" + no_name);
                            not_fit.Add("");
                            //stringBuilder.AppendLine("---不符合建筑设计防火规范GB50016-2014（2018年版）");
                        }
                    }
                    else
                    {
                        //stringBuilder.AppendLine("---该建筑不用设置消防电梯");
                        //stringBuilder.AppendLine("建筑高度≤33m");
                        //stringBuilder.AppendLine("地下室深度≤10m");
                        fit.Add("楼栋编号：" + buildFlag + "#");
                        fit.Add("建筑高度≤33m");
                        fit.Add("消防电梯不作检测");
                        fit.Add("");
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
                ////得到编辑框
                //BoundingBoxXYZ buildboundingBoxXYZ = group_element.get_BoundingBox(document.ActiveView);
                ////通过边界框创建一个OutLine
                //XYZ buildmin = buildboundingBoxXYZ.Min;
                //XYZ buildmax = buildboundingBoxXYZ.Max;

                string ok = null;
                string not_ok = null;
                string no_name = null;

                down = 0;

                if ((build_depth > 10 && build_basement_area > 3000)|| buildNumList != null)//仅地下建筑
                {
                    //stringBuilder.AppendLine("!!!该建筑地下部分应该设置消防电梯");
                    //stringBuilder.AppendLine("地下室深度>10m且面积>3000平方米");
                    //stringBuilder.AppendLine("需要判断地下建筑");

                    ok = null;
                    not_ok = null;
                    no_name = null;

                    bool down_room_flag = false;

                    foreach (var litem in build_downfloorsNumList)
                    {
                        bool viewPlan_flag = false;

                        foreach (ViewPlan viewPlan in viewPlans)
                        {
                            //stringBuilder.AppendLine("所有楼层：" + viewPlan.GetType() + " "+viewPlan.Id + " " + viewPlan.Name);
                            if (viewPlan.Name.Equals(litem.Name) &&
                                !down_room_flag)
                            {
                                viewPlan_flag = true;
                                collectorRooms = new FilteredElementCollector(document, viewPlan.Id);
                                elementsRooms = collectorRooms.OfCategory(BuiltInCategory.OST_Rooms).ToList<Element>();

                                double[] buildboundingBoxXY = get_Floor_Line_BBox("d", viewPlan, elementsGrids);

                                foreach (var roomitem in elementsRooms)
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
                                            down_room_flag = true;
                                            //stringBuilder.AppendLine("已设置消防电梯的楼层：" + litem.Name + " " + roomitem.Id + " " + roomitem.Name);
                                            ok = ok + SubBetweenString(litem.Name, "B", "-", true, false, true) + "  ";
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        if (viewPlan_flag == true &&
                            down_room_flag == false)
                        {
                            //stringBuilder.AppendLine("未设置消防电梯的楼层：" + litem.Name);
                            not_ok = not_ok + SubBetweenString(litem.Name, "B", "-", true, false, true) + "  ";
                        }

                        if (viewPlan_flag == false)
                        {
                            //stringBuilder.AppendLine("找不到该楼层平面图：" + litem.Name);
                            no_name = no_name + SubBetweenString(litem.Name, "B", "-", true, false, true) + "  ";
                        }

                        down_room_flag = false;
                    }

                    //stringBuilder.AppendLine("未设置消防电梯的楼层：" + not_ok);
                    //stringBuilder.AppendLine("已设置消防电梯的楼层：" + ok);
                    //stringBuilder.AppendLine("找不到其平面图的楼层：" + no_name);

                    if (ok != null && not_ok == null && no_name == null)
                    {
                        fit.Add("楼栋编号：" + "地下室");
                        fit.Add("深埋>10m且总建筑面积>3000平方米");
                        fit.Add("未设置消防电梯的楼层：" + not_ok);
                        fit.Add("已设置消防电梯的楼层：" + ok);
                        fit.Add("找不到其平面图的楼层：" + no_name);
                        fit.Add("");
                        //stringBuilder.AppendLine("！！！符合建筑设计防火规范GB50016-2014（2018年版）");
                    }
                    else
                    {
                        not_fit.Add("楼栋编号：" + "地下室");
                        //not_fit.Add("深埋>10m且总建筑面积>3000平方米");
                        not_fit.Add("未设置消防电梯的楼层：" + not_ok);
                        not_fit.Add("已设置消防电梯的楼层：" + ok);
                        not_fit.Add("找不到其平面图的楼层：" + no_name);
                        not_fit.Add("");
                        //stringBuilder.AppendLine("---不符合建筑设计防火规范GB50016-2014（2018年版）");
                    }
                }
                else
                {
                    //stringBuilder.AppendLine("---该建筑不用设置消防电梯");
                    //stringBuilder.AppendLine("建筑高度≤33m");
                    //stringBuilder.AppendLine("地下室深度≤10m");
                    fit.Add("楼栋编号：" + "地下室");
                    //fit.Add("深埋≤10m或总建筑面积≤3000平方米");
                    fit.Add("无需设置消防电梯");
                    fit.Add("");
                }
            }
            

            //将元素高亮
            //uiDocument.Selection.SetElementIds(elementFloorsListId);


            if (not_fit != null)
            {
                stringBuilder.AppendLine("不符合建筑设计防火规范GB50016-2014（2018年版）");
                stringBuilder.AppendLine();
                foreach (var string_item in fit)
                {
                    stringBuilder.AppendLine(string_item);
                }
                foreach (var string_item in not_fit)
                {
                    stringBuilder.AppendLine(string_item);
                }
            }
            else if (fit != null && not_fit == null)
            {
                stringBuilder.AppendLine("符合建筑设计防火规范GB50016-2014（2018年版）");
                stringBuilder.AppendLine();
                foreach (var string_item in fit)
                {
                    stringBuilder.AppendLine(string_item);
                }
                foreach (var string_item in not_fit)
                {
                    stringBuilder.AppendLine(string_item);
                }
            }
            else
            {
                stringBuilder.AppendLine("未检测到元素");
            }
            
            TaskDialog.Show("提示", stringBuilder.ToString());

            return Result.Succeeded;
        }
    }
}
