using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;


namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]

    public class H00003 : IExternalCommand
    {
        public struct building_message
        {
            public string Building;
            public bool feasible; //若该建筑有对象, 则置为true
            public double[] BBox; //该楼层的BBox
            public List<string> windows_NG;
            public List<string> windows_OK;
            public List<string> rail_OK;
            public List<string> rail_high_NG;
            public List<string> rail_distance_NG;
            public List<List<string>> whole;
        }

        int windows_NG_idx = 0;
        int windows_OK_idx = 1;
        int rail_OK_idx = 2;
        int rail_high_NG_idx = 3;
        int rail_distance_NG_idx = 4;

        public int building_comparison(building_message b1, building_message b2)
        {
            double area1 = (b1.BBox[3] - b1.BBox[1]) * (b1.BBox[2] - b1.BBox[0]);
            double area2 = (b2.BBox[3] - b2.BBox[1]) * (b2.BBox[2] - b2.BBox[0]);
            if (area1 > area2)
                return -1;
            else
                return 1;
        }

        List<String> windows_stair = new List<String>();
        List<building_message> text = new List<building_message>();
        static int mmDec = 2;
        static int mDec = 3;
        ElementCategoryFilter Stair_Filter = new ElementCategoryFilter(BuiltInCategory.OST_Stairs); // 楼梯组过滤器
        ElementCategoryFilter Floor_Filter = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming); // 模型组过滤器
        ElementCategoryFilter StairsRailing_Filter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRailing); // 栏杆过滤器
        ElementCategoryFilter Levels_Filter = new ElementCategoryFilter(BuiltInCategory.OST_Levels); // 楼层过滤器
        ElementCategoryFilter View_Filter = new ElementCategoryFilter(BuiltInCategory.OST_Views); // 视角过滤器
        ElementCategoryFilter Windows_Filter = new ElementCategoryFilter(BuiltInCategory.OST_Windows); // 窗户过滤器
        ElementCategoryFilter Fushou_Filter = new ElementCategoryFilter(BuiltInCategory.OST_ProfileFamilies); // 扶手过滤器
        ElementCategoryFilter Room_Filter = new ElementCategoryFilter(BuiltInCategory.OST_Rooms); // 房间过滤器
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //界面交互的doc
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            //实际内容的doc
            Document doc = commandData.Application.ActiveUIDocument.Document;

            View root_View = doc.ActiveView;

            //收集器
            FilteredElementCollector StairsRailing_collector = new FilteredElementCollector(doc); // 栏杆收集器
            FilteredElementCollector View_collector = new FilteredElementCollector(doc); // 视角收集器
            FilteredElementCollector Windows_collector = new FilteredElementCollector(doc); // 窗户收集器
            FilteredElementCollector Fushou_collector = new FilteredElementCollector(doc); // 扶手收集器
            FilteredElementCollector Room_collector = new FilteredElementCollector(doc); //房间收集器
            FilteredElementCollector Levels_collector = new FilteredElementCollector(doc); // 楼层收集器
            StairsRailing_collector.WherePasses(StairsRailing_Filter);
            Levels_collector.WherePasses(Levels_Filter);
            View_collector.WherePasses(View_Filter).OfClass(typeof(ViewPlan));
            Windows_collector.WherePasses(Windows_Filter);
            Fushou_collector.WherePasses(Fushou_Filter);
            Room_collector.WherePasses(Room_Filter);

            foreach (var item in View_collector)
            {
                if (item != null)
                {
                    View v = (View)item;
                    uiDoc.ActiveView = v;
                    doc = commandData.Application.ActiveUIDocument.Document;
                    break;
                }
            }
            FilteredElementCollector Grid_collector = new FilteredElementCollector(doc); // 轴网收集器
            ElementCategoryFilter Grid_Filter = new ElementCategoryFilter(BuiltInCategory.OST_Grids); // 轴网过滤器
            Grid_collector.WherePasses(Grid_Filter);
            List<Element> grid_list = new List<Element>();
            foreach (var item in Grid_collector)
            {
                if (item.Name.Contains("d-"))
                    continue;
                else
                    grid_list.Add(item);
            }

            // 轴网确定楼栋
            while (grid_list.Count() > 0)
            {
                int i = 0, _index = 0;
                while (i < grid_list[0].Name.Length)
                {
                    if (grid_list[0].Name[i] == '-')
                        break;
                    ++i;
                }
                _index = i;
                if (i == 0 || i == grid_list[0].Name.Length)
                {
                    grid_list.RemoveAt(0);
                    continue;
                }
                string pre_s = "";
                for (int j = 0; j < i; j++)
                    pre_s += grid_list[0].Name[j];
                List<Element> one_grid_list = new List<Element>();
                for (i = 0; i < grid_list.Count(); i++)
                {
                    if (grid_list[i].Name.Contains(pre_s + "-") && grid_list[i].Name.Length > _index && grid_list[i].Name[_index] == '-')
                        one_grid_list.Add(grid_list[i]);
                }

                grid_list.RemoveAll(ss => (ss.Name.Contains(pre_s + "-") && ss.Name.Length > _index && ss.Name[_index] == '-'));
                building_message temp_building = new building_message();
                {
                    //初始化
                    temp_building.Building = pre_s;
                    temp_building.windows_NG = new List<string>();
                    temp_building.windows_OK = new List<string>();
                    temp_building.rail_OK = new List<string>();
                    temp_building.rail_high_NG = new List<string>();
                    temp_building.rail_distance_NG = new List<string>();
                    temp_building.whole = new List<List<string>>();
                    temp_building.whole.Add(temp_building.windows_NG);
                    temp_building.whole.Add(temp_building.windows_OK);
                    temp_building.whole.Add(temp_building.rail_OK);
                    temp_building.whole.Add(temp_building.rail_high_NG);
                    temp_building.whole.Add(temp_building.rail_distance_NG);
                    temp_building.BBox = get_Floor_Line_BBox(one_grid_list, doc);
                }
                text.Add(temp_building);

                //string temp_ss = "";
                //for (i = 0; i < one_grid_list.Count(); i++)
                //{
                //    temp_ss += one_grid_list[i].Name + "  ";
                //}
                //temp_ss += "\n";
                //TaskDialog.Show("Tips", temp_ss);
            }
            text.Sort((x, y) => building_comparison(x, y));

            building_message unk = new building_message();
            {
                //初始化
                unk.Building = "unknown";
                unk.windows_NG = new List<string>();
                unk.windows_OK = new List<string>();
                unk.rail_OK = new List<string>();
                unk.rail_high_NG = new List<string>();
                unk.rail_distance_NG = new List<string>();
                unk.whole = new List<List<string>>();
                unk.whole.Add(unk.windows_NG);
                unk.whole.Add(unk.windows_OK);
                unk.whole.Add(unk.rail_OK);
                unk.whole.Add(unk.rail_high_NG);
                unk.whole.Add(unk.rail_distance_NG);
            }
            text.Add(unk);

            //string temp_s = "";
            //for (int i = 0; i < text.Count(); i++)
            //{
            //    temp_s += text[i].Building.ToString() + "  ";
            //}
            //TaskDialog.Show("Tips", (temp_s).ToString());

            uiDoc.ActiveView = root_View;
            doc = commandData.Application.ActiveUIDocument.Document;

            bool feasible = true;
            bool finalFlag = true;
            foreach (var item in Windows_collector)
            {
                bool SR_Flag = true;
                bool shade_flag = true;
                FamilyInstance Fi = item as FamilyInstance;
                {
                    ParameterMap P = item.ParametersMap;
                    try
                    {
                        P.get_Item("百叶片角度").AsValueString();
                    }
                    catch (Exception e)
                    {
                        shade_flag = false;
                    }
                }
                if (shade_flag == false)
                {
                    if (Fi == null)
                    {
                        // TaskDialog.Show("Tips", "Fi == null");
                        continue;
                    }
                    if ((Fi.FromRoom != null && (Fi.ToRoom == null || Fi.ToRoom.Name.Contains("阳台"))) || ((Fi.FromRoom == null || Fi.FromRoom.Name.Contains("阳台")) && Fi.ToRoom != null))
                    {
                        BoundingBoxXYZ xyz = item.get_BoundingBox(doc.ActiveView);
                        double high = 0;
                        bool high_flag1 = true, high_flag2 = true;

                        ParameterMap P = item.ParametersMap;
                        try
                        {
                            high = Math.Round(0.3048 * P.get_Item("底高度").AsDouble(), mDec);
                        }
                        catch (Exception e)
                        {
                            high_flag1 = false;
                        }

                        try
                        {
                            high = (xyz.Min.Z + xyz.Max.Z - P.get_Item("高度").AsDouble()) / 2;
                            Level L = null;
                            foreach (var item_L in Levels_collector)
                            {
                                if (item_L.Id == item.LevelId)
                                {
                                    L = (Level)item_L;
                                    break;
                                }
                            }
                            high = Math.Round((high - L.Elevation) * 0.3048, mDec);
                        }
                        catch (Exception e)
                        {
                            high_flag2 = false;
                        }

                        if (!high_flag1 && !high_flag2)
                        {
                            // TaskDialog.Show("Tips", "!high_flag1 && !high_flag2");
                            continue;
                        }

                        int floor = 1;
                        feasible = true;
                        //TaskDialog.Show("Tips", (high).ToString());
                        SR_Flag = Windows_Judge_Simple(item, Levels_collector, doc, uiDoc, ref feasible, ref floor);
                        if (item.Name.Contains("固定"))
                        {
                            // TaskDialog.Show("Tips", "item.Name.Contains(固定)");
                            continue;
                        }
                        if (feasible == true)
                        {
                            if (high >= 0.9 || SR_Flag == true)
                            {
                                insertMessage(text, item, windows_OK_idx, doc);
                            }
                            else
                            {
                                insertMessage(text, item, windows_NG_idx, doc);
                                finalFlag = false;
                            }
                        }
                    }
                }
            }


            foreach (var item in StairsRailing_collector)
            {
                bool windows_stair_flag = false;
                for (int i = 0; i < windows_stair.Count(); i++)
                    if (windows_stair[i].ToString().Equals(item.Id.ToString()))
                    {
                        windows_stair_flag = true;
                        break;
                    }
                if (windows_stair_flag != true && !BelongToStairandExist(item, doc))
                {
                    bool glassFlag = true;
                    bool disFlag = true;
                    bool highFlag = true;
                    GeometryElement gElem = item.get_Geometry(new Options());
                    double glassDis = 0;
                    double distance = 0;
                    double beamHigh = 0;
                    double levelHigh = 0;
                    double railHigh = 0;
                    double cmpHigh = 0;
                    int floor = 1;
                    if (judge_material(item, doc) == true)
                    {
                        feasible = true;
                        glassDis = Cal_StairsRailing_Glass(item, Levels_collector, doc, ref beamHigh, ref levelHigh, ref railHigh, ref floor, ref feasible);
                        if (feasible == false)
                            continue;
                        if (Math.Round(glassDis, mmDec) <= 110)
                        {
                            glassFlag = true;
                            //TaskDialog.Show("Tips", "合格, 栏杆下间隙为" + (glassDis).ToString() + "mm");
                        }
                        else
                        {
                            glassFlag = false;
                            //TaskDialog.Show("Tips", "不合格, 栏杆下间隙为" + (glassDis).ToString() + "mm");
                        }
                    }
                    else
                    {
                        feasible = true;
                        Cal_Rail_High(item, Levels_collector, doc, ref beamHigh, ref levelHigh, ref railHigh, ref floor, ref feasible);
                        if (feasible == false)
                            continue;
                        distance = 1000 * Cal_StairsRailing_Distance(gElem);
                        if (Math.Round(distance, mmDec) <= 110)
                        {
                            disFlag = true;
                            //TaskDialog.Show("Tips", "合格, 栏杆间隙为" + (distance).ToString() + "mm");
                        }
                        else
                        {
                            disFlag = false;
                            insertMessage(text, item, rail_distance_NG_idx, doc);
                            finalFlag = false;
                            //TaskDialog.Show("Tips", "不合格, 栏杆间隙为" + (distance).ToString() + "mm");
                        }
                    }

                    if (Math.Round((beamHigh - levelHigh) * 304.8, mmDec) > 450 || beamHigh < levelHigh)
                        cmpHigh = levelHigh;
                    else
                        cmpHigh = beamHigh;
                    if (floor <= 1)
                        continue;
                    else if (Math.Round((railHigh - cmpHigh) * 304.8, mmDec) >= 1100 && floor >= 7)
                        highFlag = true;
                    else if (Math.Round((railHigh - cmpHigh) * 304.8, mmDec) >= 1050 && floor < 7)
                        highFlag = true;
                    else
                        highFlag = false;
                    if (highFlag == false || glassFlag == false)
                    {
                        insertMessage(text, item, rail_high_NG_idx, doc);
                        finalFlag = false;
                    }
                    if (highFlag && glassFlag && disFlag) { }
                    insertMessage(text, item, rail_OK_idx, doc);
                    /*
                    TaskDialog.Show("Tips", "下间隙:" + (glassDis).ToString() + "mm\n" + "栏杆间隙:" + (distance).ToString() + "mm\n" + "栏杆高度:" + ((railHigh - cmpHigh) * 304.8).ToString() + "mm\n" + "楼层:" + floor.ToString() + "\n"
                        + "glassFlag:" + (glassFlag).ToString() + "\n" + "disFlag:" + (disFlag).ToString() + "\n" + "highFlag:" + (highFlag).ToString() + "\n");
                    */
                }
            }

            PrintLog(getMessage(text, finalFlag), "H00003", doc);
            return Result.Succeeded;
        }

        private double Cal_StairsRailing_Distance(GeometryElement gElem, bool center_distance_flag = false) //计算栏杆间距,center_distance_flag=true表示计算中心距
        {
            int maxNum = -1;
            int sizeMaxNum = -1;
            double real_distance = -1;
            double bbox_size = -1;
            foreach (GeometryObject gObj in gElem)
            {
                if (gObj is GeometryInstance)
                {
                    List<XYZ> xyz = new List<XYZ>();
                    List<double> distance_record = new List<double>();
                    List<int> num_record = new List<int>();
                    int rec_len = 0;

                    List<BoundingBoxXYZ> BBox = new List<BoundingBoxXYZ>();
                    List<double> bbox_record = new List<double>();
                    List<int> size_num_record = new List<int>();
                    int rec_len_bbox = 0;

                    int length = 0;
                    String s = "";
                    GeometryInstance gIns = gObj as GeometryInstance;
                    GeometryElement gIns_Elem = gIns.GetInstanceGeometry();
                    foreach (var item0 in gIns_Elem)
                    {
                        //添加到列表
                        if (item0 is Solid)
                        {
                            Solid sd = item0 as Solid;
                            if (sd.Volume != 0)
                            {
                                xyz.Add(sd.ComputeCentroid());
                                if (center_distance_flag == false)
                                    BBox.Add(sd.GetBoundingBox());
                                length++;
                            }
                        }
                    }
                    for (int j = 1; j < length; j++)
                    {
                        //计算相邻点的距离
                        double distance = Math.Sqrt((xyz[j].X - xyz[j - 1].X) * (xyz[j].X - xyz[j - 1].X) + (xyz[j].Y - xyz[j - 1].Y) * (xyz[j].Y - xyz[j - 1].Y));
                        bool flag = false;
                        //统计点距
                        for (int i = 0; i < rec_len; i++)
                        {
                            if (Math.Abs(distance_record[i] - distance) < 0.0001)
                            {
                                num_record[i] += 1;
                                flag = true;
                                break;
                            }
                        }
                        if (flag == false)
                        {
                            num_record.Add(1);
                            distance_record.Add(distance);
                            rec_len++;
                        }
                        //计算BBox的大小
                        double size = 0;
                        bool sizeFlag = false;
                        if (center_distance_flag == false)
                        {
                            if ((BBox[j].Max.X - BBox[j].Min.X) > (BBox[j].Max.Y - BBox[j].Min.Y))
                                size = BBox[j].Max.X - BBox[j].Min.X;
                            else
                                size = BBox[j].Max.Y - BBox[j].Min.Y;
                            //统计BBox大小
                            for (int i = 0; i < rec_len_bbox; i++)
                            {
                                if (Math.Abs(bbox_record[i] - size) < 0.0001)
                                {
                                    size_num_record[i] += 1;
                                    sizeFlag = true;
                                    break;
                                }
                            }
                            if (sizeFlag == false)
                            {
                                size_num_record.Add(1);
                                bbox_record.Add(size);
                                rec_len_bbox++;
                            }
                        }
                    }
                    //查找点距出现次数最多的距离为栏杆距离
                    for (int i = 0; i < rec_len; i++)
                    {
                        if (num_record[i] >= maxNum)
                        {
                            maxNum = num_record[i];
                            real_distance = distance_record[i];
                        }
                    }
                    if(center_distance_flag == false)
                    {
                        //查找栏杆宽度出现次数最多的距离
                        for (int i = 0; i < rec_len_bbox; i++)
                        {
                            if (size_num_record[i] >= sizeMaxNum)
                            {
                                sizeMaxNum = size_num_record[i];
                                bbox_size = bbox_record[i];
                            }
                        }
                    }
                }
            }
            if(center_distance_flag == false)
                return (real_distance - bbox_size) * 0.3048;
            else
                return real_distance * 0.3048;
        }

        private void Cal_Rail_High(Element Stair, FilteredElementCollector Levels_collector, Document doc, ref double beamHigh, ref double LevelHigh, ref double railHigh, ref int floor, ref bool feasible)
        {
            BoundingBoxXYZ stairBBox = Stair.get_BoundingBox(doc.ActiveView);
            if (stairBBox == null)
            {
                feasible = false;
                return;
            }
            Element beam = null;
            XYZ temp = stairBBox.Min;
            XYZ min = new XYZ(temp.X, temp.Y, temp.Z - 3.28 / 2);
            XYZ max = stairBBox.Max;
            Outline outLine = new Outline(min, max);
            BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
            ElementId viewId = doc.ActiveView.Id;
            FilteredElementCollector collector = new FilteredElementCollector(doc, viewId);
            IList<Element> elements = collector.WherePasses(boundingBoxIntersectsFilter).WherePasses(Floor_Filter).ToElements();
            StringBuilder stringBuilder = new StringBuilder();
            List<ElementId> ids = new List<ElementId>();
            foreach (Element item in elements)
            {
                stringBuilder.AppendLine(item.Name);
                ids.Add(item.Id);
                if (beam == null || (beam.get_BoundingBox(doc.ActiveView).Max.Z < item.get_BoundingBox(doc.ActiveView).Max.Z))
                {
                    beam = item;
                }
            }
            ElementId elemL = Stair.LevelId;
            Level L = null;
            foreach (var item in Levels_collector)
            {
                if (item.Id == elemL)
                {
                    L = (Level)item;
                    break;
                }
            }
            if (L == null)
            {
                feasible = false;
                return;
            }
            string levelName = L.Name;
            for (int i = 1; i < levelName.Length; i++)
            {
                int tempFloor = 0;
                if (levelName[i].CompareTo('F') == 0 || levelName[i].CompareTo('f') == 0)
                {
                    if (levelName[i - 1] <= '9' && levelName[i - 1] >= '0')
                        tempFloor = levelName[i - 1] - '0';
                    else
                        continue;
                    if (i > 1 && levelName[i - 2] <= '9' && levelName[i - 2] >= '0')
                        tempFloor += 10 * (levelName[i - 2] - '0');
                    floor = tempFloor;
                }
            }
            LevelHigh = L.Elevation;
            railHigh = Stair.get_BoundingBox(doc.ActiveView).Max.Z;
            if (beam == null)
                beamHigh = -1;
            else
                beamHigh = beam.get_BoundingBox(doc.ActiveView).Max.Z;
        }


        private double Cal_StairsRailing_Glass(Element Stair, FilteredElementCollector Levels_collector, Document doc, ref double beamHigh, ref double LevelHigh, ref double railHigh, ref int floor, ref bool feasible) //计算栏杆下间距
        {
            Solid GlassSD = null;
            GeometryElement geometry = Stair.get_Geometry(new Options());
            foreach (GeometryObject obj in geometry)
            {
                GeometryInstance geometryInstance = obj as GeometryInstance;
                if (geometryInstance != null)
                {
                    GeometryElement instanceEle = geometryInstance.GetInstanceGeometry();
                    foreach (GeometryObject solidObject in instanceEle)
                    {
                        Solid solid = solidObject as Solid;
                        if (solid != null)
                        {
                            FaceArray faceArray = solid.Faces;
                            foreach (Face face in faceArray)
                            {
                                ElementId materialId = face.MaterialElementId;
                                Material material = doc.GetElement(materialId) as Material;
                                if (material != null)
                                {
                                    GlassSD = solid;
                                    break;
                                }
                            }
                        }

                    }
                }
            }
            BoundingBoxXYZ stairBBox = Stair.get_BoundingBox(doc.ActiveView);
            if (stairBBox == null)
            {
                feasible = false;
                return -1;
            }
            Element beam = null;
            XYZ temp = stairBBox.Min;
            XYZ min = new XYZ(temp.X, temp.Y, temp.Z - 3.28 / 2);
            XYZ max = stairBBox.Max;
            Outline outLine = new Outline(min, max);
            BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
            ElementId viewId = doc.ActiveView.Id;
            FilteredElementCollector collector = new FilteredElementCollector(doc, viewId);
            IList<Element> elements = collector.WherePasses(boundingBoxIntersectsFilter).WherePasses(Floor_Filter).ToElements();
            StringBuilder stringBuilder = new StringBuilder();
            List<ElementId> ids = new List<ElementId>();
            foreach (Element item in elements)
            {
                stringBuilder.AppendLine(item.Name);
                ids.Add(item.Id);
                if (beam == null || (beam.get_BoundingBox(doc.ActiveView).Max.Z < item.get_BoundingBox(doc.ActiveView).Max.Z))
                {
                    beam = item;
                }
            }
            double glass2floor;
            ElementId elemL = Stair.LevelId;
            Level L = null;
            foreach (var item in Levels_collector)
            {
                if (item.Id == elemL)
                {
                    L = (Level)item;
                    break;
                }
            }
            if (L == null)
            {
                feasible = false;
                return -1;
            }
            LevelHigh = L.Elevation;
            string levelName = L.Name;
            for (int i = 1; i < levelName.Length; i++)
            {
                int tempFloor = 0;
                if (levelName[i].CompareTo('F') == 0 || levelName[i].CompareTo('f') == 0)
                {
                    if (levelName[i - 1] <= '9' && levelName[i - 1] >= '0')
                        tempFloor = levelName[i - 1] - '0';
                    else
                        continue;
                    if (i > 1 && levelName[i - 2] <= '9' && levelName[i - 2] >= '0')
                        tempFloor += 10 * (levelName[i - 2] - '0');
                    floor = tempFloor;
                }
            }
            railHigh = Stair.get_BoundingBox(doc.ActiveView).Max.Z;
            if (beam == null || beam.get_BoundingBox(doc.ActiveView).Max.Z < LevelHigh)
                glass2floor = (Stair.get_BoundingBox(doc.ActiveView).Max.Z - (GlassSD.GetBoundingBox().Max.Z - GlassSD.GetBoundingBox().Min.Z) - L.Elevation) * 304.8;
            else
            {
                /*
                TaskDialog.Show("Tips",
                    "GlassSD.Z : " + (Stair.get_BoundingBox(doc.ActiveView).Max.Z - (GlassSD.GetBoundingBox().Max.Z - GlassSD.GetBoundingBox().Min.Z)).ToString() + "\n" +
                    "beam_MaxZ : " + beam.get_BoundingBox(doc.ActiveView).Max.Z.ToString() + "\n");
                */
                glass2floor = (Stair.get_BoundingBox(doc.ActiveView).Max.Z - (GlassSD.GetBoundingBox().Max.Z - GlassSD.GetBoundingBox().Min.Z) - beam.get_BoundingBox(doc.ActiveView).Max.Z) * 304.8;
                beamHigh = beam.get_BoundingBox(doc.ActiveView).Max.Z;
            }
            return glass2floor;
        }

        private double[] get_Floor_Line_BBox(List<Element> grid_list,Document doc)
        {

            double[] xyxy = new double[4] { 0, 0, 0, 0 }; //记录最小和最大的(x,y)坐标
            bool[] emptyFlag = new bool[2] { true, true}; //判断xyxy是否为空
            
            bool matchFlag = false; //匹配flag
            foreach (Grid item in grid_list)
            {
                /*------------------------计算极限位置------------------------*/
                BoundingBoxXYZ BBox = null;
                BBox = item.get_BoundingBox(doc.ActiveView);

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
                /*------------------------计算极限位置------------------------*/
            }
            /*------------------------取轴网包络矩形的1.2宽高------------------------*/
            double deta_X = (xyxy[2] - xyxy[0]) * 0.1;
            double deta_Y = (xyxy[3] - xyxy[1]) * 0.1;
            xyxy[0] -= deta_X;
            xyxy[2] += deta_X;
            xyxy[1] -= deta_Y;
            xyxy[3] += deta_Y;
            /*------------------------取轴网包络矩形的1.2宽高------------------------*/
            // TaskDialog.Show("Tips", (xyxy[0]).ToString() + "\n" + (xyxy[1]).ToString() + "\n" + (xyxy[2]).ToString() + "\n" + (xyxy[3]).ToString() + "\n");
            return xyxy; //返回数组{x_min, y_min, x_max, y_max}
        }
        private bool StairsRailing_Judge(FilteredElementCollector StairsRailing_collector, FilteredElementCollector Levels_collector)
        {
            string s = "";
            foreach (var item in Levels_collector)
            {
                s += item.Name + '\n';
            }
            
            TaskDialog.Show("Level", s);
            return true;
        }
        private void Windows_Judge(Element elem, FilteredElementCollector Levels_collector, Document doc, UIDocument uiDoc) //判断窗户是否符合规范
        {
            //输入的elem为windows收集器的item
            ElementId elemL = elem.LevelId;
            Level L = null;
         
            foreach (var item in Levels_collector)
            {
                if (item.Id == elemL)
                {
                    L = (Level)item;
                    break;
                }
            }
            string levelName = L.Name;
            int floor = 1;
            for (int i = 1; i < levelName.Length; i++)
            {
                int tempFloor = 0;
                if (levelName[i].CompareTo('F') == 0 || levelName[i].CompareTo('f') == 0)
                {
                    if (levelName[i - 1] <= '9' && levelName[i - 1] >= '0')
                        tempFloor = levelName[i - 1] - '0';
                    else
                        continue;
                    if (i > 1 && levelName[i - 2] <= '9' && levelName[i - 2] >= '0')
                        tempFloor += 10 * (levelName[i - 2] - '0');
                    floor = tempFloor;
                }
            }
            //TaskDialog.Show("floor", floor.ToString());
            double floorHigh = L.Elevation;
            if (floor <= 1)
            {
                TaskDialog.Show("Judge", "this is A 1F Windows!!!");
            }
            else
            {
                BoundingBoxXYZ temp_BBox = elem.get_BoundingBox(doc.ActiveView);
                XYZ origin1 = null;
                XYZ origin2 = null;
                if (Math.Abs(temp_BBox.Min.X - temp_BBox.Max.X) > Math.Abs(temp_BBox.Min.Y - temp_BBox.Max.Y))
                {
                    origin1 = new XYZ(temp_BBox.Min.X, (temp_BBox.Min.Y + temp_BBox.Max.Y) / 2, (temp_BBox.Min.Z + temp_BBox.Max.Z) / 2);
                    origin2 = new XYZ(temp_BBox.Max.X, (temp_BBox.Min.Y + temp_BBox.Max.Y) / 2, (temp_BBox.Min.Z + temp_BBox.Max.Z) / 2);
                }
                else
                {
                    origin1 = new XYZ((temp_BBox.Min.X + temp_BBox.Max.X) / 2, temp_BBox.Min.Y, (temp_BBox.Min.Z + temp_BBox.Max.Z) / 2);
                    origin2 = new XYZ((temp_BBox.Min.X + temp_BBox.Max.X) / 2, temp_BBox.Max.Y, (temp_BBox.Min.Z + temp_BBox.Max.Z) / 2);
                }
                bool ff1 = false;
                bool ff2 = false;
                //int p = 1;
                for (int p = -1; p < 3; p+=2)
                {
                    Transaction tran = new Transaction(doc, " 拉伸");
                    tran.Start();
                    //创建一个长方体拉伸模型
                    double[] xy = new double[2] { 0, 0 };
                    if ((Math.Abs(origin1.X - origin2.X)) > (Math.Abs(origin1.Y - origin2.Y)))
                        xy[1] = p;
                    else
                        xy[0] = p;

                    // TaskDialog.Show("origin", origin1.ToString() + '\n' + origin2.ToString());
                    // 正侧
                    double deta = 3.2808 * 2;
                    XYZ p1 = new XYZ(origin1.X + xy[0] * deta, origin1.Y + xy[1] * deta, origin1.Z);
                    XYZ p2 = new XYZ(origin2.X + xy[0] * deta, origin2.Y + xy[1] * deta, origin1.Z);
                    XYZ p3 = new XYZ(origin2.X + xy[0] * deta, origin2.Y + xy[1] * deta, L.Elevation - 0.1 * (origin2.Z - L.Elevation));
                    XYZ p4 = new XYZ(origin1.X + xy[0] * deta, origin1.Y + xy[1] * deta, L.Elevation - 0.1 * (origin1.Z - L.Elevation));
                    Line line1 = Line.CreateBound(p1, p2);
                    Line line2 = Line.CreateBound(p2, p3);
                    Line line3 = Line.CreateBound(p3, p4);
                    Line line4 = Line.CreateBound(p4, p1);
                    CurveLoop loop = new CurveLoop();
                    loop.Append(line1);
                    loop.Append(line2);
                    loop.Append(line3);
                    loop.Append(line4);
                    List<CurveLoop> loops = new List<CurveLoop>() { loop };
                    XYZ direction = null;
                    if ((Math.Abs(origin1.X - origin2.X)) > (Math.Abs(origin1.Y - origin2.Y)))
                        direction = new XYZ(0, p, 0);//拉伸方向
                    else
                        direction = new XYZ(p, 0, 0);//拉伸方向
                    Solid solid = GeometryCreationUtilities.CreateExtrusionGeometry(loops, direction, 6.561679);
                    DirectShape shape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_StructuralFoundation));
                    shape.AppendShape(new List<GeometryObject>() { solid });
                    tran.Commit();

                    BoundingBoxXYZ BBox = shape.get_BoundingBox(doc.ActiveView);
                    XYZ min = BBox.Min;
                    XYZ max = BBox.Max;
                    Outline outLine = new Outline(min, max);
                    BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
                    ElementId viewId = doc.ActiveView.Id;
                    FilteredElementCollector collector = new FilteredElementCollector(doc, viewId);
                    IList<Element> elements = collector.WherePasses(boundingBoxIntersectsFilter).WherePasses(Floor_Filter).ToElements();
                    StringBuilder stringBuilder = new StringBuilder();
                    List<ElementId> ids = new List<ElementId>();
                    int count = 0;
                    foreach (Element item in elements)
                    {
                        count++;
                        stringBuilder.AppendLine(item.Name);
                        ids.Add(item.Id);
                    }
                            
                    //将交互的元素高亮
                    //uiDoc.Selection.SetElementIds(ids);

                    if (count > 0)
                    {
                        if (p == -1)
                            ff1 = true;
                        else
                            ff2 = true;
                    }

                    // TaskDialog.Show("提示", stringBuilder.ToString());
                    tran.Start();
                    doc.Delete(shape.Id);
                    tran.Commit();
                }
                if((ff1 && !ff2) || (!ff1 && ff2))
                {
                    Transaction tran = new Transaction(doc, " 拉伸");
                    tran.Start();
                    //创建一个长方体拉伸模型
                    double[] xy = new double[2] { 0, 0 };
                    if ((Math.Abs(origin1.X - origin2.X)) > (Math.Abs(origin1.Y - origin2.Y)))
                        xy[1] = 1;
                    else
                        xy[0] = 1;

                    // 正侧
                    double deta = 3.2808 * 2;
                    XYZ p1 = new XYZ(origin1.X + xy[0] * deta, origin1.Y + xy[1] * deta, origin1.Z);
                    XYZ p2 = new XYZ(origin2.X + xy[0] * deta, origin2.Y + xy[1] * deta, origin1.Z);
                    XYZ p3 = new XYZ(origin2.X + xy[0] * deta, origin2.Y + xy[1] * deta, L.Elevation - 0.1 * (origin2.Z - L.Elevation));
                    XYZ p4 = new XYZ(origin1.X + xy[0] * deta, origin1.Y + xy[1] * deta, L.Elevation - 0.1 * (origin1.Z - L.Elevation));
                    Line line1 = Line.CreateBound(p1, p2);
                    Line line2 = Line.CreateBound(p2, p3);
                    Line line3 = Line.CreateBound(p3, p4);
                    Line line4 = Line.CreateBound(p4, p1);
                    CurveLoop loop = new CurveLoop();
                    loop.Append(line1);
                    loop.Append(line2);
                    loop.Append(line3);
                    loop.Append(line4);
                    List<CurveLoop> loops = new List<CurveLoop>() { loop };
                    XYZ direction = null;
                    if ((Math.Abs(origin1.X - origin2.X)) > (Math.Abs(origin1.Y - origin2.Y)))
                        direction = new XYZ(0, -1, 0);//拉伸方向
                    else
                        direction = new XYZ(-1, 0, 0);//拉伸方向
                    Solid solid = GeometryCreationUtilities.CreateExtrusionGeometry(loops, direction, 3.2808 * 4);
                    DirectShape shape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_StructuralFoundation));
                    shape.AppendShape(new List<GeometryObject>() { solid });
                    tran.Commit();

                    BoundingBoxXYZ BBox = shape.get_BoundingBox(doc.ActiveView);
                    XYZ min = BBox.Min;
                    XYZ max = BBox.Max;
                    Outline outLine = new Outline(min, max);
                    BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
                    ElementId viewId = doc.ActiveView.Id;
                    FilteredElementCollector collector = new FilteredElementCollector(doc, viewId);
                    ElementCategoryFilter StairsRailing_Filter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRailing); // 栏杆过滤器
                    IList<Element> elements = collector.WherePasses(boundingBoxIntersectsFilter).WherePasses(StairsRailing_Filter).ToElements();
                    StringBuilder stringBuilder = new StringBuilder();
                    List<ElementId> ids = new List<ElementId>();
                    int count = 0;
                    foreach (Element item in elements)
                    {
                        count++;
                        stringBuilder.AppendLine(item.Name);
                        ids.Add(item.Id);
                    }
                    //将交互的元素高亮
                    //uiDoc.Selection.SetElementIds(ids);
                    tran.Start();
                    doc.Delete(shape.Id);
                    tran.Commit();
                    if (count > 0)
                        TaskDialog.Show("提示", "该窗户与外界相邻且有栏杆");
                    else
                        TaskDialog.Show("提示", "该窗户与外界相邻且无栏杆");
                }
                else
                {
                    TaskDialog.Show("提示", "该窗户与外界不相邻");
                }
            }
        }
        private bool Windows_Judge_Simple(Element elem, FilteredElementCollector Levels_collector, Document doc, UIDocument uiDoc, ref bool feasible, ref int floor) //判断窗户是否符合规范
        {
            //输入的elem为windows收集器的item
            ElementId elemL = elem.LevelId;
            Level L = null;

            foreach (var item in Levels_collector)
            {
                if (item.Id == elemL)
                {
                    L = (Level)item;
                    break;
                }
            }
            if (L == null)
            {
                feasible = false;
                return false;
            }
            string levelName = L.Name;
            bool floorflag = false;
            floor = 1;
            for (int i = 1; i < levelName.Length; i++)
            {
                int tempFloor = 0;
                if (levelName[i].CompareTo('F') == 0 || levelName[i].CompareTo('f') == 0)
                {
                    if (levelName[i - 1] <= '9' && levelName[i - 1] >= '0')
                        tempFloor = levelName[i - 1] - '0';
                    else
                        continue;
                    if (i > 1 && levelName[i - 2] <= '9' && levelName[i - 2] >= '0')
                        tempFloor += 10 * (levelName[i - 2] - '0');
                    floor = tempFloor;
                    floorflag = true;
                }
            }
            double floorHigh = L.Elevation;
            if (floor == 1 && floorflag == true)
            {
                return true;
            }
            else
            {
                BoundingBoxXYZ temp_BBox = elem.get_BoundingBox(doc.ActiveView);
                XYZ origin1 = null;
                XYZ origin2 = null;
                if (Math.Abs(temp_BBox.Min.X - temp_BBox.Max.X) > Math.Abs(temp_BBox.Min.Y - temp_BBox.Max.Y))
                {
                    origin1 = new XYZ(temp_BBox.Min.X, (temp_BBox.Min.Y + temp_BBox.Max.Y) / 2, (temp_BBox.Min.Z + temp_BBox.Max.Z) / 2);
                    origin2 = new XYZ(temp_BBox.Max.X, (temp_BBox.Min.Y + temp_BBox.Max.Y) / 2, (temp_BBox.Min.Z + temp_BBox.Max.Z) / 2);
                }
                else
                {
                    origin1 = new XYZ((temp_BBox.Min.X + temp_BBox.Max.X) / 2, temp_BBox.Min.Y, (temp_BBox.Min.Z + temp_BBox.Max.Z) / 2);
                    origin2 = new XYZ((temp_BBox.Min.X + temp_BBox.Max.X) / 2, temp_BBox.Max.Y, (temp_BBox.Min.Z + temp_BBox.Max.Z) / 2);
                }
                Transaction tran = new Transaction(doc, " 拉伸");
                tran.Start();
                //创建一个长方体拉伸模型
                double[] xy = new double[2] { 0, 0 };
                if ((Math.Abs(origin1.X - origin2.X)) > (Math.Abs(origin1.Y - origin2.Y)))
                    xy[1] = 1;
                else
                    xy[0] = 1;

                double deta = 3.2808 * 1;
                XYZ p1 = new XYZ(origin1.X + xy[0] * deta, origin1.Y + xy[1] * deta, origin1.Z);
                XYZ p2 = new XYZ(origin2.X + xy[0] * deta, origin2.Y + xy[1] * deta, origin1.Z);
                XYZ p3 = new XYZ(origin2.X + xy[0] * deta, origin2.Y + xy[1] * deta, L.Elevation - 0.1 * (origin2.Z - L.Elevation));
                XYZ p4 = new XYZ(origin1.X + xy[0] * deta, origin1.Y + xy[1] * deta, L.Elevation - 0.1 * (origin1.Z - L.Elevation));
                Line line1 = Line.CreateBound(p1, p2);
                Line line2 = Line.CreateBound(p2, p3);
                Line line3 = Line.CreateBound(p3, p4);
                Line line4 = Line.CreateBound(p4, p1);
                CurveLoop loop = new CurveLoop();
                loop.Append(line1);
                loop.Append(line2);
                loop.Append(line3);
                loop.Append(line4);
                List<CurveLoop> loops = new List<CurveLoop>() { loop };
                XYZ direction = null;
                if ((Math.Abs(origin1.X - origin2.X)) > (Math.Abs(origin1.Y - origin2.Y)))
                    direction = new XYZ(0, -1, 0);//拉伸方向
                else
                    direction = new XYZ(-1, 0, 0);//拉伸方向
                Solid solid = GeometryCreationUtilities.CreateExtrusionGeometry(loops, direction, 3.2808 * 4);
                DirectShape shape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_StructuralFoundation));
                shape.AppendShape(new List<GeometryObject>() { solid });
                tran.Commit();


                BoundingBoxXYZ BBox = shape.get_BoundingBox(doc.ActiveView);
                XYZ min = BBox.Min;
                XYZ max = BBox.Max;
                Outline outLine = new Outline(min, max);
                BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
                ElementId viewId = doc.ActiveView.Id;
                FilteredElementCollector collector = new FilteredElementCollector(doc, viewId);
                ElementCategoryFilter StairsRailing_Filter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRailing); // 栏杆过滤器
                IList<Element> elements = collector.WherePasses(boundingBoxIntersectsFilter).WherePasses(StairsRailing_Filter).ToElements();
                StringBuilder stringBuilder = new StringBuilder();
                List<ElementId> ids = new List<ElementId>();
                int count = 0;
                foreach (Element item in elements)
                {
                    count++;
                    stringBuilder.AppendLine(item.Name);
                    ids.Add(item.Id);
                }

                //将交互的元素高亮
                //uiDoc.Selection.SetElementIds(ids);
                //TaskDialog.Show("Judge", count.ToString());
                tran.Start();
                doc.Delete(shape.Id);
                tran.Commit();
                if (count > 0)
                {
                    for (int i = 0; i < elements.Count(); i++)
                        windows_stair.Add(elements[i].Id.ToString());
                    bool disFlag = true;
                    bool highFlag = true;
                    GeometryElement gElem = elements[0].get_Geometry(new Options());
                    double glassDis = 0;
                    double distance = 0;
                    double beamHigh = 0;
                    double levelHigh = 0;
                    double railHigh = 0;
                    Cal_Rail_High(elements[0], Levels_collector, doc, ref beamHigh, ref levelHigh, ref railHigh, ref floor, ref feasible);
                    if (feasible == false)
                        return false;
                    distance = Math.Round(1000 * Cal_StairsRailing_Distance(gElem), mmDec);
                    if (distance <= 110)
                        disFlag = true;
                    else
                        disFlag = false;

                    if (Math.Round((railHigh - levelHigh) * 304.8, mmDec) >= 900)
                        highFlag = true;
                    else
                        highFlag = false;

                    if (disFlag && highFlag)
                    {
                        // TaskDialog.Show("提示", "flag true");
                        return true;
                    }
                    else
                    {
                        // TaskDialog.Show("提示", "flag false");
                        return false;
                    }
                        
                }
                else
                {
                    // TaskDialog.Show("提示", "else false");
                    return false;
                }

            }
        }

        private bool judge_material(Element Stair, Document doc)
        {

            GeometryElement geometry = Stair.get_Geometry(new Options());
            bool flag = false;
            try
            {
                foreach (GeometryObject obj in geometry)
                {
                    GeometryInstance geometryInstance = obj as GeometryInstance;
                    if (geometryInstance != null)
                    {
                        GeometryElement instanceEle = geometryInstance.GetInstanceGeometry();
                        foreach (GeometryObject solidObject in instanceEle)
                        {
                            Solid solid = solidObject as Solid;
                            if (solid != null)
                            {
                                FaceArray faceArray = solid.Faces;
                                foreach (Face face in faceArray)
                                {
                                    ElementId materialId = face.MaterialElementId;
                                    Material material = doc.GetElement(materialId) as Material;
                                    if (material != null)
                                    {
                                        flag = true;
                                        break;
                                    }
                                }
                            }

                        }
                        if (flag)
                        {
                            break;
                        }
                    }
                    if (flag)
                    {
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                flag = false;
            }
            
            return flag;
        }

        private bool BelongToStairandExist(Element elem, Document doc)
        {
            if (elem.get_BoundingBox(doc.ActiveView) != null)
            {
                Railing railing = elem as Railing;
                if (railing.HasHost == true)
                    return true;
                else if (elem.LevelId == null)
                    return true;
                else
                {
                    Level L = null;
                    L = doc.GetElement(elem.LevelId) as Level;
                    if (L == null)
                        return true;
                    else
                        if ((elem.get_BoundingBox(doc.ActiveView).Max.Z + elem.get_BoundingBox(doc.ActiveView).Min.Z) / 2 > L.Elevation)
                            return false;
                        else
                            return true;
                }
            }
            else
                return true;
            
        }
        
        public static void PrintLog(string info, string itemNum, Document doc, string path = null)
        {
            // 获取模型编号M0000x
            Regex regex = new Regex(@"(?<=\\)M[0-9]{5}");
            string modelNum = "M0000x";
            string docPath = doc.PathName;
            if (regex.IsMatch(docPath))
                modelNum = regex.Match(docPath).ToString();

            // 默认路径为当前用户的桌面
            if (path == null)
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string date = DateTime.Now.ToString("yyyyMMdd");
                path = string.Format("{0}\\{1}-{2}-{3}.txt", desktop, itemNum, modelNum, date);
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
        public void insertMessage(List<building_message> text, Element elem, int append_Idx, Document doc)
        {
            BoundingBoxXYZ bbox= elem.get_BoundingBox(doc.ActiveView);
            double centerX = (bbox.Max.X + bbox.Min.X) / 2;
            double centerY = (bbox.Max.Y + bbox.Min.Y) / 2;

            //最后一个是未知楼栋
            for (int i = 0; i < text.Count() - 1; i++)
            {
                if (text[i].BBox[3] >= centerY && centerY >= text[i].BBox[1] && text[i].BBox[2] >= centerX && centerX >= text[i].BBox[0])
                {
                    text[i].whole[append_Idx].Add(elem.Id.ToString());
                    return;
                }
                
            }
            text[text.Count() - 1].whole[append_Idx].Add(elem.Id.ToString());
        }

        string getMessage(List<building_message> text, bool flag)
        {
            text.Sort((left, right) =>
            {
                int NumL, NumR;
                bool resultL = int.TryParse(left.Building, out NumL);
                bool resultR = int.TryParse(right.Building, out NumR);
                if (resultL && resultR)
                {
                    if (NumL > NumR) return 1;
                    else return -1;
                }
                else if (resultL && !resultR)
                    return 1;
                else if (!resultL && resultR)
                    return -1;
                else
                    return string.Compare(left.Building, right.Building);
            });
            string s = "";
            if (flag)
            {
                s += "符合5.1.5住宅建筑规范GB50368-2005\n符合5.6.2  5.6.3  5.8.1  6.1.1住宅设计规范GB 50096-2011\n";
                return s;
            }
            else
                s = "不符合5.1.5住宅建筑规范GB50368-2005\n不符合5.6.2  5.6.3  5.8.1  6.1.1住宅设计规范GB 50096-2011\n";
            for (int i = 0; i < text.Count(); i++)
            {
                s += "楼栋编号:" + text[i].Building.ToString() + "#\n";

                s += "不符合要求的窗台:";
                for (int j = 0; j < text[i].windows_NG.Count(); j++)
                {
                    s += text[i].windows_NG[j].ToString() + " ";
                }
                s += "\n";

                s += "符合要求的窗台:";
                for (int j = 0; j < text[i].windows_OK.Count(); j++)
                {
                    s += text[i].windows_OK[j].ToString() + " ";
                }
                s += "\n";

                s += "高度不符合要求的栏杆:";
                for (int j = 0; j < text[i].rail_high_NG.Count(); j++)
                {
                    s += text[i].rail_high_NG[j].ToString() + " ";
                }
                s += "\n";

                s += "垂直杆件间净距不符合要求的栏杆:";
                for (int j = 0; j < text[i].rail_distance_NG.Count(); j++)
                {
                    s += text[i].rail_distance_NG[j].ToString() + " ";
                }
                s += "\n";

                s += "符合要求的栏杆:";
                for (int j = 0; j < text[i].rail_OK.Count(); j++)
                {
                    s += text[i].rail_OK[j].ToString() + " ";
                }
                s += "\n\n";

            }
            return s;
        }
    }
}



