using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Architecture;


namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    public class H00012:IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //界面交互的doc
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            //实际内容的doc
            Document doc = commandData.Application.ActiveUIDocument.Document;
            //创建收集器
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            //通用过滤方法--获取楼梯
            ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_Stairs);
            collector.WherePasses(elementCategoryFilter);

            //判断地址
            bool locationFalg=false;
            int locCount = 0;
            

            //楼梯踏步高度规范
            StringBuilder heightRule = new StringBuilder();
            
            //楼梯踏步宽度规范
            StringBuilder widthRule = new StringBuilder();
           
            //楼梯梯段宽度规范
            StringBuilder railing_wRule = new StringBuilder();
            

            //楼梯井
            StringBuilder stairWell = new StringBuilder();
            
            //先获取所有楼号信息
            List<Object> buildingMsg = getBuildingLevelCount(uiDoc, doc, commandData);

            Dictionary<string, int> stairsIdMap = new Dictionary<string, int>();

            foreach (var item in collector)
            {
                Element stair = item as Element;
                cal_stairShaft(doc, stair, ref stairWell);

                if (locCount == 0 && stair.Document.SiteLocation.PlaceName.Contains("浙江"))
                {
                    locationFalg=true;
                    locCount++;
                }
               
                //获取属性
                ParameterSet parameterSets = stair.Parameters;
                foreach(Parameter parameter in parameterSets)
                {
                    //判断踏步高度
                    if (parameter.Definition.Name == "实际踢面高度")
                    {
                        double height = Convert.ToDouble(parameter.AsValueString());
                        
                        if (height > 175)
                        {
                            heightRule.Append(stair.Id + " "); ;                            
                        }
                        
                    }

                    //实际踏板深度(踏步宽度)
                    if (parameter.Definition.Name == "实际踏板深度")
                    {
                        double width = Convert.ToDouble(parameter.AsValueString());

                        if (width < 260)
                        {
                            widthRule.Append(stair.Id + " ");
                            
                        }

                    }

                }

                //获取踏井宽度
                Stairs stair_true = item as Stairs;

                //楼梯的栏杆
                ICollection<ElementId> roliings = getStairsRailings(doc, stair);
                if (stair_true != null)
                {
                    //楼梯的梯段
                    ICollection<ElementId> ls = stair_true.GetStairsRuns();
                    if (ls.Count() == 2)
                    {
                        List<double> stairs_w = new List<double>();

                        List<double> line_index = new List<double>();
                        foreach (ElementId l in ls)
                        {
                            Element ele = doc.GetElement(l);
                            StairsRun stairsRun = ele as StairsRun;
                            //判断梯段方向
                            CurveLoop lines = stairsRun.GetStairsPath();
                            int x_y = -1;
                            if (lines.Count() == 1)
                            {
                                foreach (Line line in lines)
                                {
                                    line_index.Add(line.Direction.X);
                                    line_index.Add(line.Direction.Y);
                                    if (Math.Abs(line.Direction.X) > 0.9)
                                    {
                                        x_y = 1;
                                    }
                                    else
                                    {
                                        x_y = 0;
                                    }
                                }
                            }

                            BoundingBoxXYZ xyz = ele.get_BoundingBox(doc.ActiveView);
                            
                            if (x_y == 0)
                            {
                                stairs_w.Add(xyz.Max.X);
                                stairs_w.Add(xyz.Min.X);
                            }
                            else
                            {
                                stairs_w.Add(xyz.Max.Y);
                                stairs_w.Add(xyz.Min.Y);
                            }
                        }
                        //判断平行
                        double line_arc = line_index[0] * line_index[2] + line_index[1] * line_index[3];
                        if (Math.Abs(line_arc) > 0.01)
                        {
                            stairs_w.Sort();
                            double ww = (stairs_w[2] - stairs_w[1]) * 0.3048;
                            //判断踏井宽度是否符合规范
                            if (ww > 0.11)
                            {
                                stairsIdMap.Add(stair.Id.ToString(), 1);
                                stairWell.Append(stair.Id.ToString() + " ");
                            }
                        }
                    }
                    if (stair_true.GetStairsRuns() != null&& roliings!=null)
                    {
                        //获取楼梯间距
                        Dictionary<string, object> result = getStairsWidth(stair_true.GetStairsRuns(), roliings, doc);

                        //获取楼栋层数
                        string[] levelMap = getElementBuildingNum(stair, uiDoc, doc, commandData, buildingMsg);
                        if (levelMap == null)
                        {
                            continue;
                        }
                        int levels = int.Parse(levelMap[1]);
                        List<int> stairCount = (List<int>)result["stairCount"];
                        List<double> weight = (List<double>)result["weight"];
                        for (int i = 0; i < stairCount.Count; i++)
                        {
                            double stair_width = weight[i];
                            int flag = stairCount[i];
                            if (levels >= 7 || flag == 2)
                            {
                                if (stair_width < 1.1)
                                {
                                    //railing_wRule.AppendLine("七层及七层以上或者两边有栏杆，楼梯梯段" + stair.Id + "净宽为" + stair_width.ToString("f4") + "不应小于1.10m。");
                                    railing_wRule.Append(stair.Id + " ");
                                }

                            }
                            else if (levels < 7)
                            {
                                if (stair_width < 1)
                                {
                                    //railing_wRule.AppendLine("六层及六层以下，楼梯梯段" + stair.Id + "净宽为" + stair_width.ToString("f4") + "不应小于1.0m。");
                                    railing_wRule.Append(stair.Id + " ");
                                }
                            }
                        }
                    }
                }

            }
            //TaskDialog.Show("楼梯踏步高度规范", heightRule.ToString());
            //TaskDialog.Show("楼梯踏步宽度规范", widthRule.ToString());

            
            //3.	检测所有扶手高度，不应小于0.9m
            //创建收集器
            FilteredElementCollector collector_railing = new FilteredElementCollector(doc);
            //通用过滤方法--获取扶手
            //ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRailing);
            //collector.WherePasses(elementCategoryFilter);
            elementCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRailing);
            collector_railing.WherePasses(elementCategoryFilter);

            //楼梯扶手高度规范
            StringBuilder railing_hRule = new StringBuilder();
            
            //楼梯扶手间距规范
            StringBuilder railing_disRule = new StringBuilder();
            

            //防攀爬措施
            StringBuilder methods = new StringBuilder();
            

            foreach (var item in collector_railing)
            {
                if(getRailMaterial(item, doc))
                {
                    continue;
                }

                Railing railing = item as Railing;

                //获取楼梯栏杆
                if (railing != null&& railing.HasHost )
                {
                    //获取主从对象
                    Element hostEle = doc.GetElement(railing.HostId);
                    

                    if (hostEle.Name.Contains("楼梯"))
                    {
                        //得到栏杆
                        Element railingElement = item as Element;
                        
                        
                        //获取栏杆几何对象
                        GeometryElement railingGeometry = railingElement.get_Geometry(new Options());
                        //获取栏杆间距
                        double railing_distance = Cal_StairsRailing_Distance(railingGeometry);
                        if (railing_distance > 0.110)
                        {
                            railing_disRule.Append(railingElement.Id + " ");
                        }

                        Element newEle = doc.GetElement(railingElement.GetTypeId());

                        BoundingBoxXYZ xyz = railingElement.get_BoundingBox(doc.ActiveView);

                        //根据boundingbox获取高度
                        double z = (xyz.Max.Z - xyz.Min.Z) * 0.3048 * 1000;
                        
                        foreach (Parameter parameter in newEle.Parameters)
                        {
                            //判断扶手高度
                            if (parameter.Definition.Name == "栏杆扶手高度")
                            {
                                double height = Convert.ToDouble(parameter.AsValueString());

                                if (height < 900 )
                                {
                                    
                                    railing_hRule.Append(railing.Id + " ");
                                    //HighLint(uiDoc, railing.HostId);
                                }
                                //判断防攀爬措施
                                if (stairsIdMap.ContainsKey(railing.HostId.ToString()) && height <450)
                                {
                                    methods.Append(railing.Id + " ");
                                    stairsIdMap.Remove(railing.HostId.ToString());
                                }

                                //判断水平梯段
                                if (height == z)
                                {
                                    foreach (Parameter railingParameter in railingElement.Parameters)
                                    {
                                        if (parameter.Definition.Name == "长度")
                                        {
                                            double railingLength = Convert.ToDouble(railingParameter.AsValueString());
                                            if (railingLength > 500 && height < 1050)
                                            {
                                                railing_hRule.Append( railing.Id + " ");
                                            }

                                            
                                        }
                                    }
                                }
                                
                                break;

                            }
                        }
                    }
                }
            }
            
            //TaskDialog.Show("楼梯扶手高度规范", railing_hRule.ToString());
            //TaskDialog.Show("楼梯扶手间距规范", railing_disRule.ToString());
            //TaskDialog.Show("楼梯梯段宽度规范", railing_wRule.ToString());
            //TaskDialog.Show("防攀爬措施", methods.ToString());

            //整合输出
            int resLength= railing_wRule.Length+ widthRule.Length+ heightRule.Length+ railing_hRule.Length+ railing_disRule.Length+ methods.Length+ stairWell.Length;
            StringBuilder resultString = new StringBuilder();
            if (resLength > 0)
            {
                

                resultString.AppendLine("不符合5.2.3住宅建筑规范GB 50368-2005");
                resultString.AppendLine("不符合6.3.2  6.3.5住宅设计规范GB 50096-2011");
                if (locationFalg)
                {
                    resultString.AppendLine("不符合5.2.2 5.2.6浙江省住宅设计标准DB33/ 1006-2017");
                }

                resultString.AppendLine("\n");
                resultString.AppendLine("楼梯梯段净宽");
                resultString.AppendLine("不满足要求的有:");
                resultString.Append(railing_wRule);
                resultString.AppendLine("\n");
                resultString.AppendLine("楼梯踏步宽度");
                resultString.AppendLine("不满足要求的有:");
                resultString.Append(widthRule);
                resultString.AppendLine("\n");
                resultString.AppendLine("楼梯踏步高度");
                resultString.AppendLine("不满足要求的有:");
                resultString.Append(heightRule);
                resultString.AppendLine("\n");
                resultString.AppendLine("楼梯扶手高度");
                resultString.AppendLine("不满足要求的有:");
                resultString.Append(railing_hRule);
                resultString.AppendLine("\n");
                resultString.AppendLine("楼梯栏杆垂直杆件间距");
                resultString.AppendLine("不满足要求的有:");
                resultString.Append(railing_disRule);
                resultString.AppendLine("\n");
                resultString.AppendLine("楼梯水平段栏杆扶手高度");
                resultString.AppendLine("不满足要求的有:");
                resultString.Append(methods);
                resultString.AppendLine("\n");
                resultString.AppendLine("楼梯井净宽大于0.11m的构件有");
                resultString.Append(stairWell);
                resultString.AppendLine("\n");
                
            }
            else
            {
                resultString.AppendLine("符合5.2.3住宅建筑规范GB 50368-2005");
                resultString.AppendLine("符合6.3.2  6.3.5住宅设计规范GB 50096-2011");
                if (locationFalg)
                {
                    resultString.AppendLine("符合5.2.2 5.2.6浙江省住宅设计标准DB33/ 1006-2017");
                }

            }
            TaskDialog.Show("楼梯栏杆规范", resultString.ToString());



            return Result.Succeeded;
        }
        //根据id高亮显示
        private void HighLint(UIDocument uiDoc, IList<ElementId> lElementIds)
        {
            var sel = uiDoc.Selection.GetElementIds();
            foreach (var item in lElementIds)
            {
                sel.Add(item);
            }
            uiDoc.Selection.SetElementIds(sel);
        }
        //根据单个id高亮显示
        private void HighLint(UIDocument uiDoc, ElementId lElementIds)
        {
            var sel = uiDoc.Selection.GetElementIds();
            
            sel.Add(lElementIds);
            
            uiDoc.Selection.SetElementIds(sel);
        }
        //高亮element显示
        private void HighLint(UIDocument uiDoc, IList<Element> lElements)
        {
            var sel = uiDoc.Selection.GetElementIds();
            foreach (var item in lElements)
            {
                sel.Add(item.Id);
            }
            uiDoc.Selection.SetElementIds(sel);

        }

        //获取楼梯栏杆间距
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
                    if (center_distance_flag == false)
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
            if (center_distance_flag == false)
                return (real_distance - bbox_size) * 0.3048;
            else
                return real_distance * 0.3048;
        }

        //计算梯段宽度
        private double getStairsRailingWidth(BoundingBoxXYZ stairXYZ,BoundingBoxXYZ railingXYZ)
        {
 
            double x_center = (railingXYZ.Max.X + railingXYZ.Min.X) / 2;
            double y_center = (railingXYZ.Max.Y + railingXYZ.Min.Y) / 2;

            double x_w = Math.Min(Math.Abs(x_center - stairXYZ.Max.X), Math.Abs(x_center - stairXYZ.Min.X));
            double y_w = Math.Min(Math.Abs(y_center - stairXYZ.Max.Y), Math.Abs(y_center - stairXYZ.Min.Y));
            //判断是否最短
            double res = Math.Min(x_w, y_w) * 0.3048;
            return res;
        }

        //获取某元素所在楼号和楼层
        private string[] getElementBuildingNum(Element el, UIDocument uiDoc, Document doc, ExternalCommandData commandData,List<Object> res)
        {
            BoundingBoxXYZ xyz = el.get_BoundingBox(doc.ActiveView);

            //List<Object> res = getBuildingLevelCount(uiDoc, doc, commandData);
            Dictionary<string, int> num = (Dictionary<string, int>)res[0];
            Dictionary<string, double[]> keyValuePairs = (Dictionary<string, double[]>)res[1];
            string index = "";
            foreach (var item in keyValuePairs)
            {
                double[] value = (double[])item.Value;
                //判断元素是否在楼栋范围内
                if (xyz.Min.X > value[0] && xyz.Min.Y > value[1] && xyz.Max.X < value[2] && xyz.Max.Y < value[3])
                {
                    index = item.Key;
                }
            }
            if (index == "")
            {
                return null;
            }
            string[] elementMsg = new string[2];
            //返回楼号
            elementMsg[0] = index;
            if (!num.ContainsKey(index))
            {
                return null;
            }
            //TaskDialog.Show("防攀爬措施", methods.ToString());
            //返回该楼号层数
            elementMsg[1] = num[index].ToString();
            return elementMsg;
        }

        //获取楼栋层数和楼栋范围
        private List<Object> getBuildingLevelCount(UIDocument uiDoc, Document doc, ExternalCommandData commandData)
        {

            FilteredElementCollector collectorLevels = new FilteredElementCollector(doc);
            List<Element> elementsLevels = collectorLevels.OfCategory(BuiltInCategory.OST_Levels).ToList<Element>();

            //*****************************************获取建筑栋数以及每栋轴网大小
            StringBuilder stringBuilder = new StringBuilder();
            SortedSet<string> build = new SortedSet<string>();//存放建筑楼栋的编号，1，2，3...

            FilteredElementCollector Grid_collector = new FilteredElementCollector(doc); // 轴网收集器       
            List<Element> elementsGrids = Grid_collector.OfCategory(BuiltInCategory.OST_Grids).ToList<Element>();
            List<Element> GridsList = new List<Element>();

            double[] xyxy = new double[4] { 0, 0, 0, 0 }; //记录最小和最大的(x,y)坐标
            bool[] emptyFlag = new bool[2] { true, true }; //判断xyxy是否为空



            foreach (var item in elementsGrids)
            {
                //stringBuilder.AppendLine(item.Name);
                string[] array1 = item.Name.Split(new char[] { '-' });
                if (item.Name.Contains('-'))
                {
                    GridsList.Add(item);
                    build.Add(array1[0]);
                }
            }

            //按轴网名字排序
            GridsList.Sort(SortItem);


            //轴网在二维视图下，根据轴网分好楼栋后，需要转回三维视图
            //***************************************视图转换******************************************
            
            View view = null;

            FilteredElementCollector col_view = new FilteredElementCollector(doc);
            col_view.OfClass(typeof(ViewPlan));//获得所有平面图

            List<ElementId> ids = new List<ElementId>();
            foreach (Element elem in col_view.ToElements())
            {
                ViewPlan viewPlan = elem as ViewPlan;
                if (viewPlan == null) continue;
                ViewFamilyType viewType = doc.GetElement(viewPlan.GetTypeId()) as ViewFamilyType;
                if (viewType == null) continue;
                if (viewType.Name != "楼层平面") continue;
                if (elem.Name.Contains("#1F"))
                {
                    view = viewPlan;
                    break;
                }
            }
            //***************************************视图转换******************************************
            Dictionary<string, double[]> dictionary = new Dictionary<string, double[]>();
            foreach (string bulidName in build)
            {
                double[] xxyy = get_Floor_Line_BBox(bulidName, doc,view);
                dictionary[bulidName] = xxyy;
            }

            ////****************计算楼栋和楼层模块*****************************
            int GroupsNumber = build.Count;
            stringBuilder.AppendLine(GroupsNumber.ToString());
            if (GroupsNumber != 0)
            {
                stringBuilder.AppendLine("住宅建筑数量：" + GroupsNumber);
            }
            else
            {
                stringBuilder.AppendLine("住宅建筑数量：" + 0);
            }
            stringBuilder.AppendLine();


            //int[] num = new int[GroupsNumber + 1];
            Dictionary<string, int> bulidNameToLevels = new Dictionary<string, int>();//存放建筑楼栋的名字和对应层数
            foreach (Element item in elementsLevels)
            {
                if (item.Name.Contains("F") && !item.Name.Contains("S"))
                {
                    MySubNum(item.Name.ToString(), ref bulidNameToLevels);//计算每栋楼层数函数
                    stringBuilder.AppendLine("梁：" + item.Name);
                }

            }


            //uiDoc.ActiveView = root_view;
            //doc = commandData.Application.ActiveUIDocument.Document;

            ////""""""""""""计算楼栋和楼层模块******************************

            //*****************************************获取建筑栋数以及每栋轴网大小
            //TaskDialog.Show("防攀爬措施", stringBuilder.ToString());
            List<Object> list = new List<Object>();
            list.Add(bulidNameToLevels);
            list.Add(dictionary);
            return list;
        }

        //两个item对象名比较，排序
        private static int SortItem(Element item1, Element item2)
        {
            //传入的对象为列表中的对象
            //进行两两比较，用左边的和右边的 按条件 比较
            //返回值规则与接口方法相同
            if (item1.Name != item2.Name)
            {
                return item1.Name.CompareTo(item2.Name);
            }
            else
            {
                return 0;
            }
        }

        //根据轴网分楼栋
        //根据轴网确定建筑楼栋
        private double[] get_Floor_Line_BBox(string bulidName, Document doc, View viewPlan)
        {
            //******************************************************************
            double[] xyxy = new double[4] { 0, 0, 0, 0 }; //记录最小和最大的(x,y)坐标
            bool[] emptyFlag = new bool[2] { true, true }; //判断xyxy是否为空
            StringBuilder stringBuilder = new StringBuilder();
            FilteredElementCollector Grid_collector = new FilteredElementCollector(doc, viewPlan.Id).OfCategory(BuiltInCategory.OST_Grids);


            //楼号, 目前还是用int来表示楼号
            //if (bulidName != "0")
            int flag = 0;
            if (bulidName != "0")
            {
                bool matchFlag = false; //匹配flag

                foreach (var item in Grid_collector)
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


                    if (matchFlag == true)
                    {
                        if (flag == 0)
                        {
                            Grid grid = item as Grid;
                            Outline outline = grid.GetExtents();
                            XYZ maxPoint = outline.MaximumPoint;
                            XYZ minPoint = outline.MinimumPoint;
                            xyxy[0] = minPoint.X;
                            xyxy[1] = minPoint.Y;
                            xyxy[2] = maxPoint.X;
                            xyxy[3] = maxPoint.Y;
                            flag = 1;

                        }
                        else
                        {
                            Grid grid = item as Grid;
                            Outline outline = grid.GetExtents();
                            XYZ maxPoint = outline.MaximumPoint;
                            XYZ minPoint = outline.MinimumPoint;
                            xyxy[0] = (xyxy[0] < minPoint.X) ? xyxy[0] : minPoint.X;
                            xyxy[1] = (xyxy[1] < minPoint.Y) ? xyxy[1] : minPoint.Y;
                            xyxy[2] = (xyxy[2] > maxPoint.X) ? xyxy[2] : maxPoint.X;
                            xyxy[3] = (xyxy[3] > maxPoint.Y) ? xyxy[3] : maxPoint.Y;
                        }

                    }
                    /*------------------------计算极限位置------------------------*/
                }
            }
            /*------------------------取轴网包络矩形的1.2宽高------------------------*/
           
            //TaskDialog.Show("提示", stringBuilder.ToString());
            return xyxy; //返回数组{x_min, y_min, x_max, y_max}
        }


        //根据level字符串算出楼层数
        private void MySubNum(string str, ref Dictionary<string, int> bulidNameToLevels)
        {

            string[] splitStr = str.Split(new char[] { '、', 'F', '#' });

            for (int i = 0; i < splitStr.Length - 2; i++)
            {
                if (splitStr[i] == "")
                {
                    continue;
                }

                if (splitStr[i].Contains('-') || NameStringOrNot(splitStr[i]))
                {
                    if (int.Parse(splitStr[i]) < 0)
                    {
                        int a = int.Parse(splitStr[i - 1]);
                        int b = System.Math.Abs(int.Parse(splitStr[i]));
                        for (int count = a + 1; count <= b; count++)
                        {
                            if (bulidNameToLevels.ContainsKey(count.ToString()))
                            {
                                bulidNameToLevels[count.ToString()]++;
                            }
                            else
                            {
                                bulidNameToLevels[count.ToString()] = 1;
                            }

                        }
                    }
                    else
                    {
                        if (bulidNameToLevels.ContainsKey(splitStr[i]))
                        {
                            bulidNameToLevels[splitStr[i]]++;
                        }
                        else
                        {
                            bulidNameToLevels[splitStr[i]] = 1;
                        }
                    }
                }
                else
                {
                    if (bulidNameToLevels.ContainsKey(splitStr[i]))
                    {
                        bulidNameToLevels[splitStr[i]]++;
                    }
                    else
                    {
                        bulidNameToLevels[splitStr[i]] = 1;
                    }
                }
            }
        }

        private bool NameStringOrNot(string str)
        {
            for (int i = 0; i < str.Length; i++)
            {
                if (!(str[i] > '0' && str[i] < '9'))
                {
                    return false;
                }
            }
            return true;
        }

        //获取楼梯相关栏杆
        private List<ElementId> getStairsRailings(Document doc, Element ele)
        {
            //创建收集器
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            //通用过滤方法--获取楼梯
            ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRailing);
            collector.WherePasses(elementCategoryFilter);
            Dictionary<string, List<ElementId>> stairMap = new Dictionary<string, List<ElementId>>();
            foreach (var item in collector)
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
        
        //获取楼梯间距
        private Dictionary<string, object> getStairsWidth(ICollection<ElementId> runs_Stair, ICollection<ElementId> roliings, Document doc)
        {

            List<int> stairCount = new List<int>();
            List<double> res = new List<double>();
            //楼梯栏杆投影线
            IList<Curve> railingLine = new List<Curve>();
            foreach (ElementId railingId in roliings)
            {
                //Railing railing = stair.Document.GetElement(railingId) as Railing;
                Railing railing = doc.GetElement(railingId) as Railing;

                //获取楼梯所有栏杆的投影线
                IList<Curve> stairline = new List<Curve>();
                stairline = railing.GetPath();
                foreach (Curve i in stairline)
                {
                    railingLine.Add(i);
                }

            }

            //获取梯段
            foreach (ElementId runId in runs_Stair)
            {
                //初始化楼梯宽度
                double weight = 0;
                int flag = -1;
                StairsRun run = doc.GetElement(runId) as StairsRun;
                //获取一个梯段的投影线
                CurveLoop run_curveloops = run.GetFootprintBoundary();
                List<Curve> run_lines = new List<Curve>();
                //一个楼梯梯段最多有两个栏杆
                Curve c1 = null, c2 = null;
                foreach (Curve runline in run_curveloops)
                {
                    run_lines.Add(runline);
                }

                foreach (Curve line in railingLine)
                {
                    if (line_in_rectangle(line, run_lines))
                    {
                        if (c1 == null)
                        {
                            c1 = line;
                        }
                        else
                        {
                            c2 = line;
                        }

                    }
                }

                //当楼梯梯段之内有两个栏杆的时候计算两个栏杆之间的宽度
                if (c1 != null && c2 != null)
                {
                    if (Math.Abs(c1.GetEndPoint(0).X - c1.GetEndPoint(1).X) < 0.1)
                    {
                        weight = Math.Abs(c1.GetEndPoint(0).X - c2.GetEndPoint(1).X);
                    }
                    else
                    {
                        weight = Math.Abs(c1.GetEndPoint(0).Y - c2.GetEndPoint(1).Y);
                    }
                    flag = 2;
                }
                else if (c1 == null && c2 == null)
                {
                    //楼梯梯段没有栏杆
                    weight = run.ActualRunWidth;
                    flag = 0;
                }
                else
                {
                    //楼梯梯段只有一侧栏杆
                    weight = railtowall(c1, run_lines);
                    flag = 1;
                }

                weight *= 0.3048;
                res.Add(weight);
                stairCount.Add(flag);
            }
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("stairCount", stairCount);
            result.Add("weight", res);
            return result;
        }

        //判断线是否在矩形框内
        private bool line_in_rectangle(Curve c, List<Curve> rectangle)
        {

            double t_x, t_y, t_z;
            bool flag = false;
            t_x = (c.GetEndPoint(0).X + c.GetEndPoint(1).X) / 2;
            t_y = (c.GetEndPoint(0).Y + c.GetEndPoint(1).Y) / 2;
            t_z = rectangle[0].GetEndPoint(0).Z;
            XYZ point = new XYZ(t_x, t_y, t_z);
            flag = IsInsideOutline(point, rectangle);
            return flag;
        }

        //判断点是否在包围框之内
        public bool IsInsideOutline(XYZ TargetPoint, List<Curve> lines)
        {
            bool result = true;
            int insertCount = 0;
            Line rayLine = Line.CreateBound(TargetPoint, TargetPoint.Add(XYZ.BasisX * 1000));
            foreach (var areaLine in lines)
            {
                SetComparisonResult interResult = areaLine.Intersect(rayLine, out IntersectionResultArray resultArray);
                IntersectionResult insPoint = resultArray?.get_Item(0);
                if (insPoint != null)
                {
                    insertCount++;
                }
            }
            //如果次数为偶数就在外面，次数为奇数就在里面
            if (insertCount % 2 == 0)//偶数
            {
                return result = false;
            }
            return result;
        }
        
        //计算栏杆到墙的宽度
        private double railtowall(Curve c, List<Curve> rectangle)
        {
            double weight = 0;
            foreach (Curve line in rectangle)
            {
                double temp = 0;
                double vec1_x = Convert.ToDouble(c.GetEndPoint(0).X) - Convert.ToDouble(c.GetEndPoint(1).X);
                vec1_x = Math.Round(vec1_x, 1);
                double vec1_y = Convert.ToDouble(c.GetEndPoint(0).Y) - Convert.ToDouble(c.GetEndPoint(1).Y);
                vec1_y = Math.Round(vec1_y, 1);
                double vec2_x = Convert.ToDouble(line.GetEndPoint(0).X) - Convert.ToDouble(line.GetEndPoint(1).X);
                vec2_x = Math.Round(vec2_x, 1);
                double vec2_y = Convert.ToDouble(line.GetEndPoint(0).Y) - Convert.ToDouble(line.GetEndPoint(1).Y);
                vec2_y = Math.Round(vec2_y, 1);

                if ((vec1_x * vec2_x + vec1_y * vec2_y) != 0)
                {
                    //平行就计算之间的距离

                    if (Math.Abs(c.GetEndPoint(0).X - c.GetEndPoint(1).X) < 0.1)
                    {
                        //竖向线条
                        temp = Math.Abs(c.GetEndPoint(0).X - line.GetEndPoint(1).X);

                    }
                    else
                    {
                        //横向线条
                        temp = Math.Abs(c.GetEndPoint(0).Y - line.GetEndPoint(1).Y);
                    }
                    if (weight < temp)
                    {
                        weight = temp;
                    }

                }
                else continue;


            }

            return weight;
        }
    
        //判断栏杆是否为玻璃
        private bool getRailMaterial(Element el, Document doc)
        {
            GeometryElement geometry = el.get_Geometry(new Options());

            bool flag = false;
            if (geometry == null)
            {
                return false;
            }
            foreach (GeometryObject obj in geometry)
            {
                //得到solid
                GeometryInstance geometryInstance = obj as GeometryInstance;
                if (geometryInstance != null)
                {
                    GeometryElement instanceEle = geometryInstance.GetInstanceGeometry();
                    //得到solid
                    foreach (GeometryObject solidObject in instanceEle)
                    {
                        Solid solid = solidObject as Solid;
                        if (solid != null)
                        {
                            //得到每个面
                            FaceArray faceArray = solid.Faces;
                            foreach (Face face in faceArray)
                            {
                                //得到材料id，再得到材料
                                ElementId materialId = face.MaterialElementId;
                                Material material = doc.GetElement(materialId) as Material;
                                if (material != null)
                                {
                                    if ("玻璃".Equals(material.MaterialCategory))
                                    {
                                        flag = true;
                                        break;
                                    }
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
            return flag;
        }

        public void cal_stairShaft(Document doc, Element el, ref StringBuilder stairWell)
        {
            //楼梯的梯段
            List<ElementId> ls = new List<ElementId>();
            List<ElementId> stairListId = new List<ElementId>();
            stairListId.Add(el.Id);
            Stairs stairs = el as Stairs;
            if (stairs != null)
            {
                ICollection<ElementId> stairsLandings = stairs.GetStairsLandings();
                foreach (ElementId stairLandingId in stairsLandings)
                {
                    StairsLanding stairLanding = doc.GetElement(stairLandingId) as StairsLanding;
                    IList<StairsComponentConnection> stairRuns = stairLanding.GetConnections();
                    if (stairRuns.Count == 1)
                    {
                        //创建收集器
                        FilteredElementCollector collector = new FilteredElementCollector(doc);
                        //通用过滤方法--获取楼梯
                        ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_Stairs);

                        BoundingBoxXYZ boundingBoxXYZ = stairLanding.get_BoundingBox(doc.ActiveView);
                        if (boundingBoxXYZ == null|| boundingBoxXYZ.Max==null|| boundingBoxXYZ.Min == null)
                        {
                            return;
                        }
                        ElementId id = Create_solid(doc, boundingBoxXYZ.Max, boundingBoxXYZ.Min);

                        Element landElement = doc.GetElement(id);

                        ElementIntersectsElementFilter filter = new ElementIntersectsElementFilter(landElement);
                        collector.WherePasses(filter);
                        collector.WherePasses(elementCategoryFilter);


                        if (collector.Count() == 2)
                        {
                            //楼梯的梯段
                            ls.Add(stairRuns[0].PeerElementId);
                            List<Element> stairList = new List<Element>(collector);
                            foreach (Element stair in stairList)
                            {
                                if (stair.Id.ToString() == el.Id.ToString())
                                {
                                    if (!stairList.Remove(stair))
                                    {
                                        return;
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                            //相邻楼梯
                            Stairs nextStair = stairList[0] as Stairs;
                            stairListId.Add(nextStair.Id);

                            ICollection<ElementId> stairsRuns = nextStair.GetStairsRuns();
                            foreach (ElementId stairRunId in stairsRuns)
                            {
                                StairsRun stairsRun = doc.GetElement(stairRunId) as StairsRun;
                                IList<StairsComponentConnection> connectionStairLands = stairsRun.GetConnections();
                                if (connectionStairLands.Count == 1)
                                {
                                    ls.Add(stairRunId);
                                }
                            }
                        }
                        //删除碰撞
                        Transaction tran = new Transaction(doc, "删除");
                        tran.Start();
                        doc.Delete(id);
                        tran.Commit();
                    }
                }
            }
            if (ls.Count != 2)
            {
                return;
            }
            List<double> stairs_w = new List<double>();
            List<double> line_index = new List<double>();
            foreach (ElementId l in ls)
            {
                Element ele = doc.GetElement(l);
                StairsRun stairsRun = ele as StairsRun;
                //判断梯段方向
                CurveLoop lines = stairsRun.GetStairsPath();
                int x_y = -1;
                if (lines.Count() == 1)
                {
                    foreach (Line line in lines)
                    {
                        line_index.Add(line.Direction.X);
                        line_index.Add(line.Direction.Y);
                        if (Math.Abs(line.Direction.X) > 0.9)
                        {
                            x_y = 1;
                        }
                        else
                        {
                            x_y = 0;
                        }
                    }
                }

                BoundingBoxXYZ xyz = ele.get_BoundingBox(doc.ActiveView);

                if (x_y == 0)
                {
                    stairs_w.Add(xyz.Max.X);
                    stairs_w.Add(xyz.Min.X);
                }
                else
                {
                    stairs_w.Add(xyz.Max.Y);
                    stairs_w.Add(xyz.Min.Y);
                }
            }
            //判断平行
            double line_arc = line_index[0] * line_index[2] + line_index[1] * line_index[3];
            if (Math.Abs(line_arc) > 0.01)
            {
                stairs_w.Sort();
                double ww = (stairs_w[2] - stairs_w[1]) * 0.3048;
                //判断踏井宽度是否符合规范
                if (ww > 0.11)
                {

                    stairWell.Append("(" + stairListId[0] + "," + stairListId[1] + ") ");
                }
            }
            else
            {
                return;
            }

        }

        public static ElementId Create_solid(Document document, XYZ max, XYZ min)
        {
            ElementId id;
            Solid solid;
            double temp = 0.1 / 0.3048;
            min = new XYZ(min.X - temp, min.Y - temp, min.Z - temp);
            max = new XYZ(max.X + temp, max.Y + temp, max.Z + temp);
            using (Transaction tran = new Transaction(document, "拉伸"))
            {
                tran.Start();
                Line line1 = Line.CreateBound(min, new XYZ(max.X, min.Y, min.Z));
                Line line2 = Line.CreateBound(new XYZ(max.X, min.Y, min.Z), new XYZ(max.X, max.Y, min.Z));
                Line line3 = Line.CreateBound(new XYZ(max.X, max.Y, min.Z), new XYZ(min.X, max.Y, min.Z));
                Line line4 = Line.CreateBound(new XYZ(min.X, max.Y, min.Z), min);
                CurveLoop loop = new CurveLoop();
                loop.Append(line1);
                loop.Append(line2);
                loop.Append(line3);
                loop.Append(line4);
                List<CurveLoop> loops = new List<CurveLoop>() { loop };
                XYZ direction = new XYZ(0, 0, 1);//拉伸方向
                solid = GeometryCreationUtilities.CreateExtrusionGeometry(loops, direction, max.Z - min.Z);
                DirectShape shape = DirectShape.CreateElement(document, new ElementId(BuiltInCategory.OST_StructuralFoundation));
                shape.AppendShape(new List<GeometryObject>() { solid });
                id = shape.Id;
                tran.Commit();
            }
            return id;
        }
    }
    
}
