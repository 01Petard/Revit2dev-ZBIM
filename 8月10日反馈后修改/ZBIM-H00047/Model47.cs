using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ZBIMUtils;

namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    public class Model47 
    {
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
        //跳到一个名字中包含“viewPlan_name”的平面视图
        public static ViewPlan get2DView(Autodesk.Revit.DB.Document document, string viewPlan_name)
        {
            ViewPlan workView = null;
            Type type = document.ActiveView.GetType();
            if (!type.Equals(typeof(ViewPlan)))
            {
                ViewPlan viewPlan = null;
                //过滤所有的2D视图
                FilteredElementCollector collector = new FilteredElementCollector(document).OfClass(typeof(ViewPlan));
                foreach (ViewPlan view in collector)
                {
                    if (view.Name.Contains(viewPlan_name))
                    {
                        viewPlan = view;
                        workView = viewPlan;
                    }
                }

            }
            return workView;
        }
        public static void switchto2D(UIDocument uiDocument, Autodesk.Revit.DB.Document document, string viewPlan_name)
        {
            Type type = document.ActiveView.GetType();
            if (!type.Equals(typeof(ViewPlan)))
            {
                ViewPlan viewPlan = null;
                //过滤所有的2D视图
                FilteredElementCollector collector = new FilteredElementCollector(document).OfClass(typeof(ViewPlan));
                foreach (ViewPlan view in collector)
                {
                    if (view.Name.Contains(viewPlan_name))
                    {
                        viewPlan = view;
                    }
                }
                uiDocument.ActiveView = viewPlan;//设置所取得的3D视图为当前视图
            }
        }
        //跳到一个名字中包含“viewPlan_name”的三维视图
        public static void switchto3D(UIDocument uiDocument, Autodesk.Revit.DB.Document document, string viewPlan_name)
        {
            Type type = document.ActiveView.GetType();
            if (!type.Equals(typeof(View3D)))
            {
                View3D view3D = null;
                //过滤所有的3D视图
                FilteredElementCollector collector = new FilteredElementCollector(document).OfClass(typeof(View3D));
                foreach (View3D view in collector)
                {
                    if (view.Name.Contains(viewPlan_name))
                    {
                        view3D = view;
                    }
                }
                uiDocument.ActiveView = view3D;//设置所取得的3D视图为当前视图
            }
        }
        //获得每栋建筑中某种类型的面积分区
        public static Dictionary<string, List<Area>> GetBuildArea(Autodesk.Revit.DB.Document document, string AreaTypeName, List<Element> elementsGrids)
        {
            Dictionary<string, List<Area>> FireArea_buildNum = new Dictionary<string, List<Area>>();
            //根据轴网获取了所有的可能的楼栋
            foreach (var ele in elementsGrids)
            {
                string build_num_might = Utils.StartSubString(ele.Name, 0, "-");
                if (build_num_might != "" || build_num_might != null || build_num_might != " ")
                {
                    if (!FireArea_buildNum.ContainsKey(build_num_might))
                    {
                        FireArea_buildNum.Add(build_num_might, new List<Area>());
                    }
                }
            }
            FireArea_buildNum.Add("B1", new List<Area>());
            FireArea_buildNum.Add("B2", new List<Area>());
            FireArea_buildNum.Add("B3", new List<Area>());
            FireArea_buildNum.Add("B4", new List<Area>());
            FireArea_buildNum.Add("B5", new List<Area>());
            FireArea_buildNum.Add("b1", new List<Area>());
            FireArea_buildNum.Add("b2", new List<Area>());
            FireArea_buildNum.Add("b3", new List<Area>());
            FireArea_buildNum.Add("b4", new List<Area>());
            FireArea_buildNum.Add("b5", new List<Area>());
            int area_num = 0;
            string area_buildnum;
            //过滤得到所有面积分区Area，然后按照名字过滤，得到指定类型的面积分区
            FilteredElementCollector AreaCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Areas);
            foreach (Area area in AreaCollector)
            {
                if (area.AreaScheme.Name.Contains(AreaTypeName))
                {
                    if (area.Name.Contains("B"))
                    {
                        area_buildnum = Utils.LastSubEndString(area.Name, AreaTypeName);//先分离出防火分区后面的字段，例如“B1-1 1”
                        area_buildnum = Utils.StartSubString(area_buildnum, 0, "-", false, true);//然后获得楼号，例如“B1”
                        FireArea_buildNum[area_buildnum].Add(area);
                        //stringBuilder.AppendLine("防火分区Id：" + area.Id + "，名称：" + area.Name + "，楼号：" + area_buildnum + "#");
                        area_num++;
                    }
                    else
                    {
                        area_buildnum = Utils.LastSubEndString(area.Name, AreaTypeName);//先分离出防火分区后面的字段，例如“1#-2F-1 10”
                        area_buildnum = Utils.StartSubString(area_buildnum, 0, "#-");//然后获得楼号，例如“1”
                        FireArea_buildNum[area_buildnum].Add(area);
                        //stringBuilder.AppendLine("防火分区Id：" + area.Id + "，名称：" + area.Name + "，楼号：" + area_buildnum + "#");
                        area_num++;
                    }
                }
            }
            return FireArea_buildNum;
        }
        //查看每栋建筑中某种类型的面积分区
        public static void ShowBuildArea(Dictionary<string, List<Area>> FireArea_buildNum, StringBuilder stringBuilder)
        {
            stringBuilder.AppendLine("查看每栋楼每层的面积分区Area：");
            foreach (var item in FireArea_buildNum)
            {
                stringBuilder.AppendLine($"楼号：{item.Key}#");
                foreach (Area area in item.Value)
                {
                    string area_firearea = area.LookupParameter("面积").AsValueString();// 得到area的面积，类似这种字段：“114514m²”
                    if (area.Name.Contains("B") || area.Name.Contains("Y") || area.Name.Contains("S"))
                    {
                        string area_level = Utils.LastSubEndString(area.Name, "-");
                        stringBuilder.AppendLine("第" + area_level + "个：面积分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea);
                    }
                    else
                    {
                        string area_level = Utils.SubBetweenString(area.Name, "-", "F");
                        stringBuilder.AppendLine("第" + area_level + "层：面积分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea);
                    }
                }
            }
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
        //查看某一种的面积分区的所有视图
        public static void ShowAreaView(List<ViewPlan> AreaViewPlan, StringBuilder stringBuilder)
        {
            foreach (var view in AreaViewPlan)
            {
                stringBuilder.AppendLine("视图名：" + view.Name + "，视图ID：" + view.Id + "，视图类型：" + view.LookupParameter("类型").AsValueString());
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
                        int j1 = Convert.ToInt16(Utils.StartSubString(elementsGridsSort[jj].Name, 0, "-")); //将数字字符转为int类型比较
                        int j2 = Convert.ToInt16(Utils.StartSubString(elementsGridsSort[jj + 1].Name, 0, "-"));
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
        //查看所有建筑的轴网
        public static void ShowBuildGrid(List<Element> elementsGridsSort, StringBuilder stringBuilder)
        {
            foreach (var ele in elementsGridsSort)
            {
                stringBuilder.AppendLine(ele.Name);
            }

        }
        //通过管道的“系统缩写”属性，过滤得到所有这些类型的管道（“系统缩写”通过数组传入，可以过滤好几种“系统缩写”的管道哦！）
        public static List<Pipe> GetPipeByAbbr(Autodesk.Revit.DB.Document document, List<string> PipeAbbrList)
        {
            double FirePipeMaxX; //消防立管的最大X坐标
            double FirePipeMinX; //消防立管的最小X坐标
            double FirePipeMaxY; //消火的最大Y坐标
            double FirePipeMinY; //消防立管的最小Y坐标
            double FirePipeMaxZ; //消防立管的最大Z坐标
            double FirePipeMinZ; //消防立管的最小Z坐标
            double FirePipeXLength; //消防立管的X轴长度
            double FirePipeYLength; //消防立管的Y轴长度
            double FirePipeZLength; //消防立管的Z轴长度

            FilteredElementCollector PipeCollector = new FilteredElementCollector(document).OfClass(typeof(Autodesk.Revit.DB.Plumbing.Pipe)).OfCategory(BuiltInCategory.OST_PipeCurves); //获得所有的管道;
            List<Pipe> PipeList = new List<Pipe>();   //存放指定缩写的管道数组
            int pipe_num = 0;
            foreach (Pipe Pipe in PipeCollector)//从管道中筛选出指定缩写的管道，并放入PipeList中
            {
                string PipeAbbr = Pipe.LookupParameter("系统缩写").AsString();//消防立管的缩写
                foreach (string abbr in PipeAbbrList)
                {
                    if (PipeAbbr.Contains(abbr))
                    {
                        BoundingBoxXYZ FirePipeBoundingBoxXYZ = Pipe.get_BoundingBox(document.ActiveView);
                        XYZ FirePipeXYZMax = FirePipeBoundingBoxXYZ.Max; //取到Max元组
                        XYZ FirePipeXYZMin = FirePipeBoundingBoxXYZ.Min; //取到Min元组
                        FirePipeMaxX = FirePipeXYZMax.X;
                        FirePipeMinX = FirePipeXYZMin.X;
                        FirePipeMaxY = FirePipeXYZMax.Y;
                        FirePipeMinY = FirePipeXYZMin.Y;
                        FirePipeMaxZ = Math.Round(Utils.FootToMeter(FirePipeXYZMax.Z), 2);
                        FirePipeMinZ = Math.Round(Utils.FootToMeter(FirePipeXYZMin.Z), 2);
                        FirePipeXLength = FirePipeMaxX - FirePipeMinX;
                        FirePipeYLength = FirePipeMaxY - FirePipeMinY;
                        FirePipeZLength = FirePipeMaxZ - FirePipeMinZ;
                        if (FirePipeMaxZ >= 0 && FirePipeMinZ >= 0)//首先在地坪上
                        {
                            if (FirePipeZLength > 2.0 || FirePipeYLength > 0.5 || FirePipeXLength > 0.5)//考虑竖管和横管
                            {
                                PipeList.Add(Pipe);
                                pipe_num++;
                                //stringBuilder.AppendLine("找到管道：" + FirePipe.Id + "，" + FirePipe.LevelId + "，" + abbr + "，X：（" + FirePipeMinX + "，" + FirePipeMaxX + "），Y：（" + FirePipeMinY + "，" + FirePipeMaxY + "），Z：(" + FirePipeMinZ + "，" + FirePipeMaxZ + ")");
                            }
                        }
                    }
                }

            }
            return PipeList;
        }
        //根据过滤好的管道数组，判断每栋建筑有哪些管道
        public static Dictionary<string, List<Pipe>> GetBuildPipe(Autodesk.Revit.DB.Document document, List<Pipe> PipeList, List<Element> elementsGridsSort, List<Element> elementsGrids)
        {
            double buildMinX = 0; //建筑最小的X坐标
            double buildMinY = 0; //建筑最小的Y坐标
            double buildMaxX = 0; //建筑最大的X坐标
            double buildMaxY = 0; //建筑最大的Y坐标
            double buildMinZ = 0; //建筑最大的Z坐标
            double buildMaxZ = 0; //建筑最大的Z坐标

            double PipeMaxX; //消防立管的最大X坐标
            double PipeMinX; //消防立管的最小X坐标
            double PipeMaxY; //消火的最大Y坐标
            double PipeMinY; //消防立管的最小Y坐标
            double PipeMaxZ; //消防立管的最大Z坐标
            double PipeMinZ; //消防立管的最小Z坐标
            double PipeXLength; //消防立管的X轴长度
            double PipeYLength; //消防立管的Y轴长度
            double PipeZLength; //消防立管的Z轴长度

            Dictionary<string, List<Pipe>> Pipe_buildNum = new Dictionary<string, List<Pipe>>();
            //根据轴网获取了所有的可能的楼栋
            foreach (var ele in elementsGrids)
            {
                string build_num_might = Utils.StartSubString(ele.Name, 0, "-");
                if (build_num_might != "" || build_num_might != null || build_num_might != " ")
                {
                    if (!Pipe_buildNum.ContainsKey(build_num_might))
                    {
                        Pipe_buildNum.Add(build_num_might, new List<Pipe>());
                    }
                }
            }
            Pipe_buildNum.Add("B1", new List<Pipe>());
            Pipe_buildNum.Add("B2", new List<Pipe>());
            Pipe_buildNum.Add("B3", new List<Pipe>());
            Pipe_buildNum.Add("B4", new List<Pipe>());
            Pipe_buildNum.Add("B5", new List<Pipe>());
            Pipe_buildNum.Add("b1", new List<Pipe>());
            Pipe_buildNum.Add("b2", new List<Pipe>());
            Pipe_buildNum.Add("b3", new List<Pipe>());
            Pipe_buildNum.Add("b4", new List<Pipe>());
            Pipe_buildNum.Add("b5", new List<Pipe>());
            string item_name;
            string pre_item_name = "0";

            //Dictionary<Element, List<double>> FirePipeXYZ = new Dictionary<Element, List<double>>();
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
                    item_name = Utils.StartSubString(item.Name, 0, "-");
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
                        foreach (var Pipe in PipeList)
                        {
                            BoundingBoxXYZ FirePipeBoundingBoxXYZ = Pipe.get_BoundingBox(document.ActiveView);
                            XYZ FirePipeXYZMax = FirePipeBoundingBoxXYZ.Max; //取到Max元组
                            XYZ FirePipeXYZMin = FirePipeBoundingBoxXYZ.Min; //取到Min元组
                            PipeMinX = FirePipeXYZMin.X;
                            PipeMaxX = FirePipeXYZMax.X;
                            PipeMinY = FirePipeXYZMin.Y;
                            PipeMaxY = FirePipeXYZMax.Y;
                            PipeMinZ = FirePipeXYZMin.Z;
                            PipeMaxZ = FirePipeXYZMax.Z;
                            PipeXLength = Math.Round(Utils.FootToMeter(Math.Abs(PipeMaxX - PipeMinX)), 2);
                            PipeYLength = Math.Round(Utils.FootToMeter(Math.Abs(PipeMaxY - PipeMinY)), 2);
                            PipeZLength = Math.Round(Utils.FootToMeter(Math.Abs(PipeMaxZ - PipeMinZ)), 2);
                            //stringBuilder.AppendLine("FirePipe.LevelId：" + FirePipe.LevelId + "，FirePipe.Name：" + FirePipe.Name + "，FirePipeAbbr：" + FirePipe.LookupParameter("系统缩写").AsString() + "，FirePipe.Id：" + FirePipe.Id + "，方位：" + "X:(" + FirePipeMinX + "," + FirePipeMaxX + ")" + "，Y:(" + FirePipeMinY + "," + FirePipeMaxY + ")" + "，Z:(" + FirePipeMinZ + "," + FirePipeMaxZ + ")" + "，Z轴长度：" + FirePipeZLength + "m");
                            //stringBuilder.AppendLine($"{FirePipeXLength},{FirePipeYLength},{FirePipeZLength}");
                            //stringBuilder.AppendLine($"管道：({Math.Round(FirePipeMinX,2)},{Math.Round(FirePipeMaxX,2)}),({Math.Round(FirePipeMinY,2)},{Math.Round(FirePipeMaxY,2)}),({Math.Round(FirePipeMinZ,2)},{Math.Round(FirePipeMaxZ,2)})");
                            //stringBuilder.AppendLine("现在item.Name是：" + item.Name + "，pre_item_name是：" + pre_item_name);
                            //stringBuilder.AppendLine($"楼栋：({Math.Round(buildMinX,2)},{Math.Round(buildMaxX, 2)}),({Math.Round(buildMinY, 2)},{Math.Round(buildMaxY, 2)}),({Math.Round(buildMinZ, 2)},{Math.Round(buildMaxZ,2)})");
                            if (PipeMinX > buildMinX && PipeMinY > buildMinY && PipeMaxX < buildMaxX && PipeMaxY < buildMaxY && !Pipe_buildNum[item_name].Contains(Pipe))
                            {
                                if ((PipeXLength < 1 && PipeYLength < 1 && PipeZLength >= 0.1) || PipeYLength > 0.1 || PipeXLength > 0.1) //筛选出垂直的管道（X,Y偏移量不超过1m，Z轴高度大于0.5m），去掉水平的管道
                                {
                                    Pipe_buildNum[pre_item_name].Add(Pipe);

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
                    //给最后一栋楼匹配管道
                    foreach (var FirePipe in PipeList)
                    {
                        BoundingBoxXYZ FirePipeBoundingBoxXYZ = FirePipe.get_BoundingBox(document.ActiveView);
                        XYZ FirePipeXYZMax = FirePipeBoundingBoxXYZ.Max; //取到Max元组
                        XYZ FirePipeXYZMin = FirePipeBoundingBoxXYZ.Min; //取到Min元组
                        PipeMaxX = FirePipeXYZMax.X;
                        PipeMinX = FirePipeXYZMin.X;
                        PipeMaxY = FirePipeXYZMax.Y;
                        PipeMinY = FirePipeXYZMin.Y;
                        PipeMaxZ = FirePipeXYZMax.Z;
                        PipeMinZ = FirePipeXYZMin.Z;
                        PipeXLength = Math.Round(Utils.FootToMeter(Math.Abs(PipeMaxX - PipeMinX)), 2);
                        PipeYLength = Math.Round(Utils.FootToMeter(Math.Abs(PipeMaxY - PipeMinY)), 2);
                        PipeZLength = Math.Round(Utils.FootToMeter(Math.Abs(PipeMaxZ - PipeMinZ)), 2);
                        //stringBuilder.AppendLine("FirePipe.LevelId：" + FirePipe.LevelId + "，FirePipe.Name：" + FirePipe.Name + "，FirePipeAbbr：" + FirePipe.LookupParameter("系统缩写").AsString() + "，FirePipe.Id：" + FirePipe.Id + "，方位：" + "X:(" + FirePipeMinX + "," + FirePipeMaxX + ")" + "，Y:(" + FirePipeMinY + "," + FirePipeMaxY + ")" + "，Z:(" + FirePipeMinZ + "," + FirePipeMaxZ + ")" + "，Z轴长度：" + FirePipeZLength + "m");
                        //stringBuilder.AppendLine($"{FirePipeXLength},{FirePipeYLength},{FirePipeZLength}");
                        //stringBuilder.AppendLine($"管道：({Math.Round(FirePipeMinX, 2)},{Math.Round(FirePipeMaxX, 2)}),({Math.Round(FirePipeMinY, 2)},{Math.Round(FirePipeMaxY, 2)}),({Math.Round(FirePipeMinZ, 2)},{Math.Round(FirePipeMaxZ, 2)})");
                        //stringBuilder.AppendLine("现在item.Name是：" + item.Name + "，pre_item_name是：" + pre_item_name);
                        //stringBuilder.AppendLine($"楼栋：({Math.Round(buildMinX, 2)},{Math.Round(buildMaxX, 2)}),({Math.Round(buildMinY, 2)},{Math.Round(buildMaxY, 2)}),({Math.Round(buildMinZ, 2)},{Math.Round(buildMaxZ, 2)})");
                        if (PipeMinX > buildMinX && PipeMinY > buildMinY && PipeMaxX < buildMaxX && PipeMaxY < buildMaxY && !Pipe_buildNum[item_name].Contains(FirePipe))
                        {
                            if ((PipeXLength < 1 && PipeYLength < 1 && PipeZLength >= 0.1) || PipeYLength > 0.1 || PipeXLength > 0.1) //筛选出垂直的管道（X,Y偏移量不超过1m，Z轴高度大于0.5m），去掉水平的管道
                            {
                                Pipe_buildNum[pre_item_name].Add(FirePipe);
                                //stringBuilder.AppendLine(pre_item_name + "#：" + FirePipe.Id + "，" + FirePipe.LookupParameter("系统缩写").AsString() + "，X：（" + FirePipeMaxX + "，" + FirePipeMinX + "），Y：（" + FirePipeMaxY + "，" + FirePipeMinY + "），Z：(" + FirePipeMinZ + "，" + FirePipeMaxZ + ")");
                            }
                        }
                    }
                }
            }
            return Pipe_buildNum;
        }
        //查看每栋建筑有哪些管道
        public static void ShowBuildPipe(Autodesk.Revit.DB.Document document, Dictionary<string, List<Pipe>> Pipe_buildNum, StringBuilder stringBuilder)
        {
            //查看每栋楼的管道
            double PipeMaxX; //管道的最大X坐标
            double PipeMinX; //管道的最小X坐标
            double PipeMaxY; //管道最大Y坐标
            double PipeMinY; //管道的最小Y坐标
            double PipeMaxZ; //管道的最大Z坐标
            double PipeMinZ; //管道的最小Z坐标
            double PipeXLength; //管道的X轴长度
            double PipeYLength; //管道的Y轴长度
            double PipeZLength; //管道的Z轴长度
            foreach (var item in Pipe_buildNum)
            {
                stringBuilder.AppendLine($"楼号：{item.Key}#");
                foreach (var FirePipe in item.Value)
                {
                    BoundingBoxXYZ FirePipeBoundingBoxXYZ = FirePipe.get_BoundingBox(document.ActiveView);
                    XYZ FirePipeXYZMax = FirePipeBoundingBoxXYZ.Max; //取到Max元组
                    XYZ FirePipeXYZMin = FirePipeBoundingBoxXYZ.Min; //取到Min元组
                    PipeMaxX = FirePipeXYZMax.X;
                    PipeMinX = FirePipeXYZMin.X;
                    PipeMaxY = FirePipeXYZMax.Y;
                    PipeMinY = FirePipeXYZMin.Y;
                    PipeMaxZ = FirePipeXYZMax.Z;
                    PipeMinZ = FirePipeXYZMin.Z;
                    stringBuilder.AppendLine("找到管道：" + FirePipe.Id + "，" + FirePipe.LevelId + "，" + FirePipe.LookupParameter("系统缩写").AsString() + "，X：（" + PipeMinX + "，" + PipeMaxX + "），Y：（" + PipeMinY + "，" + PipeMaxY + "），Z：(" + PipeMinZ + "，" + PipeMaxZ + ")");

                }
            }
        }
        public static List<ElementId> GetAreaViewId(Document document, string AreaTypeName, bool isContainBasement = true)//获得某种类型面积分区的视图View
        {
            List<ElementId> AreaViewPlan = new List<ElementId>();
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
                                AreaViewPlan.Add(viewPlan.Id);
                            }
                            else
                            {
                                if (!viewPlan.Name.Contains("B"))
                                {
                                    //stringBuilder.AppendLine("视图名：" + viewPlan.Name + "，视图ID：" + viewPlan.Id + "，视图类型：" + viewPlan_category);
                                    //viewPlan2D = viewPlan;
                                    AreaViewPlan.Add(viewPlan.Id);
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        //有些视图类型不是通过AsValueString属性获得的，因此这一步需要处理异常
                        //stringBuilder.AppendLine(e.Message);
                    }
                }
            }
            return AreaViewPlan;
        }


        //public bool CoreExecute(ExternalCommandData commandData, List<string> stringList)
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            List<string> stringList = new List<string>();

            //界面交互的document
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            //实际内容的document
            Autodesk.Revit.DB.Document document = commandData.Application.ActiveUIDocument.Document;

            //切换到二维平面
            //switchto2D(uiDocument,document);

            // 轴网收集器
            FilteredElementCollector collectorGrid = new FilteredElementCollector(document);
            //过滤出每栋楼的轴网
            List<Element> elementsGrids = collectorGrid.OfCategory(BuiltInCategory.OST_Grids).ToList<Element>();

            View viewPlan1F = null;
            foreach (ViewPlan view in new FilteredElementCollector(document).OfClass(typeof(ViewPlan)))
            {
                if (view.Name.Contains("#1F")) { viewPlan1F = view; }
            }
            bool switched = true;
            //Utils.SwitchTo2D(document, uiDocument, ref viewPlan1F, ref switched);

            StringBuilder stringBuilder = new StringBuilder();

            //==========防火分区FireArea==========
            Dictionary<string, List<Area>> FireArea_buildNum = GetBuildArea(document, "防火分区", elementsGrids);
            //查看每栋楼防火分区
            //ShowBuildArea(FireArea_buildNum, stringBuilder);


            //==========防火分区的ViewPlan==========
            //List<ViewPlan> FireAreaViewPlan = GetAreaView(document, "防火分区", false);
            List<ElementId> FireAreaViewPlan = GetAreaViewId(document, "防火分区", false);
            //查看防火分区的ViewPlan
            //ShowAreaView(FireAreaViewPlan, stringBuilder);


            //==========轴网过滤、排序==========
            List<Element> elementsGridsSort = GetBuildGrid(document);
            //查看排序后的轴网
            //ShowBuildGrid(elementsGridsSort, stringBuilder);


            //==========获得喷淋管==========
            List<string> PipeAbbrList = new List<string> { "ZP", "PL" };
            List<Pipe> FirePipeList = GetPipeByAbbr(document, PipeAbbrList);


            //==========找到每个轴网范围内的喷淋管道==========
            Dictionary<string, List<Pipe>> FirePipe_buildNum = GetBuildPipe(document, FirePipeList, elementsGridsSort, elementsGrids);
            //查看每栋楼的喷淋管
            //ShowBuildPipe(document, FirePipe_buildNum, stringBuilder);



            //存储符合与不符合国标的输出
            //List<string> fit = new List<string>();
            //List<string> not_fit = new List<string>();

            List<Element> elementGridsList = new List<Element>();

            ElementId viewid = document.ActiveView.Id;

            FilteredElementCollector collectorLevels = new FilteredElementCollector(document, viewid);
            FilteredElementCollector collectorAreas = new FilteredElementCollector(document, viewid);
            FilteredElementCollector collectorRooms = new FilteredElementCollector(document, viewid);
            FilteredElementCollector collectorViews = new FilteredElementCollector(document, viewid);

            FilteredElementCollector collectorViewSchedules = new FilteredElementCollector(document);
            FilteredElementCollector collectorViewSections = new FilteredElementCollector(document);
            FilteredElementCollector collectorViewPlans = new FilteredElementCollector(document);

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
                    buildNum = Utils.StartSubString(item.Name, 0, "-", false, false);
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
                    buildNum = Utils.StartSubString(item.Name, 0, "-", false, false);
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

            //Utils.CloseCurrentView(uiDocument);
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
                                up_levelname = Utils.StartSubString(levelitem_Name, 0, "#-", false, false);
                                if (up_levelname.Contains("、"))
                                {
                                    up_levelname = Utils.LastSubEndString(up_levelname, "、", false, false);
                                }
                                before = int.Parse(up_levelname);

                                up_levelname = Utils.SubBetweenString(levelitem_Name, "#-", "#", false, false, false);
                                after = int.Parse(up_levelname);

                                if (after - before <= 0)
                                {
                                    stringList.Add("标高命名有误\n");
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
                                stringList.Add("标高命名有误\n");
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
                            if ((up_levelname.Contains("屋面") || up_levelname.Contains("屋顶") || up_levelname.Contains("屋面层") || up_levelname.Contains("屋顶层")) &&
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
                        else { }
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
                            levelName_str = Utils.SubBetweenString(levelName_str, "#", "F", false, true, true);
                        }
                        levelName_str = Utils.StartSubString(levelName_str, 0, "F", false, true);
                        int levelName_int = Convert.ToInt16(levelName_str);

                        while (levelName_str_next.Contains("#"))
                        {
                            levelName_str_next = Utils.SubBetweenString(levelName_str_next, "#", "F", false, true, true);
                        }
                        levelName_str_next = Utils.StartSubString(levelName_str_next, 0, "F", false, true);
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
                    //not_fit.Add("楼栋编号：" + buildFlag + "#");
                    //not_fit.Add("找不到室外地坪，无法判断");
                    //not_fit.Add("");
                    continue;
                }
                else if (parapet == null)
                {
                    //stringBuilder.AppendLine("找不到女儿墙，无法判断");
                    //not_fit.Add("楼栋编号：" + buildFlag + "#");
                    //not_fit.Add("找不到女儿墙，无法判断");
                    //not_fit.Add("");
                    continue;
                }
                else if (roof == null)
                {
                    //stringBuilder.AppendLine("找不到屋面，无法判断");
                    //not_fit.Add("楼栋编号：" + buildFlag + "#");
                    //not_fit.Add("找不到屋面，无法判断");
                    //not_fit.Add("");
                    continue;
                }
                else
                {
                    build_height = Math.Abs(Utils.FootToMeter(outdoor_floor.Elevation) - Utils.FootToMeter(roof.Elevation));
                }

                //输出建筑的高度
                stringBuilder.AppendLine("");
                //stringBuilder.Append(buildFlag + "建筑高度");
                //stringBuilder.Append("楼栋编号：" + buildFlag + "#，建筑高度：" + Math.Round(build_height, 2) + "m");
                stringBuilder.Append(buildFlag + "# 建筑高度：" + Math.Round(build_height, 2) + "m");
                stringList.Add(buildFlag + "# 建筑高度：" + Math.Round(build_height, 2) + "m");

                if (build_height <= 27)
                {
                    stringList.Add("（≤27m），单、多层民用建筑");
                    stringBuilder.AppendLine("（≤27m），单、多层民用建筑");
                }
                else if (build_height > 27 && build_height <= 54)
                {
                    stringList.Add("（27-54m），二类高层住宅");
                    stringBuilder.AppendLine("（27-54m），二类高层住宅");
                }
                else
                {
                    stringList.Add("（＞54m），一类高层住宅");
                    stringBuilder.AppendLine("（＞54m），一类高层住宅");
                }


                StringBuilder stringBuilder_pass = new StringBuilder();
                StringBuilder stringBuilder_notpass = new StringBuilder();
                List<String> str_pass=new List<string>();
                List<String> str_notpass = new List<string>();

                //没排序前的楼层号存放数组
                List<int> string_pass = new List<int>();
                List<int> string_notpass = new List<int>();
                //没排序前每层楼的检测结果内容
                List<string> string_pass_result = new List<string>();
                List<string> string_notpass_result = new List<string>();

                //排序后，将每层楼和楼的检测结果内容一一结合后，再输出
                string level_name_num;
                double FireAreaMinX;
                double FireAreaMaxX;
                double FireAreaMinY;
                double FireAreaMaxY;
                double FireAreaMinZ;
                double FireAreaMaxZ;

                double FirePipeMaxX; //消防立管的最大X坐标
                double FirePipeMinX; //消防立管的最小X坐标
                double FirePipeMaxY; //消火的最大Y坐标
                double FirePipeMinY; //消防立管的最小Y坐标
                double FirePipeMaxZ; //消防立管的最大Z坐标
                double FirePipeMinZ; //消防立管的最小Z坐标
                double FirePipeXLength; //消防立管的X轴长度
                double FirePipeYLength; //消防立管的Y轴长度
                double FirePipeZLength; //消防立管的Z轴长度

                View CurrnetView = null;


                double level_height = 0;
                //排序后，将每层楼和楼的检测结果内容一一结合后，再输出
                if (build_height > 27)
                {
                    stringList.Add("不符合的防火分区：\n");
                    foreach (Level litem in build_upfloorsNumList)
                    { //从这里开始遍历建筑的楼层
                        //levelNum++;
                        //stringBuilder.AppendLine("litem.Name:" + litem.Name + "，litem.Id:" + litem.Id);
                        string level_name = litem.Name; //仅用于记录楼层的名字，1#、2#、3#
                        while (level_name.Contains("#")) { level_name = Utils.SubBetweenString(level_name, "#", "F", false, true, true); }
                        level_name_num = Utils.StartSubString(level_name, 0, "F");
                        if (litem.Name.Contains("避难层")) { level_height = 56.4; }
                        else
                        {
                            //利用楼层的“名称”属性获得高度，然后将String转成Double，就是真正的楼层对地高度
                            //获得小数点"."后面的一个字符
                            int point_idx = litem.Name.IndexOf(".");
                            string nextpoint = litem.Name.Substring(point_idx + 3, 2);//获得小数点后三个起始的字符，作为SubBetweenString的截断字符，长度只需取2
                            level_height = Convert.ToDouble(Utils.SubBetweenString(litem.Name, "F", nextpoint, false, false));
                        }

                        //判断防火分区内是否存在喷淋管
                        foreach (Area area in FireArea_buildNum[buildFlag])
                        {
                            bool exist_FirePipe = false;
                            double area_firearea = Convert.ToDouble(area.LookupParameter("面积").AsValueString()); //得到Area指向的防火分区的面积，类似这种字段：“114514”
                            string area_level = Utils.SubBetweenString(area.Name, "-", "F");
                            if (area_level == level_name_num)
                            {
                                foreach (ElementId viewPlanId in FireAreaViewPlan)
                                {
                                    try
                                    {
                                        //拿到防火分区的坐标



                                        BoundingBoxXYZ FireAreaBoundingBoxXYZ = area.get_BoundingBox(document.ActiveView);
                                        FireAreaMinX = area.get_BoundingBox(document.ActiveView).Min.X;
                                        FireAreaMaxX = area.get_BoundingBox(document.ActiveView).Max.X;
                                        FireAreaMinY = area.get_BoundingBox(document.ActiveView).Min.Y;
                                        FireAreaMaxY = area.get_BoundingBox(document.ActiveView).Max.Y;
                                        FireAreaMinZ = area.get_BoundingBox(document.ActiveView).Min.Z;
                                        FireAreaMaxZ = area.get_BoundingBox(document.ActiveView).Max.Z;
                                        //stringBuilder.AppendLine(area_level + "F：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡");
                                        //stringBuilder.AppendLine("X：(" + Math.Round(FireAreaMinX, 2) + "，" + Math.Round(FireAreaMaxX, 2) + ")，Y：(" + Math.Round(FireAreaMinY, 2) + "，" + Math.Round(FireAreaMaxY, 2) + ")，Z：(" + Math.Round(FireAreaMinZ, 2) + "，" + Math.Round(FireAreaMaxZ, 2) + ")");
                                        //拿到喷淋管的坐标
                                        //stringBuilder.AppendLine(FirePipe_buildNum[buildFlag].Count.ToString());
                                        foreach (Element FirePipe in FirePipe_buildNum[buildFlag])
                                        {
                                            BoundingBoxXYZ FirePipeBoundingBoxXYZ = FirePipe.get_BoundingBox(document.ActiveView);
                                            FirePipeMinX = FirePipe.get_BoundingBox(document.ActiveView).Min.X;
                                            FirePipeMaxX = FirePipe.get_BoundingBox(document.ActiveView).Max.X;
                                            FirePipeMinY = FirePipe.get_BoundingBox(document.ActiveView).Min.Y;
                                            FirePipeMaxY = FirePipe.get_BoundingBox(document.ActiveView).Max.Y;
                                            FirePipeMinZ = FirePipe.get_BoundingBox(document.ActiveView).Min.Z;
                                            FirePipeMaxZ = FirePipe.get_BoundingBox(document.ActiveView).Max.Z;
                                            //stringBuilder.AppendLine("找到管道：" + FirePipe.Id + "，" + FirePipe.LevelId + "，" + "，X：（" + FirePipeMinX + "，" + FirePipeMaxX + "），Y：（" + FirePipeMinY + "，" + FirePipeMaxY + "），Z：(" + FirePipeMinZ + "，" + FirePipeMaxZ + ")");
                                            //stringBuilder.AppendLine($"喷淋管Id：{FirePipe.Id}，名称：{FirePipe.Name}，X：({Math.Round(FirePipeMinX, 2)}，{Math.Round(FirePipeMaxX, 2)})，Y：({Math.Round(FirePipeMinY, 2)}，{Math.Round(FirePipeMaxY, 2)})，Z：({Math.Round(FirePipeMinZ, 2)}，{Math.Round(FirePipeMaxZ, 2)})，缩写：{FirePipe.LookupParameter("系统缩写").AsString()}");
                                            FirePipeZLength = Math.Round(Utils.FootToMeter(Math.Abs(FirePipeMaxZ - FirePipeMinZ)), 2);
                                            //如果喷淋管在防火分区的高度范围和坐标范围内，就认为此时的防火分区内存在喷淋管
                                            if ((FirePipeMaxZ > FireAreaMinZ && FirePipeMinZ < FireAreaMinZ) || (FirePipeMaxZ < FireAreaMaxZ && FirePipeMinZ > FireAreaMinZ))
                                            {

                                                if (FireAreaMinX < FirePipeMinX && FireAreaMinY < FirePipeMinY && FireAreaMaxX > FirePipeMaxX && FireAreaMaxY > FirePipeMaxY)
                                                {
                                                    exist_FirePipe = true;
                                                    //stringBuilder.AppendLine($"喷淋管Id：{FirePipe.Id}，名称：{FirePipe.Name}，X：({Math.Round(FirePipeXYZ[FirePipe][0], 2)}，{Math.Round(FirePipeXYZ[FirePipe][1], 2)})，Y：({Math.Round(FirePipeXYZ[FirePipe][2], 2)}，{Math.Round(FirePipeXYZ[FirePipe][3], 2)})，Z：({Math.Round(FirePipeXYZ[FirePipe][4], 2)}，{Math.Round(FirePipeXYZ[FirePipe][5], 2)})，缩写：{FirePipe.LookupParameter("系统缩写").AsString()}");
                                                }
                                            }
                                            //Utils.SwitchTo2D(document, uiDocument, ref CurrnetView, ref switched);
                                            Utils.CloseCurrentView(uiDocument);
                                        }
                                    }
                                    catch { }
                                }
                               // Utils.CloseCurrentView(uiDocument);
                                if (exist_FirePipe == true)
                                {
                                    //stringBuilder.AppendLine("防火分区存在喷淋管");
                                    if (area_firearea >= 3000)
                                    {
                                        H00047_pass = false;
                                        /*stringList.Add(area_level + "F：防火分区Id：");
                                        stringList.Add(area.Id.ToString());
                                        stringList.Add("，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡" + "\n");
*/
                                        stringBuilder_notpass.AppendLine(area_level + "F：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡");
                                        str_notpass.Add(area_level + "F：防火分区Id：");
                                        str_notpass.Add(area.Id.ToString());
                                        str_notpass.Add("，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡" + "\n");
                                    }
                                    else
                                    {
                                        stringBuilder_pass.AppendLine(area_level + "F：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡");
                                        str_pass.Add(area_level + "F：防火分区Id：");
                                        str_pass.Add(area.Id.ToString());
                                        str_pass.Add("，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡"+"\n");

                                    }
                                }
                                else
                                {
                                    //stringBuilder.AppendLine("防火分区没有喷淋管");
                                    if (area_firearea >= 1500)
                                    {
                                        H00047_pass = false;
                                        str_notpass.Add(area_level + "F：防火分区Id：");
                                        str_notpass.Add(area.Id.ToString());
                                        str_notpass.Add("，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡" + "\n");
                                        stringBuilder_notpass.AppendLine(area_level + "F：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡");
                                    }
                                    else
                                    {
                                        stringBuilder_pass.AppendLine(area_level + "F：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡");
                                        str_pass.Add(area_level + "F：防火分区Id：");
                                        str_pass.Add(area.Id.ToString());
                                        str_pass.Add("，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡"+"\n");
                                    }
                                }
                            }
                        }
                    }
                    stringBuilder.Append("不符合的防火分区：\n" + stringBuilder_notpass);
                    stringBuilder.Append("符合的防火分区：\n" + stringBuilder_pass);
                    stringList.Add("不符合的防火分区：\n" );
                    stringList.AddRange(str_notpass);
                    stringList.Add("\n");
                    stringList.Add("符合的防火分区：\n");
                    stringList.AddRange(str_pass);
                    stringList.Add("\n");
                }
                else
                {
                    stringList.Add("不符合的防火分区：\n");
                    foreach (Level litem in build_upfloorsNumList)
                    { //从这里开始遍历建筑的楼层
                        //levelNum++;
                        //stringBuilder.AppendLine("litem.Name:" + litem.Name + "，litem.Id:" + litem.Id);
                        string level_name = litem.Name; //仅用于记录楼层的名字，1#、2#、3#
                        while (level_name.Contains("#")) { level_name = Utils.SubBetweenString(level_name, "#", "F", false, true, true); }
                        level_name_num = Utils.StartSubString(level_name, 0, "F");
                        if (litem.Name.Contains("避难层")) { level_height = 56.4; }
                        else
                        {
                            //利用楼层的“名称”属性获得高度，然后将String转成Double，就是真正的楼层对地高度
                            //获得小数点"."后面的一个字符
                            int point_idx = litem.Name.IndexOf(".");
                            string nextpoint = litem.Name.Substring(point_idx + 3, 2);//获得小数点后三个起始的字符，作为SubBetweenString的截断字符，长度只需取2
                            level_height = Convert.ToDouble(Utils.SubBetweenString(litem.Name, "F", nextpoint, false, false));
                        }

                        //判断防火分区内是否存在喷淋管
                        foreach (Area area in FireArea_buildNum[buildFlag])
                        {
                            bool exist_FirePipe = false;
                            double area_firearea = Convert.ToDouble(area.LookupParameter("面积").AsValueString()); //得到Area指向的防火分区的面积，类似这种字段：“114514”
                            string area_level = Utils.SubBetweenString(area.Name, "-", "F");
                            if (area_level == level_name_num)
                            {
                                foreach (ElementId viewPlanId in FireAreaViewPlan)

                                {
                                    try
                                    {
                                        //拿到防火分区的坐标

                                        BoundingBoxXYZ FireAreaBoundingBoxXYZ = area.get_BoundingBox(document.ActiveView);
                                        FireAreaMinX = area.get_BoundingBox(document.ActiveView).Min.X;
                                        FireAreaMaxX = area.get_BoundingBox(document.ActiveView).Max.X;
                                        FireAreaMinY = area.get_BoundingBox(document.ActiveView).Min.Y;
                                        FireAreaMaxY = area.get_BoundingBox(document.ActiveView).Max.Y;
                                        FireAreaMinZ = area.get_BoundingBox(document.ActiveView).Min.Z;
                                        FireAreaMaxZ = area.get_BoundingBox(document.ActiveView).Max.Z;
                                        //stringBuilder.AppendLine(area_level + "F：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡");
                                        //stringBuilder.AppendLine("X：(" + Math.Round(FireAreaMinX, 2) + "，" + Math.Round(FireAreaMaxX, 2) + ")，Y：(" + Math.Round(FireAreaMinY, 2) + "，" + Math.Round(FireAreaMaxY, 2) + ")，Z：(" + Math.Round(FireAreaMinZ, 2) + "，" + Math.Round(FireAreaMaxZ, 2) + ")");
                                        //拿到喷淋管的坐标
                                        //stringBuilder.AppendLine(FirePipe_buildNum[buildFlag].Count.ToString());
                                        foreach (Element FirePipe in FirePipe_buildNum[buildFlag])
                                        {
                                            BoundingBoxXYZ FirePipeBoundingBoxXYZ = FirePipe.get_BoundingBox(document.ActiveView);
                                            FirePipeMinX = FirePipe.get_BoundingBox(document.ActiveView).Min.X;
                                            FirePipeMaxX = FirePipe.get_BoundingBox(document.ActiveView).Max.X;
                                            FirePipeMinY = FirePipe.get_BoundingBox(document.ActiveView).Min.Y;
                                            FirePipeMaxY = FirePipe.get_BoundingBox(document.ActiveView).Max.Y;
                                            FirePipeMinZ = FirePipe.get_BoundingBox(document.ActiveView).Min.Z;
                                            FirePipeMaxZ = FirePipe.get_BoundingBox(document.ActiveView).Max.Z;
                                            //stringBuilder.AppendLine("找到管道：" + FirePipe.Id + "，" + FirePipe.LevelId + "，" + "，X：（" + FirePipeMinX + "，" + FirePipeMaxX + "），Y：（" + FirePipeMinY + "，" + FirePipeMaxY + "），Z：(" + FirePipeMinZ + "，" + FirePipeMaxZ + ")");
                                            //stringBuilder.AppendLine($"喷淋管Id：{FirePipe.Id}，名称：{FirePipe.Name}，X：({Math.Round(FirePipeMinX, 2)}，{Math.Round(FirePipeMaxX, 2)})，Y：({Math.Round(FirePipeMinY, 2)}，{Math.Round(FirePipeMaxY, 2)})，Z：({Math.Round(FirePipeMinZ, 2)}，{Math.Round(FirePipeMaxZ, 2)})，缩写：{FirePipe.LookupParameter("系统缩写").AsString()}");
                                            FirePipeZLength = Math.Round(Utils.FootToMeter(Math.Abs(FirePipeMaxZ - FirePipeMinZ)), 2);
                                            //如果喷淋管在防火分区的高度范围和坐标范围内，就认为此时的防火分区内存在喷淋管
                                            if ((FirePipeMaxZ > FireAreaMinZ && FirePipeMinZ < FireAreaMinZ) || (FirePipeMaxZ < FireAreaMaxZ && FirePipeMinZ > FireAreaMinZ))
                                            {

                                                if (FireAreaMinX < FirePipeMinX && FireAreaMinY < FirePipeMinY && FireAreaMaxX > FirePipeMaxX && FireAreaMaxY > FirePipeMaxY)
                                                {
                                                    exist_FirePipe = true;
                                                    //stringBuilder.AppendLine($"喷淋管Id：{FirePipe.Id}，名称：{FirePipe.Name}，X：({Math.Round(FirePipeXYZ[FirePipe][0], 2)}，{Math.Round(FirePipeXYZ[FirePipe][1], 2)})，Y：({Math.Round(FirePipeXYZ[FirePipe][2], 2)}，{Math.Round(FirePipeXYZ[FirePipe][3], 2)})，Z：({Math.Round(FirePipeXYZ[FirePipe][4], 2)}，{Math.Round(FirePipeXYZ[FirePipe][5], 2)})，缩写：{FirePipe.LookupParameter("系统缩写").AsString()}");
                                                }
                                            }
                                        }
                                        //Utils.SwitchTo2D(document, uiDocument, ref CurrnetView, ref switched);
                                        Utils.CloseCurrentView(uiDocument);
                                    }
                                    catch { }
                                }
                               //s Utils.CloseCurrentView(uiDocument);
                                if (exist_FirePipe == true)
                                {
                                    //stringBuilder.AppendLine("防火分区存在喷淋管");
                                    if (area_firearea >= 5000)
                                    {
                                        H00047_pass = false;
                                        str_notpass.Add(area_level + "F：防火分区Id：");
                                        str_notpass.Add(area.Id.ToString());
                                        str_notpass.Add("，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡" + "\n");

                                        stringBuilder_notpass.AppendLine(area_level + "F：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡");
                                    }
                                    else
                                    {
                                        str_pass.Add(area_level + "F：防火分区Id：");
                                        str_pass.Add(area.Id.ToString());
                                        str_pass.Add("，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡" + "\n");
                                        stringBuilder_pass.AppendLine(area_level + "F：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡");
                                    }
                                }
                                else
                                {
                                    //stringBuilder.AppendLine("防火分区没有喷淋管");
                                    if (area_firearea >= 2500)
                                    {
                                        H00047_pass = false;
                                        str_notpass.Add(area_level + "F：防火分区Id：");
                                        str_notpass.Add(area.Id.ToString());
                                        str_notpass.Add("，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡" + "\n");

                                        stringBuilder_notpass.AppendLine(area_level + "F：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡");
                                    }
                                    else
                                    {
                                        str_pass.Add(area_level + "F：防火分区Id：");
                                        str_pass.Add(area.Id.ToString());
                                        str_pass.Add("，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡" + "\n");
                                        stringBuilder_pass.AppendLine(area_level + "F：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea.ToString() + "㎡");
                                    }
                                }
                            }
                        }
                    }
                    stringBuilder.Append("不符合的防火分区：\n" + stringBuilder_notpass);
                    stringBuilder.Append("符合的防火分区：\n" + stringBuilder_pass);
                    stringList.Add("不符合的防火分区：\n");
                    stringList.AddRange(str_notpass);
                   

                    stringList.Add("符合的防火分区：\n");
                    stringList.AddRange(str_pass);
                    stringList.Add("\n");
                }
                //Utils.CloseCurrentView(uiDocument);
            }

            //Utils.CloseCurrentView(uiDocument);

            if (H00047_pass == true)
            {
                H00047_result.AppendLine("符合5.3.1建筑设计防火规范 GB50016-2014（2018年版）");
            }
            else
            {
                H00047_result.AppendLine("不符合5.3.1建筑设计防火规范 GB50016-2014（2018年版）");
            }
            stringBuilder = H00047_result.AppendLine(stringBuilder.ToString());

            //切换到三维平面
           // switchto3D(uiDocument, document, "{三维}");
            //打印输出
           // Utils.PrintLog(stringBuilder.ToString(), "H00047", document);

            Regex regex = new Regex(@"(?<=\\)M[0-9]{5}");
            string modelNum = "M0000x";
            string docPath = document.PathName;
            if (regex.IsMatch(docPath))
            {
                modelNum = regex.Match(docPath).ToString();
            }

           // TaskDialog.Show($"H00047强条{modelNum}号模型检测", stringBuilder.ToString());

            //return H00047_pass;
            return Result.Succeeded;
        }
    }
}
