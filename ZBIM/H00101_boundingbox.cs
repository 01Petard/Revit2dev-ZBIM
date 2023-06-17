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

    public class H00101_boundingbox : IExternalCommand
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
        //获取楼梯相关栏杆
        private static List<ElementId> getStairsRailings(Document doc, Stairs ele, FilteredElementCollector StairsRailingCollector)
        {
            Dictionary<string, List<ElementId>> stairMap = new Dictionary<string, List<ElementId>>();
            foreach (var item in StairsRailingCollector)
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
                            //stairsRailingList.Add(stairsRailing.Id);
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
        private static List<Element> GetItemById(ElementId elementId, FilteredElementCollector elementCollector)
        {
            List<Element> element_list = new List<Element>();

            foreach (Element element in elementCollector) 
            {
                if (element.Id.Equals(elementId))
                {
                    element_list.Add(element);
                    break;
                }
            }
            return element_list;
        }
        private static List<Element> GetItemByIds(List<ElementId> elementIds, FilteredElementCollector elementCollector)
        {
            
            List<ElementId> temp_element_list = elementIds;
            List<Element> element_list = new List<Element>();

            foreach (ElementId elementId in temp_element_list)
            {
                foreach (Element element in elementCollector)
                {
                    if (element.Id.Equals(elementId))
                    {
                        element_list.Add(element);
                    }
                }

            }
            return element_list;
        }
        public static void switchto2D(UIDocument uiDocument, Document document)
        {
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
        }
        public static void switchto3D(UIDocument uiDocument, Document document)
        {
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
        }




        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            Document document = commandData.Application.ActiveUIDocument.Document;
            StringBuilder stringBuilder = new StringBuilder();
            StringBuilder stringBuilder_pass = new StringBuilder();
            StringBuilder stringBuilder_notpass = new StringBuilder(); 
            StringBuilder stringBuilder_might = new StringBuilder();
            StringBuilder stringBuilder_result = new StringBuilder();
            bool H00101 = true;

            //切换到二维平面
            switchto2D(uiDocument, document);


            //过滤：栏杆
            ElementCategoryFilter StairsRailingFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRailing);
            //过滤：楼梯
            //ElementCategoryFilter StairsFilter = new ElementCategoryFilter(BuiltInCategory.OST_Stairs);
            //过滤：平面
            //ElementCategoryFilter StairsLandingsFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsLandings);
            //过滤：梯段
            //ElementCategoryFilter StairsRunsFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRuns);
            //过滤：墙体
            ElementCategoryFilter WallsFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);


            //收集器：栏杆
            FilteredElementCollector StairsRailingCollector = new FilteredElementCollector(document).WherePasses(StairsRailingFilter);
            //FilteredElementCollector StairsRailingCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StairsRailing);
            //收集器：楼梯
            //FilteredElementCollector StairsCollector = new FilteredElementCollector(document).WherePasses(StairsFilter);
            //FilteredElementCollector StairsCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Stairs);
            //收集器：平面
            //FilteredElementCollector StairsLandingsCollector = new FilteredElementCollector(document).WherePasses(StairsLandingsFilter);
            //FilteredElementCollector StairsLandingsCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StairsLandings);
            //收集器：梯段
            //FilteredElementCollector StairsRunsCollector = new FilteredElementCollector(document).WherePasses(StairsRunsFilter);
            //收集器：墙体
            FilteredElementCollector WallsCollector = new FilteredElementCollector(document).WherePasses(WallsFilter);
            //FilteredElementCollector WallsCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Walls);


            /*
            foreach (var item in StairsCollector)
            {
                try
                {
                    Stairs stairs = item as Stairs;
                    List<ElementId> StairsRoliingsId_List = getStairsRailings(document, stairs, StairsRailingCollector);
                    List<ElementId> StairsRunsId_List = stairs.GetStairsRuns().ToList();
                    List<ElementId> StairsLandingsId_List = stairs.GetStairsLandings().ToList();

                    List<Element> StairsRoliings_List = GetItemByIds(StairsRoliingsId_List, StairsRailingCollector);
                    List<Element> StairsRuns_List = GetItemByIds(StairsRunsId_List, StairsRunsCollector);
                    List<Element> StairsLandings_List = GetItemByIds(StairsLandingsId_List, StairsLandingsCollector);

                    stringBuilder.Append($"楼梯：{stairs.Id}");
                    foreach (var stairsroliings in StairsRoliings_List)
                    {
                        stringBuilder.Append($"，栏杆：{stairsroliings.Id}");
                    }
                    foreach (var stairsruns in StairsRuns_List)
                    {
                        stringBuilder.Append($"，梯段：{stairsruns.Id}");
                    }
                    foreach (var stairslandings in StairsLandings_List)
                    {
                        stringBuilder.Append($"，平面：{stairslandings.Id}");
                    }
                    
                    stringBuilder.AppendLine("");
                }
                catch (Exception e)
                {
                    //stringBuilder.Append("," + e.Message);
                }
            }
            */


            //==========获得墙体Walls的BoundingBox==========
            //根据每块墙的BoundingXYZ坐标判断
            //键：墙的id，值：坐标信息，例如：{Xmin:0,Xmax:1,Ymin:2,Ymax:3,...}
            Dictionary<Element, Dictionary<string, double>> WallBoundingXYZ_Dict = new Dictionary<Element, Dictionary<string, double>>();
            double WallMinX;
            double WallMaxX;
            double WallMinY;
            double WallMaxY;
            double WallMinZ;
            double WallMaxZ;
            int Wall_findNum = 0;
            int Wall_notfindNum = 0;

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
                    WallMinX = Math.Round(Get_Right_Number(min.X), 2);
                    WallMaxX = Math.Round(Get_Right_Number(max.X), 2);
                    WallMinY = Math.Round(Get_Right_Number(min.Y), 2);
                    WallMaxY = Math.Round(Get_Right_Number(max.Y), 2);
                    WallMinZ = Math.Round(Get_Right_Number(min.Z), 2);
                    WallMaxZ = Math.Round(Get_Right_Number(max.Z), 2);
                    WallBoundingXYZ_Dict[wall].Add("WallMinX", WallMinX);
                    WallBoundingXYZ_Dict[wall].Add("WallMaxX", WallMaxX);
                    WallBoundingXYZ_Dict[wall].Add("WallMinY", WallMinY);
                    WallBoundingXYZ_Dict[wall].Add("WallMaxY", WallMaxY);
                    WallBoundingXYZ_Dict[wall].Add("WallMinZ", WallMinZ);
                    WallBoundingXYZ_Dict[wall].Add("WallMaxZ", WallMaxZ);
                    //stringBuilder.AppendLine($"墙体名：{wall.Name}，{wall.Id}，X：({WallMinX},{WallMaxX})，Y：({WallMinY},{WallMaxY})，Z：({WallMinZ},{WallMaxZ})");
                    Wall_findNum++;

                } catch (Exception)
                {
                    Wall_notfindNum++;
                    //stringBuilder.AppendLine("找不到的墙体：" + wall.Id.ToString() + "，" + wall.Name.ToString());
                    WallBoundingXYZ_Dict.Remove(wall);
                }
            }
            //stringBuilder.AppendLine($"找得到的墙体数量：{Wall_findNum}，找不到的墙体数量：{Wall_notfindNum}");


            //==========获得栏杆Railing的BoundingBox==========
            Dictionary<Element, Dictionary<string, double>> RailingBoundingXYZ_Dict = new Dictionary<Element, Dictionary<string, double>>();
            double RailingMinX;
            double RailingMaxX;
            double RailingMinY;
            double RailingMaxY;
            double RailingMinZ;
            double RailingMaxZ;
            double RailingXLength;
            double RailingYLength;
            int find_Railing = 0;
            int notfind_Railing = 0;
            foreach (Element item_railing in StairsRailingCollector)
            {
                try
                {
                    Railing railing = item_railing as Railing;
                    if (!RailingBoundingXYZ_Dict.ContainsKey(railing))
                    {
                        RailingBoundingXYZ_Dict.Add(railing, new Dictionary<string, double>());
                    }
                    try
                    {
                        BoundingBoxXYZ boundingBoxXYZ = railing.get_BoundingBox(document.ActiveView);
                        XYZ max = boundingBoxXYZ.Max;
                        XYZ min = boundingBoxXYZ.Min;
                        RailingMinX = Math.Round(Get_Right_Number(min.X), 2);
                        RailingMaxX = Math.Round(Get_Right_Number(max.X), 2);
                        RailingMinY = Math.Round(Get_Right_Number(min.Y), 2);
                        RailingMaxY = Math.Round(Get_Right_Number(max.Y), 2);
                        RailingMinZ = Math.Round(Get_Right_Number(min.Z), 2);
                        RailingMaxZ = Math.Round(Get_Right_Number(max.Z), 2);
                        RailingXLength = Math.Round(Utils.FootToMeter(Math.Abs(RailingMinX - RailingMaxX)), 2);
                        RailingYLength = Math.Round(Utils.FootToMeter(Math.Abs(RailingMinY - RailingMaxY)), 2);
                        RailingBoundingXYZ_Dict[railing].Add("RailingMinX", RailingMinX);
                        RailingBoundingXYZ_Dict[railing].Add("RailingMaxX", RailingMaxX);
                        RailingBoundingXYZ_Dict[railing].Add("RailingMinY", RailingMinY);
                        RailingBoundingXYZ_Dict[railing].Add("RailingMaxY", RailingMaxY);
                        RailingBoundingXYZ_Dict[railing].Add("RailingMinZ", RailingMinZ);
                        RailingBoundingXYZ_Dict[railing].Add("RailingMaxZ", RailingMaxZ);
                        RailingBoundingXYZ_Dict[railing].Add("RailingXLength", RailingXLength);
                        RailingBoundingXYZ_Dict[railing].Add("RailingYLength", RailingYLength);
                        //stringBuilder.AppendLine($"栏杆名：{railing.Name}，{railing.Id}，X：({RailingMinX},{RailingMaxX})，Y：({RailingMinY},{RailingMaxY})，Z：({RailingMinZ},{RailingMaxZ})");
                        find_Railing++;
                    }
                    catch (Exception)
                    {
                        notfind_Railing++;
                        //stringBuilder.AppendLine("找不到的栏杆：" + railing.Id.ToString() + "，" + railing.Name.ToString());
                        RailingBoundingXYZ_Dict.Remove(railing);
                    }
                }catch { }
            }
            //stringBuilder.AppendLine($"找得到的栏杆数量：{find_Railing}，找不到的栏杆数量：{notfind_Railing}");


            

            //==========开始遍历==========
            double up_distance = 0;
            double bottom_distance = 0;
            double left_distance = 0;
            double right_distance = 0;

            double z_bia_up_distance = 0;    // 墙与栏杆的顶面差距距离
            double z_bia_bottom_distance = 0;// 墙与栏杆的底面差距距离



            foreach (KeyValuePair<Element, Dictionary<string, double>> item_stairsRoliings in RailingBoundingXYZ_Dict)
            {
                up_distance = 0;
                bottom_distance = 0;
                left_distance = 0;
                right_distance = 0;
                Railing railing = item_stairsRoliings.Key as Railing;
                RailingMinX = item_stairsRoliings.Value["RailingMinX"];
                RailingMaxX = item_stairsRoliings.Value["RailingMaxX"];
                RailingMinY = item_stairsRoliings.Value["RailingMinY"];
                RailingMaxY = item_stairsRoliings.Value["RailingMaxY"];
                RailingMinZ = item_stairsRoliings.Value["RailingMinZ"];
                RailingMaxZ = item_stairsRoliings.Value["RailingMaxZ"];
                RailingXLength = item_stairsRoliings.Value["RailingXLength"];
                RailingYLength = item_stairsRoliings.Value["RailingYLength"];
                foreach (KeyValuePair<Element, Dictionary<string, double>> item_wall in WallBoundingXYZ_Dict)
                {
                    Wall wall = item_wall.Key as Wall;
                    WallMinX = item_wall.Value["WallMinX"];
                    WallMaxX = item_wall.Value["WallMaxX"];
                    WallMinY = item_wall.Value["WallMinY"];
                    WallMaxY = item_wall.Value["WallMaxY"];
                    WallMinZ = item_wall.Value["WallMinZ"];
                    WallMaxZ = item_wall.Value["WallMaxZ"];
                    //判断墙在栏杆的哪个相对方向，以栏杆为中心

                    up_distance = Math.Round(Utils.FootToMeter(WallMinY - RailingMaxY), 2);      //北
                    bottom_distance = Math.Round(Utils.FootToMeter(RailingMinY - WallMaxY), 2);  //南
                    left_distance = Math.Round(Utils.FootToMeter(RailingMinX - WallMaxX), 2);    //西
                    right_distance = Math.Round(Utils.FootToMeter(WallMinX - RailingMaxX), 2);   //东

                    z_bia_up_distance = Math.Round(Utils.FootToMeter(Math.Abs(RailingMinZ - WallMinZ)), 2);
                    z_bia_bottom_distance = Math.Round(Utils.FootToMeter(Math.Abs(WallMaxZ - RailingMaxZ)), 2);

                    if (z_bia_up_distance <= 1 || z_bia_bottom_distance <= 1)
                    {
                        //墙在栏杆上方（北侧）
                        if (up_distance < 0.04 && up_distance >= 0 && WallMaxX >= RailingMinX && RailingMaxX >= WallMinX)
                        {
                            //stringBuilder_notpass.AppendLine($"{railing.Id}，{wall.Id}，{up_distance}，北，Z上：{z_bia_up_distance}，Z下：{z_bia_bottom_distance}");
                            stringBuilder_notpass.AppendLine($"{railing.Id}，{wall.Id}，{up_distance * 1000}mm，北");
                            H00101 = false;
                        }
                        //墙在栏杆下方（南侧）
                        else if (bottom_distance < 0.04 &&bottom_distance >= 0 && WallMaxX >= RailingMinX && RailingMaxX >= WallMinX)
                        {
                            //stringBuilder_notpass.AppendLine($"{railing.Id}，{wall.Id}，{bottom_distance}，南，Z上：{z_bia_up_distance}，Z下：{z_bia_bottom_distance}");
                            stringBuilder_notpass.AppendLine($"{railing.Id}，{wall.Id}，{bottom_distance * 1000}mm，南");
                            H00101 = false;
                        }
                        //墙在栏杆左侧（西侧）
                        else if (left_distance < 0.04 && left_distance >= 0 && RailingMaxY >= WallMinY && WallMaxY >= RailingMinY)
                        {
                            //stringBuilder_notpass.AppendLine($"{railing.Id}，{wall.Id}，{left_distance}，西，Z上：{z_bia_up_distance}，Z下：{z_bia_bottom_distance}");
                            stringBuilder_notpass.AppendLine($"{railing.Id}，{wall.Id}，{left_distance * 1000}mm，西");
                            H00101 = false;
                        }
                        //墙在栏杆右侧（东侧）
                        else if (right_distance < 0.04 && right_distance >= 0 && RailingMaxY >= WallMinY && WallMaxY >= RailingMinY)
                        {
                            //stringBuilder_notpass.AppendLine($"{railing.Id}，{wall.Id}，{right_distance}，东，Z上：{z_bia_up_distance}，Z下：{z_bia_bottom_distance}");
                            stringBuilder_notpass.AppendLine($"{railing.Id}，{wall.Id}，{right_distance * 1000}mm，东");
                            H00101 = false;
                        }
                        //位置无关的墙
                        //else
                        //{
                            //stringBuilder_notpass.AppendLine($"未知墙：{railing.Id}：{railing.Name}，{wall.Id}：{wall.Name}，上：{up_distance}，下：{bottom_distance}，左：{left_distance}，右：{right_distance}，Z上：{z_bia_up_distance}，Z下：{z_bia_bottom_distance}");
                        //}
                    }
                }
                /*
                if (RailingXLength > 0.50 && RailingYLength > 0.50)
                {
                    //stringBuilder.AppendLine($"栏杆：{railing.Id}，{railing.Name}，{RailingXLength}，{RailingYLength}，X：({RailingMinZ}，{RailingMaxX})，Y：({RailingMinY}，{RailingMaxY})，Z：({RailingMinZ}，{RailingMaxZ})");
                    stringBuilder_notpass.AppendLine($"==================================================================================================");
                    stringBuilder_notpass.AppendLine($"栏杆：{railing.Id}，{railing.Name}");
                    int not_pass_num = 0;
                    Wall temp_wall = null;
                    foreach (KeyValuePair<Element, Dictionary<string, double>> item_wall in WallBoundingXYZ_Dict)
                    {
                        Wall wall = item_wall.Key as Wall;
                        temp_wall = wall;
                        WallMinX = item_wall.Value["WallMinX"];
                        WallMaxX = item_wall.Value["WallMaxX"];
                        WallMinY = item_wall.Value["WallMinY"];
                        WallMaxY = item_wall.Value["WallMaxY"];
                        WallMinZ = item_wall.Value["WallMinZ"];
                        WallMaxZ = item_wall.Value["WallMaxZ"];
                        //判断墙在栏杆的哪个相对方向，以栏杆为中心
                        up_distance = Math.Round(Utils.FootToMeter(WallMaxY - RailingMinY), 2);
                        bottom_distance = Math.Round(Utils.FootToMeter(RailingMaxY - WallMinY), 2);
                        left_distance = Math.Round(Utils.FootToMeter(RailingMinX - WallMaxX), 2);
                        right_distance = Math.Round(Utils.FootToMeter(WallMinX - RailingMaxX), 2);
                        z_bia_up_distance = Math.Round(Utils.FootToMeter(Math.Abs(RailingMinZ - WallMinZ)), 2);
                        z_bia_bottom_distance = Math.Round(Utils.FootToMeter(Math.Abs(WallMaxZ - RailingMaxZ)), 2);
                        //栏杆距墙面需≥0.04米，＜0.04米则不符合
                        if (z_bia_up_distance < 1 && z_bia_bottom_distance < 1)
                        {
                            if ((up_distance >= 0 && up_distance < 0.04) || (bottom_distance > 0 && bottom_distance < 0.04) || (left_distance > 0 && left_distance < 0.04) || (right_distance > 0 && right_distance < 0.04))
                            {
                                not_pass_num++;
                                stringBuilder_notpass.AppendLine($"墙：{wall.Id}，{wall.Name}，上：{up_distance}，下：{bottom_distance}，左：{left_distance}，右：{right_distance}，Z上：{z_bia_up_distance}，Z下：{z_bia_bottom_distance}");
                            }

                        }
                    }
                }
                */
                //stringBuilder.AppendLine($"栏杆：{stairs.Key.Id}");
                //========== 遍历墙，判断每块墙的边缘与楼梯边缘的距离 ==========
                //bool exist_wall = false;

            }


            //stringBuilder.AppendLine("靠墙扶手边缘距墙面完成面净距不符合要求：\n" + stringBuilder_notpass);

            //stringBuilder.AppendLine("可能符合要求的栏杆：\n" + stringBuilder_might);
            stringBuilder.AppendLine("不符合要求的栏杆：\n" + stringBuilder_notpass);
            //stringBuilder.AppendLine("符合要求的栏杆：\n" + stringBuilder_pass);



            if (H00101 == true)
            {
                stringBuilder_result.AppendLine("符合5.3.3民用建筑通用规范 GB55031-2022");
            }
            else
            {
                stringBuilder_result.AppendLine("不符合5.3.3民用建筑通用规范 GB55031-2022");
            }
            stringBuilder_result.AppendLine("———————————————————————————————————————————————————————————");
            stringBuilder_result.AppendLine("输出格式说明：{栏杆Id}，{墙Id}，栏杆到墙面距离（单位：mm），{墙相对于栏杆的位置（eg.“北”代表墙在栏杆的北侧）}");
            stringBuilder_result.AppendLine("———————————————————————————————————————————————————————————");

            stringBuilder = stringBuilder_result.AppendLine(stringBuilder.ToString());

            //切换到三维平面
            switchto3D(uiDocument, document);


            Utils.PrintLog(stringBuilder.ToString(), "H00101", document);
            TaskDialog.Show("H00101强条检测", stringBuilder.ToString());
            return Result.Succeeded;
            
        }
    }
}
