using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ZBIMUtils;


namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    public class H00043 : IExternalCommand
    {
        // 获取房间内的所有门，返回一个包含房间内所有门的列表
        public static List<FamilyInstance> GetDoorsInRoom(Room room, Document doc)
        {
            List<FamilyInstance> doors = new List<FamilyInstance>();
            IList<IList<BoundarySegment>> loops = room.GetBoundarySegments(new SpatialElementBoundaryOptions());
            foreach (IList<BoundarySegment> loop in loops)
            {
                foreach (BoundarySegment segment in loop)
                {
                    Wall wall = doc.GetElement(segment.ElementId) as Wall;
                    if (wall != null)
                    {
                        // 处理内嵌墙：
                        var embededWallIds = wall.FindInserts(false, false, true, false);
                        foreach (var embededWallId in embededWallIds)
                        {
                            Wall embededWall = doc.GetElement(embededWallId) as Wall;
                            if (embededWall != null && embededWall.CurtainGrid != null)
                            {
                                ICollection<ElementId> panelIds
                                    = embededWall.CurtainGrid.GetPanelIds();
                                foreach (var panelId in panelIds)
                                {
                                    var door = doc.GetElement(panelId) as FamilyInstance;
                                    if (door != null
                                        && door.Category.Id.IntegerValue
                                        == (int)BuiltInCategory.OST_Doors)
                                        doors.Add(door);
                                }
                            }
                        }
                        // 处理孔洞：
                        var openingIds = wall.FindInserts(true, false, false, false);
                        foreach (var openingId in openingIds)
                        {
                            var door = doc.GetElement(openingId) as FamilyInstance;
                            if (door != null
                                && door.Category.Id.IntegerValue
                                == (int)BuiltInCategory.OST_Doors)
                                doors.Add(door);
                        }
                    }
                }
            }
            // 后处理：删除重复项，以及不属于房间的元素。只保留带“FM”的门
            List<FamilyInstance> resoluts = new List<FamilyInstance>();
            foreach (var item in doors)
            {
                if (resoluts.Exists(x => x.Id.Equals(item.Id)))
                {
                    continue;
                }


                if (item.ToRoom != null && item.ToRoom.Id.Equals(room.Id))
                {
                    resoluts.Add(item);
                }
                else if (item.FromRoom != null && item.FromRoom.Id.Equals(room.Id))
                {
                    resoluts.Add(item);
                }
                /*
                if (item.Name.Contains("FM"))
                {
                    
                }
                */

            }
            return resoluts;
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
            ElementId viewid = document.ActiveView.Id;
            StringBuilder stringBuilder = new StringBuilder();
            Boolean H00043_pass = true;
            StringBuilder H00043_result = new StringBuilder();


            //切换到二维平面
            switchto2D(uiDocument, document);




            //////////////////////获得地下楼层的Level的数组////////////////////////////
            //获得所有“楼层平面”Level
            FilteredElementCollector collectorLevels = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Levels);
            //获得所有地下的“楼层平面”Level
            List<Element> level_list = new List<Element>();
            //过滤不需要的“楼层平面”Level
            foreach (var level in collectorLevels)
            {
                if (level.Name.Contains("B") && !level.Name.Contains("(S") && !level.Name.Contains("结") && !level.Name.Contains("Y"))
                {
                    //stringBuilder.AppendLine(level.Name);
                    level_list.Add(level);
                }
            }
            //对“楼层平面”Level进行排序
            for (int i = level_list.Count - 1; i > 0; i--)
            {
                for (int j = 0; j < i; j++)
                {
                    int level_num = Convert.ToInt16(level_list[j].Name[1]);
                    int next_level_num = Convert.ToInt16(level_list[j + 1].Name[1]);
                    if (level_num > next_level_num)
                    {
                        Element level = level_list[j];
                        level_list[j] = level_list[j + 1];
                        level_list[j + 1] = level;
                    }
                }
            }
            /*
            //查看“楼层平面”Level的名字
            foreach (var level in level_list)
            {
                stringBuilder.AppendLine(level.Name);
            }
            */
            //////////////////////获得地下楼层的Level的数组////////////////////////////



            //////////////////////获得地下楼层的ViewPlan数组////////////////////////////
            FilteredElementCollector collectorViewPlan = new FilteredElementCollector(document).OfClass(typeof(ViewPlan));//过滤所有的2D视图
            IList<Element> viewPlans = collectorViewPlan.OfClass(typeof(ViewPlan)).ToElements();
            List<ViewPlan> viewPlan_list = new List<ViewPlan>();
            //过滤出所有地面以下的ViewPlan
            foreach (ViewPlan viewPlan in collectorViewPlan)
            {
                if (viewPlan.Name.Contains("B") &&
                    !viewPlan.Name.Contains("(S") &&
                    !viewPlan.Name.Contains("结") &&
                    !viewPlan.Name.Contains("Y")
                   )
                {
                    //stringBuilder.AppendLine(viewPlan.Name);
                    viewPlan_list.Add(viewPlan);
                }
            }
            //对ViewPlan进行排序
            for (int i = viewPlan_list.Count - 1; i > 0; i--)
            {
                for (int j = 0; j < i; j++)
                {
                    int viewPlan_num = Convert.ToInt16(viewPlan_list[j].Name[1]);
                    int next_viewPlan_num = Convert.ToInt16(viewPlan_list[j + 1].Name[1]);
                    if (viewPlan_num > next_viewPlan_num)
                    {
                        ViewPlan viewPlan = viewPlan_list[j];
                        viewPlan_list[j] = viewPlan_list[j + 1];
                        viewPlan_list[j + 1] = viewPlan;
                    }
                }
            }
            //去除重复的ViewPlan，得到viewPlan_list2
            List<ViewPlan> viewPlan_list2 = new List<ViewPlan>();
            try {
                viewPlan_list2.Add(viewPlan_list[0]);
            }
            catch(Exception ex) { }
            for (int i = 1; i < viewPlan_list.Count; i++)
            {
                if (viewPlan_list[i - 1].Name != viewPlan_list[i].Name)
                {
                    viewPlan_list2.Add(viewPlan_list[i]);
                }
            }
            /*
            //查看该模型中地下的楼层平面
            foreach (ViewPlan viewPlan in viewPlan_list2)
            {
                stringBuilder.AppendLine(viewPlan.Name);
            }
            */
            //////////////////////获得地下楼层的ViewPlan数组////////////////////////////



            //////////////////////匹配Level和ViewPlan，对每个楼层的房间进行过滤，并进行检测////////////////////////////

            StringBuilder stringBuilder_pass_floor = new StringBuilder();
            StringBuilder stringBuilder_notpass_floor = new StringBuilder();
            StringBuilder stringBuilder_pass_elevator = new StringBuilder();
            StringBuilder stringBuilder_notpass_elevator = new StringBuilder();

            StringBuilder stringBuilder_notpass_floor_door = new StringBuilder();
            StringBuilder stringBuilder_notpass_elevator_door = new StringBuilder();

            Dictionary<Element, List<Element>> notpass_floor_door_dict = new Dictionary<Element, List<Element>>();
            Dictionary<Element, List<Element>> notpass_elevator_door_dict = new Dictionary<Element, List<Element>>();

            List<ElementId> pass_floor_list = new List<ElementId>();
            List<ElementId> notpass_floor_list = new List<ElementId>();
            List<ElementId> pass_elevator_list = new List<ElementId>();
            List<ElementId> notpass_elevator_list = new List<ElementId>();

            List<Element> notpass_floor_door_list = new List<Element>();
            List<Element> notpass_elevator_door_list = new List<Element>();

            //存储所有不需要的门的Id（“XX井”）
            List<Element> other_door_list = new List<Element>();
            List<ElementId> other_doorId_list = new List<ElementId>();

            int notpass_floor_num = 0;
            int notpass_elevator_num = 0;

            //存储符合的房间
            List<Element> room_floor_list = new List<Element>();
            List<Element> room_elevator_list = new List<Element>();
            //存储符合的房间的房门
            List<FamilyInstance> room_doors_list = new List<FamilyInstance>();
            //double roomMinX; //房间最小的X坐标
            //double roomMinY; //房间最小的Y坐标
            //double roomMaxX; //房间最大的X坐标
            //double roomMaxY; //房间最大的Y坐标
            double roomMaxZ; //房间最大的Y坐标
            double roomMinZ; //房间最大的Y坐标


            //stringBuilder.Append("====================所有“XX井”的房门信息====================");
            //获取所有“XX井”房间的房门，将房门Id存到List数组里
            foreach (var litem in level_list)
            {
                foreach (ViewPlan viewPlan in viewPlan_list2)
                {
                    if (viewPlan.Name.Equals(litem.Name))
                    {
                        //stringBuilder.AppendLine("\n" + viewPlan.Name + "：");
                        //过滤当前ViewPlan的房间
                        FilteredElementCollector collectorRooms = new FilteredElementCollector(document, viewPlan.Id);
                        List<Element> elementsRooms = collectorRooms.OfCategory(BuiltInCategory.OST_Rooms).ToList<Element>();
                        //遍历当前ViewPlan的房间
                        foreach (Element room in elementsRooms)
                        {
                            room_doors_list = new List<FamilyInstance>();
                            BoundingBoxXYZ roomBoundingBoxXYZ = room.get_BoundingBox(document.ActiveView);//计算轴网建筑的占地范围
                            XYZ roomXYZMax = roomBoundingBoxXYZ.Max; //取到Max元组
                            XYZ roomXYZMin = roomBoundingBoxXYZ.Min; //取到Min元组
                            roomMaxZ = roomXYZMax.Z;
                            roomMinZ = roomXYZMin.Z;
                            //设置过滤房间的条件
                            if (roomMaxZ < 0 || roomMinZ < 0)
                            {
                                if (room.Name.Contains("井"))
                                {
                                    room_doors_list = Utils.GetDoorsInRoom((Room)room, document);//获得某个房间房门列表
                                    if (room_doors_list.Count < 1)//如果房门数量小于1，就认为该房间没有房门
                                    {
                                        //stringBuilder.AppendLine("（无房门）");
                                    }
                                    else//如果房门数量大于等于1，就认为该房间有房门
                                    {
                                        other_door_list.Add(room);//将含有“井”且有房门的房间加入数组中
                                        //stringBuilder.AppendLine(room.Name + "，" + room.Id + "，防火门数量：" + room_doors_list.Count);
                                        foreach (var item in room_doors_list)
                                        {
                                            try
                                            {
                                                if (!other_doorId_list.Contains(item.Id))
                                                {
                                                    other_doorId_list.Add(item.Id);//将所有“XX井”的房门Id存到数组中，以备后面遍历去除这些门
                                                }
                                                string fire_rating = item.Symbol.LookupParameter("防火等级").AsString();//获取每扇房门的防火等级
                                                //stringBuilder.AppendLine("      " + item.Name + "，" + item.Id + "，" + fire_rating);
                                            }
                                            catch (Exception e)
                                            {
                                                //stringBuilder.AppendLine($"{e.Message}");
                                                //stringBuilder.AppendLine("      " + item.Name + "，" + item.Id + "，无防火等级");
                                            }

                                        }
                                    }
                                }
                            }
                            else//标高大于0.00的房间，理论上是不存在的，因为含“B”楼层的房间都标高都小于0.00，所以不会执行else里的语句
                            {
                                //stringBuilder.AppendLine("发现标高大于0.00的房间，roomId：" + room.Id + "，room.Name：" + room.Name + "，Zmax：" + roomMaxZ + "，Zmin：" + roomMinZ);
                            }
                        }
                        //当前ViewPlan视图的房间遍历结束

                    }
                }
            }


            //stringBuilder.Append("====================所有“楼梯间”和“电梯间”的房门信息====================");
            //获取地下每一层特定房间名的房门，删掉不需要门，检测这些房门的防火等级
            foreach (var litem in level_list)
            {
                foreach (ViewPlan viewPlan in viewPlan_list2)
                {
                    if (viewPlan.Name.Equals(litem.Name))
                    {
                        //stringBuilder.AppendLine("\n" + viewPlan.Name + "：");
                        //过滤当前ViewPlan的房间
                        FilteredElementCollector collectorRooms = new FilteredElementCollector(document, viewPlan.Id);
                        List<Element> elementsRooms = collectorRooms.OfCategory(BuiltInCategory.OST_Rooms).ToList<Element>();
                        //遍历当前ViewPlan的房间
                        foreach (Element room in elementsRooms)
                        {
                            BoundingBoxXYZ roomBoundingBoxXYZ = room.get_BoundingBox(document.ActiveView);//计算轴网建筑的占地范围
                            XYZ roomXYZMax = roomBoundingBoxXYZ.Max; //取到Max元组
                            XYZ roomXYZMin = roomBoundingBoxXYZ.Min; //取到Min元组
                            roomMaxZ = roomXYZMax.Z;
                            roomMinZ = roomXYZMin.Z;
                            //设置过滤房间的条件
                            if (roomMaxZ < 0 || roomMinZ < 0)
                            {
                                room_doors_list = new List<FamilyInstance>();
                                if (room.Name.Contains("楼梯间"))
                                {
                                    room_doors_list = GetDoorsInRoom((Room)room, document);//获得某个房间房门列表
                                    if (room_doors_list.Count < 1)//如果房门数量小于1，就认为该房间没有房门
                                    {
                                        //stringBuilder.AppendLine("（无房门）");
                                    }
                                    else//如果房间有房门，就判断所有房门的防火等级
                                    {
                                        int notpass_flag = 0;
                                        notpass_floor_door_list = new List<Element>();
                                        //stringBuilder.AppendLine(room.Name + "，" + room.Id + "，防火门数量：" + room_doors_list.Count);
                                        //去除和和“XX井”相连的门，如果room_doors_list中存在和other_door_list一致的门，就删除这扇门
                                        foreach (var door in room_doors_list)
                                        {
                                            foreach (var other_door in other_door_list)
                                            {
                                                if (door.Id == other_door.Id)
                                                {
                                                    room_doors_list.Remove(door);
                                                }
                                            }
                                        }
                                        //检测该房间每扇门的防火等级
                                        foreach (var item in room_doors_list)
                                        {
                                            try
                                            {
                                                string fire_rating = item.Symbol.LookupParameter("防火等级").AsString();//获取每扇房门的防火等级
                                                if (fire_rating.Contains("甲") || fire_rating.Contains("乙"))
                                                {
                                                    if (!pass_floor_list.Contains(room.Id))
                                                    {
                                                        pass_floor_list.Add(room.Id);
                                                    }
                                                    //stringBuilder.AppendLine("      " + item.Name + "，" + item.Id + "，" + fire_rating);
                                                }
                                                else
                                                {
                                                    H00043_pass = false;
                                                    if (!notpass_floor_list.Contains(room.Id))
                                                    {
                                                        notpass_floor_list.Add(room.Id);
                                                        notpass_floor_num++;
                                                    }
                                                    notpass_flag = 1;
                                                    notpass_floor_door_list.Add(item);//将所有该房间不符合的门的Id存到数组中，数组存到键为该房间Id的字典中
                                                    //stringBuilder.AppendLine("      " + item.Name + "，" + item.Id + "，" + fire_rating + "（不符合要求）");
                                                }
                                            }
                                            catch (Exception e)//catch中的报的异常是没有“防火等级”属性的门
                                            {
                                                //stringBuilder.AppendLine($"{e.Message}");
                                                //stringBuilder.AppendLine("      " + item.Name + "，" + item.Id + "，无防火等级");
                                            }

                                        }
                                        if (notpass_flag == 1)
                                        {
                                            notpass_floor_door_dict.Add(room, notpass_floor_door_list);
                                        }

                                    }
                                }
                                else if (
                                    room.Name.Contains("电梯厅") ||
                                    room.Name.Contains("消防电梯前室") ||
                                    room.Name.Contains("合用前室") ||
                                    room.Name.Contains("三合一前室")
                                    )
                                {
                                    room_doors_list = GetDoorsInRoom((Room)room, document);//获得某个房间房门列表
                                    if (room_doors_list.Count < 1)//如果房门数量小于1，就认为该房间没有房门
                                    {
                                        //stringBuilder.AppendLine("（无房门）");
                                    }
                                    else//如果房间有房门，就判断所有房门的防火等级
                                    {
                                        int notpass_flag = 0;
                                        notpass_elevator_door_list = new List<Element>();
                                        //stringBuilder.AppendLine(room.Name + "，" + room.Id + "，防火门数量：" + room_doors_list.Count);
                                        //去除和和“XX井”相连的门，如果room_doors_list中存在和other_door_list一致的门，就删除这扇门
                                        foreach (var door in room_doors_list)
                                        {
                                            foreach (var other_door in other_door_list)
                                            {
                                                if (door.Id == other_door.Id)
                                                {
                                                    room_doors_list.Remove(door);
                                                }
                                            }
                                        }
                                        //检测每扇门的防火等级
                                        foreach (var item in room_doors_list)
                                        {
                                            try
                                            {
                                                string fire_rating = item.Symbol.LookupParameter("防火等级").AsString();//获取每扇房门的防火等级
                                                if (fire_rating.Contains("甲") || fire_rating.Contains("乙"))
                                                {
                                                    if (!pass_elevator_list.Contains(room.Id))
                                                    {
                                                        pass_elevator_list.Add(room.Id);
                                                    }
                                                    //stringBuilder.AppendLine("      " + item.Name + "，" + item.Id + "，" + fire_rating);
                                                }
                                                else
                                                {
                                                    H00043_pass = false;
                                                    notpass_flag = 1;
                                                    if (!notpass_elevator_list.Contains(room.Id))
                                                    {
                                                        notpass_elevator_list.Add(room.Id);
                                                        notpass_elevator_num++;
                                                    }

                                                    notpass_elevator_door_list.Add(item);//将所有该房间不符合的门的Id存到数组中，数组存到键为该房间Id的字典中
                                                    //stringBuilder.AppendLine("      " + item.Name + "，" + item.Id + "，" + fire_rating + "（不符合要求）");
                                                }
                                            }
                                            catch (Exception e)//catch中的报的异常是没有“防火等级”属性的门
                                            {
                                                //stringBuilder.AppendLine($"{e.Message}");
                                                //stringBuilder.AppendLine("      " + item.Name + "，" + item.Id + "，无防火等级");
                                            }
                                        }
                                        if (notpass_flag == 1)
                                        {
                                            notpass_elevator_door_dict.Add(room, notpass_elevator_door_list);
                                        }
                                    }
                                }
                                else//其他名字的房间，不需要判断
                                {
                                    //stringBuilder.AppendLine("不需要的房间：" + room.Name);
                                }
                            }
                            else//标高大于0.00的房间，理论上是不存在的，因为含“B”楼层的房间都标高都小于0.00，所以不会执行else里的语句
                            {
                                //stringBuilder.AppendLine("发现标高大于0.00的房间，roomId：" + room.Id + "，room.Name：" + room.Name + "，Zmax：" + roomMaxZ + "，Zmin：" + roomMinZ);
                            }
                        }
                        //当前ViewPlan视图的房间遍历结束

                    }
                }
            }


            //stringBuilder.AppendLine("不符合的楼梯间个数：" + notpass_floor_num);
            //stringBuilder.AppendLine("不符合的电梯间个数：" + notpass_elevator_num);
            //stringBuilder.AppendLine("=================以上为调试信息=================\n\n");

            /*
            foreach (ElementId roomId in notpass_floor_list)
            {
                stringBuilder_notpass_floor.AppendLine(roomId.ToString());
            }
            foreach (ElementId roomId in notpass_elevator_list)
            {
                stringBuilder_notpass_elevator.AppendLine(roomId.ToString());
            }
            */
            foreach (KeyValuePair<Element, List<Element>> entry in notpass_floor_door_dict)
            {
                stringBuilder_notpass_floor.Append(entry.Key.Id.ToString() + "：");
                foreach (var floor_door in entry.Value)
                {
                    stringBuilder_notpass_floor.Append(floor_door.Id.ToString() + "  ");
                }
                stringBuilder_notpass_floor.AppendLine("");
            }
            foreach (KeyValuePair<Element, List<Element>> entry in notpass_elevator_door_dict)
            {
                stringBuilder_notpass_elevator.Append(entry.Key.Id.ToString() + "：");
                foreach (var elevator_door in entry.Value)
                {
                    stringBuilder_notpass_elevator.Append(elevator_door.Id.ToString() + "  ");
                }
                stringBuilder_notpass_elevator.AppendLine("");
            }

            stringBuilder.AppendLine("不符合的楼梯间（房间Id：房门Id）：\n" + stringBuilder_notpass_floor);
            stringBuilder.AppendLine("不符合的电梯间（房间Id：房门Id）：\n" + stringBuilder_notpass_elevator);


            //////////////////////匹配Level和ViewPlan，对每个楼层的房间进行过滤，并进行检测////////////////////////////


            if (H00043_pass == false)
            {
                H00043_result.AppendLine("不符合6.9.6住宅设计规范GB50096-2011\n");
            }
            else
            {
                H00043_result.AppendLine("符合6.9.6住宅设计规范GB50096-2011\n");
            }
            //stringBuilder = H00043_result.AppendLine("房间数量：" + room_num + "\n\n" + stringBuilder.ToString());
            stringBuilder = H00043_result.AppendLine(stringBuilder.ToString());
            
            //切换到三维平面
            switchto3D(uiDocument, document);
            //打印输出
            Utils.PrintLog(stringBuilder.ToString(), "H00043", document);

            TaskDialog.Show("H00043强条检测", stringBuilder.ToString());

            return Result.Succeeded;
        }



    }
}
