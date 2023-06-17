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
using ZBIMUtils;


namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    public class MyUtil : IExternalCommand
    {

        //获得每栋建筑中某种类型的面积分区
        public static Dictionary<string, List<Area>> GetBuildArea(Autodesk.Revit.DB.Document document, string AreaTypeName)
        {
            Dictionary<string, List<Area>> FireArea_buildNum = new Dictionary<string, List<Area>>
            {
                { "1", new List<Area>() },
                { "2", new List<Area>() },
                { "3", new List<Area>() },
                { "4", new List<Area>() },
                { "5", new List<Area>() },
                { "6", new List<Area>() },
                { "7", new List<Area>() },
                { "8", new List<Area>() },
                { "9", new List<Area>() },
                { "10", new List<Area>() },
                { "11", new List<Area>() },
                { "12", new List<Area>() },
                { "13", new List<Area>() },
                { "14", new List<Area>() },
                { "15", new List<Area>() },
                { "16", new List<Area>() },
                { "17", new List<Area>() },
                { "18", new List<Area>() },
                { "19", new List<Area>() },
                { "20", new List<Area>() },
                { "21", new List<Area>() },
                { "22", new List<Area>() },
                { "23", new List<Area>() },
                { "24", new List<Area>() },
                { "25", new List<Area>() },
                { "26", new List<Area>() },
                { "27", new List<Area>() },
                { "28", new List<Area>() },
                { "29", new List<Area>() },
                { "30", new List<Area>() },
                { "31", new List<Area>() },
                { "32", new List<Area>() },
                { "33", new List<Area>() },
                { "34", new List<Area>() },
                { "35", new List<Area>() },
                { "B1", new List<Area>() },
                { "B2", new List<Area>() },
                { "B3", new List<Area>() },
                { "B4", new List<Area>() },
                { "B5", new List<Area>() },
                { "B6", new List<Area>() },
                { "Y1", new List<Area>() },
                { "Y2", new List<Area>() },
                { "Y3", new List<Area>() },
                { "Y4", new List<Area>() },
                { "Y5", new List<Area>() },
                { "Y6", new List<Area>() },
                { "Y7", new List<Area>() },
                { "Y8", new List<Area>() },
                { "Y9", new List<Area>() },
                { "S1", new List<Area>() },
                { "S2", new List<Area>() },
                { "S3", new List<Area>() },
                { "S4", new List<Area>() },
                { "S5", new List<Area>() },
                { "S6", new List<Area>() },
                { "S7", new List<Area>() },
                { "S8", new List<Area>() },
                { "S9", new List<Area>() }
            };


            
            

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




        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            Document document = commandData.Application.ActiveUIDocument.Document;


            Utils.GetBasementHeight(document);





            StringBuilder stringBuilder = new StringBuilder();
            TaskDialog.Show("H00018强条检测", stringBuilder.ToString());
            return Result.Succeeded;
        }
    }
}
