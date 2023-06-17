using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ZBIMUtils;

namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    public class FilterArea47_hzx : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;  //界面交互的document
            Document document = commandData.Application.ActiveUIDocument.Document;  //实际内容的document
            StringBuilder stringBuilder = new StringBuilder();
            ElementId viewid = document.ActiveView.Id;

            //========存放每栋楼的防火分区Area=========
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
                { "B1", new List<Area>() },
                { "B2", new List<Area>() },
                { "B3", new List<Area>() },
                { "Y1", new List<Area>() },
                { "Y2", new List<Area>() },
                { "Y3", new List<Area>() },
                { "S1", new List<Area>() },
                { "S2", new List<Area>() },
                { "S3", new List<Area>() }
            };
            int area_num = 0;
            string area_buildnum;
            //过滤得到所有面积分区Area，然后按照名字过滤，得到防火分区
            FilteredElementCollector AreaCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Areas);
            foreach (Area area in AreaCollector)
            {
                if (area.AreaScheme.Name.Contains("防火分区"))
                {
                    if (area.Name.Contains("B"))
                    {
                        area_buildnum = Utils.LastSubEndString(area.Name, "防火分区");//先分离出防火分区后面的字段，例如“B1-1 1”
                        area_buildnum = Utils.StartSubString(area_buildnum, 0, "-");//然后获得楼号，例如“B1”
                        FireArea_buildNum[area_buildnum].Add(area);
                        stringBuilder.AppendLine("防火分区Id：" + area.Id + "，名称：" + area.Name + "，楼号：" + area_buildnum + "#");
                        area_num ++;
                    }
                    else 
                    {
                        area_buildnum = Utils.LastSubEndString(area.Name, "防火分区");//先分离出防火分区后面的字段，例如“1#-2F-1 10”
                        area_buildnum = Utils.StartSubString(area_buildnum, 0, "#-");//然后获得楼号，例如“1”
                        FireArea_buildNum[area_buildnum].Add(area);
                        stringBuilder.AppendLine("防火分区Id：" + area.Id + "，名称：" + area.Name +"，楼号：" + area_buildnum + "#");
                        area_num ++;
                    }
                    
                }
            }

            
            stringBuilder.AppendLine("一共有" + area_num +"个防火分区\n\n");
            stringBuilder.AppendLine("查看每栋楼每层的防火分区Area：");
            foreach (var item in FireArea_buildNum)
            {
                stringBuilder.AppendLine($"楼号：{item.Key}#");
                foreach (Area area in item.Value)
                {
                    string area_firearea = area.LookupParameter("面积").AsValueString();  // 得到area的面积，类似这种字段：“114514m²”
                    if (area.Name.Contains("B") || area.Name.Contains("Y") || area.Name.Contains("S"))
                    {
                        string area_level = Utils.LastSubEndString(area.Name, "-");
                        stringBuilder.AppendLine("第" + area_level + "个：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea);
                    }
                    else
                    {
                        string area_level = Utils.SubBetweenString(area.Name, "-", "F");
                        stringBuilder.AppendLine("第" + area_level + "层：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea);
                    }
                }
            }
            
            
            /*
            //========存放每栋楼的防火分区Tag=========
            Dictionary<string, List<SpatialElementTag>> FireTag_buildNum = new Dictionary<string, List<SpatialElementTag>>
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
            int tag_num = 0;
            //过滤得到所有AreaTag，area.Name==Tag.TagText后半部分
            FilteredElementCollector SpatialElementTagCollector = new FilteredElementCollector(document).OfClass(typeof(Autodesk.Revit.DB.SpatialElementTag));
            foreach (SpatialElementTag Tag in SpatialElementTagCollector)
            {
                if (Tag.TagText.Contains("防火"))
                {
                    tag_num++;
                    string view_name = Tag.View.Name; //得到Tag所在标高，类似这种字段：“B2-9.300（建）”
                    string tag_firearea = StartSubString(Tag.TagText, 0, "m²", true);// 得到Tag的面积，类似这种字段：“114514m²”
                    //stringBuilder.Append(tag_num + "、标签名Id：" + Tag.Id.ToString() + "，面积：" + tag_firearea + "，标签TagText：" + Tag.TagText + "，标签Name：" + Tag.Name + "，所属视图：" + view_name);
                    stringBuilder.Append(tag_num + "、标签Id：" + Tag.Id.ToString()  + "，标签TagText：" + Tag.TagText + "，面积：" + tag_firearea + "，所属视图：" + view_name);

                    if (Tag.TagText.Contains("B") || Tag.TagText.Contains("Y") || Tag.TagText.Contains("S"))
                    {
                        string tag_level = SubBetweenString(Tag.TagText, "防火分区", "-");// 类似字段：“B1”，地下的Tag只有层数信息，没有楼号
                        stringBuilder.AppendLine("，层号：" + tag_level);
                        FireTag_buildNum[tag_level].Add(Tag);
                    }
                    else
                    {
                        string temp_tag_level = SubBetweenString(Tag.TagText, "防火分区", "F-");// 类似字段：“5#-3”
                        string tag_level = LastSubEndString(temp_tag_level, "-"); //进一步，提取层号，类似字段：“3”
                        string tag_build_num = StartSubString(temp_tag_level, 0, "#"); // 提取楼号，类似字段：“5”
                        stringBuilder.AppendLine("，楼号：" + tag_build_num + "#" + "，层号：" + tag_level + "F");
                        FireTag_buildNum[tag_build_num].Add(Tag);
                    }
                }

            }
            stringBuilder.AppendLine("一共有" + tag_num + "个防火分区标签\n\n");


            stringBuilder.AppendLine("查看每栋楼每层的防火分区Tag：");
            foreach (var item in FireTag_buildNum)
            {
                stringBuilder.AppendLine($"楼号：{item.Key}#");
                foreach (SpatialElementTag Tag in item.Value)
                {
                    string view_name = Tag.View.Name; //得到Tag所在标高，类似这种字段：“B2-9.300（建）”
                    string tag_firearea = StartSubString(Tag.TagText, 0, "m²", true);// 得到Tag的面积，类似这种字段：“114514m²”
                    if (Tag.TagText.Contains("B") || Tag.TagText.Contains("Y") || Tag.TagText.Contains("S"))
                    {
                        string tag_level = LastSubEndString(Tag.TagText, "-");
                        stringBuilder.AppendLine("第" + tag_level + "个：标签名Id：" + Tag.Id.ToString() + "，面积：" + tag_firearea + "，标签TagText：" + Tag.TagText + "，所属视图：" + view_name);
                    }
                    else
                    {
                        string tag_level = SubBetweenString(Tag.TagText, "-", "F");
                        stringBuilder.AppendLine("第" + tag_level + "层：标签名Id：" + Tag.Id.ToString() + "，面积：" + tag_firearea + "，标签TagText：" + Tag.TagText + "，所属视图：" + view_name);
                    }
                }
            }
            */


            TaskDialog.Show("H00047强条检测", stringBuilder.ToString());

            return Result.Succeeded;
        }
    }
}
