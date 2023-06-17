using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ZBIM
{
    [Transaction(TransactionMode.Manual)]
    public class FilterArea47 : IExternalCommand
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











        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //界面交互的document
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            //实际内容的document
            Document document = commandData.Application.ActiveUIDocument.Document;
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
            //过滤得到所有面积分区Area，然后按照名字过滤，得到防火分区
            FilteredElementCollector AreaCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Areas);
            List<Element> areaLevels = new List<Element>();
            foreach (Area area in AreaCollector)
            {
                string up_levelname = null;
                int before = 0;
                int after = 0;
                if (area.AreaScheme.Name.Contains("防火分区"))
                {
                    areaLevels.Add(area.Level);
                    foreach (Level levelitem in areaLevels)//elementsLevels是所有楼层，我们需要按照buildFlag楼号来分类到对应的楼层
                    {
                        if (levelitem.Name.Contains("#-"))//处理楼层的名字
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
                        }
                        else
                        {
                            up_levelname = levelitem.Name;
                        }
                    }
                    //stringBuilder.AppendLine("原标高：" + area.Level.Name + "，处理后的标高：" + up_levelname);
                    string area_build;
                    string[] area_builds = Regex.Split(up_levelname, "、");
                    if (area.Level.Name.Contains("B"))
                    {
                        area_build = StartSubString(area.Level.Name, 0, "-");
                        FireArea_buildNum[area_build].Add(area);
                        stringBuilder.AppendLine("防火分区Id：" + area.Id + "，名称：" + area.Name + "，原标高：" + area.Level.Name + "，处理后的标高：" + up_levelname + "每栋楼号：" + area_build);
                    }
                    else 
                    {
                        stringBuilder.Append("防火分区Id：" + area.Id + "，名称：" + area.Name + "，原标高：" + area.Level.Name + "，处理后的标高：" + up_levelname + "，每栋楼号：");
                        foreach (string temp_area_build in area_builds)
                        {
                            if (temp_area_build.Contains("F"))
                            {
                                area_build = StartSubString(temp_area_build, 0, "#", true);
                            }
                            else
                            {
                                area_build = temp_area_build;
                            }
                            //去掉area_build里的“#”
                            area_build = StartSubString(area_build, 0, "#");
                            stringBuilder.Append(area_build + " ");
                            FireArea_buildNum[area_build].Add(area);
                        }
                        stringBuilder.AppendLine();
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
                    string area_firearea = area.LookupParameter("面积").AsValueString();// 得到area的面积，类似这种字段：“114514m²”
                    if (area.Name.Contains("B") || area.Name.Contains("Y") || area.Name.Contains("S"))
                    {
                        string area_level = LastSubEndString(area.Name, "-");
                        stringBuilder.AppendLine("第" + area_level + "个：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea);
                    }
                    else
                    {
                        string area_level = SubBetweenString(area.Name, "-", "F");
                        stringBuilder.AppendLine("第" + area_level + "层：防火分区Id：" + area.Id + "，名称：" + area.Name + "，面积：" + area_firearea);
                    }
                }
            }
            
            

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



            TaskDialog.Show("H00047强条检测", stringBuilder.ToString());

            return Result.Succeeded;
        }
    }
}
