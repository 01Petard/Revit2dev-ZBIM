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
    public class FilterViewPlan47 : IExternalCommand
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


            List<ViewPlan> FireAreaViewPlan = new List<ViewPlan>();
            stringBuilder.AppendLine(document.ActiveView.GetType().Equals(typeof(ViewPlan)).ToString());

            //ViewPlan viewPlan2D = null;
            if (document.ActiveView.GetType().Equals(typeof(ViewPlan)))
            {
                FilteredElementCollector ViewPlanCollector = new FilteredElementCollector(document).OfClass(typeof(ViewPlan));
                foreach (var viewPlan in ViewPlanCollector)
                {
                    string viewPlan_category = viewPlan.LookupParameter("类型").AsValueString();
                    try
                    {
                        if (viewPlan_category.Contains("防火分区") && !viewPlan.Name.Contains("B"))
                        {
                            //stringBuilder.AppendLine("视图名：" + viewPlan.Name + "，视图ID：" + viewPlan.Id + "，视图类型：" + viewPlan_category);
                            //viewPlan2D = viewPlan;
                            FireAreaViewPlan.Add((ViewPlan)viewPlan);
                        }
                    }catch(Exception e){
                        //有些视图类型不是通过AsValueString属性获得的，因此这一步需要处理异常
                        //stringBuilder.AppendLine(e.Message);
                    }
                    
                }
            }

            stringBuilder.AppendLine("查看防火分区的ViewPlan");
            foreach (var view in FireAreaViewPlan)
            {
                stringBuilder.AppendLine("视图名：" + view.Name + "，视图ID：" + view.Id + "，视图类型：" + view.LookupParameter("类型").AsValueString());
            }

            TaskDialog.Show("FilterViewPlan", stringBuilder.ToString());

            return Result.Succeeded;
        }
    }
}
