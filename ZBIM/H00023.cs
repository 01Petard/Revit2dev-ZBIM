using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;

namespace H00008
{
    [Transaction(TransactionMode.Manual)]
    class H00023: IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) {
            //界面交互的doc
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            //实际内容的doc
            Document doc = commandData.Application.ActiveUIDocument.Document;

            //1、创建收集器
            FilteredElementCollector FireAlarmcollector = new FilteredElementCollector(doc);
            FilteredElementCollector FireAlarmcollector_levelid = new FilteredElementCollector(doc);
            //2、快速过滤方法
            FireAlarmcollector.OfCategory(BuiltInCategory.OST_FireAlarmDevices).OfClass(typeof(FamilyInstance));
            ICollection<ElementId> levelIds = FireAlarmcollector_levelid.OfClass(typeof(Level)).ToElementIds();
            

            string txt = "";
            string notpass = "";
            int notpass_count = 0;
            foreach (ElementId levelid in levelIds) {                
                Element level_ele = doc.GetElement(levelid);
                string temptxt = "";
                //txt += "********************";
                //txt += "楼层名："+level_ele.Name+",楼层ID："+levelid.ToString()+"\n";
                bool flag = FindSpecificLevelFireAlarm(doc, FireAlarmcollector, levelid, ref temptxt, ref notpass);
                if (flag == true)
                {//该楼层存在消火栓
                    txt += "********************\n";
                    txt += "楼层名：" + level_ele.Name + ",楼层ID：" + levelid.ToString() + "\n";
                    txt += "该楼层包含的消火栓箱信息如下：\n";
                    txt += temptxt;
                }
                else
                {//该楼层不存在消火栓
                    notpass_count++;
                }

                //FireAlarmcollector_levelid = new FilteredElementCollector(doc);
                //ElementLevelFilter filter = new ElementLevelFilter(levelid);
                //ICollection<ElementId> ele_founds = FireAlarmcollector_levelid.WherePasses(filter).ToElementIds();
            }

            if (notpass_count != 0)
            {//存在不合格的楼层
                string pass = "";
                if (levelIds.Count != notpass_count)
                    pass += "##################################################\n存在" + (levelIds.Count - notpass_count).ToString() + "个合格设置消火栓楼层\n" + "楼层信息如下：\n" + txt;
                else
                    pass += "##################################################\n存在0个合格设置消火栓楼层,当前建筑无消火栓箱！\n";
                TaskDialog.Show("H00023_消防给水及消火栓系统技术规范GB50974-2014", "不合格！存在" + notpass_count.ToString() + "个未设置消火栓楼层\n" + "楼层信息如下：\n" + notpass + pass);

            }
            else {
                TaskDialog.Show("H00023_消防给水及消火栓系统技术规范GB50974-2014", "合格！\n合格楼层信息如下：\n"  + txt);
            }
            return Result.Succeeded;
        }

        public bool FindSpecificLevelFireAlarm(Document doc, FilteredElementCollector FireAlarmcollector, ElementId levelid, ref string temptxt, ref string notpass) {
            bool flag = false;
            foreach (var item in FireAlarmcollector) {
                if (item.LevelId == levelid) {
                    flag = true;
                    temptxt += item.Name + " "+item.Id.ToString()+"\n";
                }
            }
            if (flag == false) {
                Element level_ele = doc.GetElement(levelid);
                notpass += "楼层名：" + level_ele.Name + ",楼层ID：" + levelid.ToString()+"\n";
            }
            
            return flag;
        }


    }
}
