using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace Stair
{

    [Transaction(TransactionMode.Manual)]
    class StairLandingHeightAssessment : IExternalCommand
    {

        private double StairsLandLimitHeight = 2.2;
        private double StairsRunLimitHeight = 2.0;
        private double SuspensionHeight = 1.0;//悬浮高度，使创建的立方体离开地面
        private int ScaleFactor = 10;//创建立方体的缩放因子
        ExternalCommandData commandDataGloble;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            commandDataGloble = commandData;
            DateTime dt1 = DateTime.Now;

            StairMonitor();

            DateTime dt2 = DateTime.Now;
            TimeSpan ts = dt2.Subtract(dt1);
            TaskDialog.Show("程序耗时：{0}ms.", ts.TotalMilliseconds.ToString());

            
            return Result.Succeeded;
        }

        private void StairMonitor()
        {
            //TaskDialog.Show("楼梯数量","");
            //p.ShowDialog();
            //界面交互的doc
            UIDocument uiDoc = commandDataGloble.Application.ActiveUIDocument;
            //实际内容的doc
            Document doc = commandDataGloble.Application.ActiveUIDocument.Document;
            
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_Stairs);

            collector.WherePasses(elementCategoryFilter);
            //TaskDialog.Show("楼梯数量", collector2.GetElementCount().ToString());
            List<ElementId> errElements = new List<ElementId>();
            int passNum = 0;

            //非相关元素
            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            List<Element> excludingIds = collector2.OfCategory(BuiltInCategory.OST_StructuralColumns).ToList<Element>();
            ICollection<ElementId> idsToExclude = new List<ElementId>();
            foreach (var element in excludingIds)
            {
                idsToExclude.Add(element.Id);
            }
            foreach (var item in collector)
            {
                //if (Pressindex == 10) break;
                Autodesk.Revit.DB.Architecture.Stairs stairs = item as Autodesk.Revit.DB.Architecture.Stairs;
                if (stairs != null)
                {
                    IList<ElementId> stairsRun = stairs.GetStairsRuns().ToList<ElementId>();
                    if (stairsRun.Count != 0)
                    {
                        //HighLint(uiDoc, stairsRun);
                        foreach (ElementId stairsRunId in stairsRun)
                        {
                            if (!IsIndivied(doc.GetElement(stairsRunId), doc, uiDoc, StairsRunLimitHeight, idsToExclude))
                            {
                                errElements.Add(stairsRunId);
                            }
                            else
                            {
                                passNum++;
                            }
                        }
                    }
                    IList<ElementId> stairsLand = stairs.GetStairsLandings().ToList<ElementId>();
                    if (stairsLand.Count != 0)
                    {
                        //HighLint(uiDoc, stairsLand);
                        foreach (ElementId stairsLandId in stairsLand)
                        {
                            if (!IsIndivied(doc.GetElement(stairsLandId), doc, uiDoc, StairsLandLimitHeight, idsToExclude))
                            {
                                errElements.Add(stairsLandId);

                            }
                            else
                            {
                                passNum++;
                            }
                        }
                    }
                }
               
            }
            CheckErrElements(errElements,uiDoc,passNum);
            //IsIndivied(doc.GetElement(new ElementId(2602683)), doc, uiDoc, 2.2);
        }
        public void CheckErrElements(List<ElementId> errElements, UIDocument uiDoc,int passNum)
        {
            if (errElements.Count != 0)
            {
                TaskDialog.Show("楼梯检测结果", "楼梯平台上部及下部过道处的净高不应小于2.0m，梯段净高不应小于2.2m。->检测不通过");
                HighLint(uiDoc, errElements);
                String err = "楼梯组件检测通过数为" + passNum + "\n" + "未通过测试的楼梯编号:\n";
                int num = 0;
                foreach (ElementId s in errElements)
                {
                    if (num == 3)
                    {
                        err += s.ToString() + "\n";
                        num = 0;
                    }
                    else
                    {
                        err += s.ToString() + " ";
                        num++;
                    }

                }
                TaskDialog.Show("楼梯检测问题编号", err);
            }
            else
            {
                TaskDialog.Show("楼梯检测结果", "楼梯平台上部及下部过道处的净高不应小于2.0m，梯段净高不应小于2.2m。->检测通过");
            }
        }

        public bool IsIndivied(Element element,Document doc, UIDocument uiDoc,double height, ICollection<ElementId> idsToExclude)
        {
            Options options = new Options();
            GeometryElement geometry = element.get_Geometry(options);
            Face facetemp = null;

            try
            {
                foreach (GeometryObject obj in geometry)
                {
                    Solid solid = obj as Solid;
                    if (solid != null)
                    {
                        FaceArray faceArray = solid.Faces;
                        foreach (Face face in faceArray)
                        {
                            if (facetemp == null)
                            {
                                facetemp = faceArray.get_Item(0);
                            }
                            if (GreaterThan(face, facetemp))
                            {
                                facetemp = face;//face在facetemp上面且面为顶面
                            }
                        }
                        if (facetemp == null)
                        {
                            throw new Exception("face 为 null"); 
                        }
                    }
                }
                List<XYZ> facePoint = GetFacePoints(facetemp);

                ElementId elementCreateId;
                elementCreateId = CreateCurve(doc, facePoint, height);
                Element elementCreate = doc.GetElement(elementCreateId);


                List<Element> lElements = GetIntersectsElements(elementCreate, doc, idsToExclude);
                
                DeleteCurve(doc, elementCreate);//删除刚刚创建的立方
                return CheckRes(uiDoc, lElements);
            }
            catch (Exception e)
            {
               
                //TaskDialog.Show("异常结果", element.Id.ToString()+e.ToString());
                //Environment.Exit(0);
                return false;
            }
        }

        private bool GreaterThan(Face face, Face facetemp)
        {
            Mesh mesh = face.Triangulate();
            Mesh meshTemp = facetemp.Triangulate();
            List<XYZ> pointList = new List<XYZ>();
            double Ztemp;
            if (mesh!=null)
            {
                Ztemp = mesh.Vertices[0].Z;
                foreach (XYZ ii in mesh.Vertices)
                {
                    if (ii.Z!= Ztemp)//说明是不同一平面
                    {
                        return false;
                    }
                }
                return mesh.Vertices[0].Z > meshTemp.Vertices[0].Z;//如果是同一平面的点则和faceTemp对比
            }
            return false;
        }
        //高亮显示
        private void HighLint(UIDocument uiDoc, IList<ElementId> lElementIds)
        {
            var sel = uiDoc.Selection.GetElementIds();
            foreach (var item in lElementIds)
            {
                sel.Add(item);
            }
            uiDoc.Selection.SetElementIds(sel);
        }
        //高亮显示
        public void HighLint(UIDocument uiDoc, IList<Element> lElements)
        {
            var sel = uiDoc.Selection.GetElementIds();
            foreach (var item in lElements)
            {
                sel.Add(item.Id);
            }
            uiDoc.Selection.SetElementIds(sel);        
        }
        //检测结果展示
        public bool CheckRes(UIDocument uiDoc, List<Element> lElements)
        {
            
            if (lElements.Count != 0)
            {               
                //TaskDialog.Show("楼梯检测结果", "楼梯平台上部及下部过道处的净高不应小于2.0m，梯段净高不应小于2.2m。->检测不通过");
                HighLint(uiDoc, lElements);
                return false;
            }
            else
            {
                //TaskDialog.Show("楼梯检测结果", "楼梯平台上部及下部过道处的净高不应小于2.0m，梯段净高不应小于2.2m。->检测通过");
                return true;
            }
        }
        public static List<Element> GetIntersectsElements(Element e, Document document, ICollection<ElementId> idsToExclude)
        {
            var collector = new FilteredElementCollector(document).OfClass(typeof(FamilyInstance));//定义碰撞可能发生的类型
            // ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_StructuralColumns);//不可能碰撞的
            //todo 划定可能碰撞的元素。
            ElementIntersectsElementFilter iFilter = new ElementIntersectsElementFilter(e, false);   
            var intersectElements = collector
                .Excluding(idsToExclude)
                .WherePasses(iFilter)
                .ToList();
            return intersectElements;
        }
        public List<Element> GetInstanceListInSpace(Element elem, int hight, FilteredElementCollector collector)//1、当前元素  2、高度 3.碰撞检测目标
        {

            ElementIntersectsElementFilter filter = new ElementIntersectsElementFilter(elem);//用API过滤相交元素
            List<Element> lstElem = collector.WherePasses(filter).ToElements().ToList();
            return lstElem;
        }
        public List<XYZ> GetFacePoints(Face temFace)
        {
            Mesh mesh = temFace.Triangulate();
            List<XYZ> pointList = new List<XYZ>();
            
            if (temFace != null)
            {          
                foreach (XYZ ii in mesh.Vertices)
                {     
                        pointList.Add(ii);
                }
            }
            double Xmin= 1000000, Xmax = -1000000, Ymin = 1000000, Ymax = -1000000, Z = -1000000;//设置最小值
            for (int i=0;i<pointList.Count;i++)
            {
                XYZ temp = pointList.ElementAt(i);
                Xmin = Math.Min(Xmin,temp.X);
                Xmax = Math.Max(Xmax, temp.X);
                Ymin = Math.Min(Ymin, temp.Y);
                Ymax = Math.Max(Ymax, temp.Y);
                Z = temp.Z;
            }
            while (pointList.Count > 0)
            {
                pointList.RemoveAt(0);
            }
            pointList.Add(new XYZ(Xmin,Ymin,Z));
            pointList.Add(new XYZ(Xmax, Ymin, Z));
            pointList.Add(new XYZ(Xmax, Ymax, Z));
            pointList.Add(new XYZ(Xmin, Ymax, Z));
            if (pointList.Count<4)
            {
                //TaskDialog.Show("点结果少于4个", pointList.Count.ToString());
                throw new Exception("点结果少于4个"); 
            }
            return pointList;
        }

        public ElementId CreateCurve(Document doc,List<XYZ> p,double height)//创建立方体
        {
            //创建一个立方体拉伸模型
            using (Transaction tran = new Transaction(doc, "拉伸"))
            {
                XYZ up = new XYZ(0,0, SuspensionHeight / 0.3048);
                tran.Start();
                XYZ midOfP0 = (p[0] + (p[2] - p[0]) / ScaleFactor)  + up;
                XYZ midOfP1 = (p[1] + (p[3] - p[1]) / ScaleFactor)  + up;
                XYZ midOfP2 = (p[2] - (p[2] - p[0]) / ScaleFactor) + up;
                XYZ midOfP3 = (p[3] - (p[3] - p[1]) / ScaleFactor) + up;
                //midOfP0 = p[0] - (p[3] - p[2]) / 2;
                //Line line1 = Line.CreateBound(p[0], p[1]);
                //Line line2 = Line.CreateBound(p[1], p[2]);
                //Line line3 = Line.CreateBound(p[2], p[3]);
                //Line line4 = Line.CreateBound(p[3], p[0]);
                Line line1 = Line.CreateBound(midOfP0, midOfP1);

                Line line2 = Line.CreateBound(midOfP1, midOfP2);
                Line line3 = Line.CreateBound(midOfP2, midOfP3);
                Line line4 = Line.CreateBound(midOfP3, midOfP0);
                CurveLoop loop = new CurveLoop();
                loop.Append(line1);
                loop.Append(line2);
                loop.Append(line3);
                loop.Append(line4);
                List<CurveLoop> loops = new List<CurveLoop>() { loop };
                XYZ direction = new XYZ(0, 0, 1);//拉伸方向
                Solid solid = GeometryCreationUtilities.CreateExtrusionGeometry(loops, direction, (height-SuspensionHeight)/ 0.3048);
                ElementId categoryId = new ElementId(BuiltInCategory.OST_StructuralFoundation);
                DirectShape shape = DirectShape.CreateElement(doc, categoryId);
                shape.AppendShape(new List<GeometryObject>() { solid });
                tran.Commit();
                //TaskDialog.Show("查看结果", shape.Id.ToString());
                return shape.Id;
            }
        }
        public void DeleteCurve(Document doc, Element element)//删除立方体
        {
            using (Transaction tran = new Transaction(doc, "删除"))
            {
                tran.Start();
                doc.Delete(element.Id);
                tran.Commit();
            }
        }
    }
    
}
