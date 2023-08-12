using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB.Architecture;
using ZBIMUtils;


namespace UiLable
{

    [Transaction(TransactionMode.Manual)]
    class StairLandingHeightAssessment : IExternalCommand
    {
        private void HighLint(UIDocument uiDoc, IList<ElementId> lElementIds)
        {
            var sel = uiDoc.Selection.GetElementIds();
            foreach (var item in lElementIds)
            {
                sel.Add(item);
            }
            uiDoc.Selection.SetElementIds(sel);
        }

        public bool IsIndivied(Element element,Document doc, UIDocument uiDoc,double height)
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
                List<XYZ> facePoint = GetFaceCenter(facetemp);

                ElementId elementCreateId;
                elementCreateId = CreateCurve(doc, facePoint, height);
                Element elementCreate = doc.GetElement(elementCreateId);
                List<Element> lElements = GetIntersectsElements(elementCreate, doc);
                
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
                return true;
            }
            else
            {
                //TaskDialog.Show("楼梯检测结果", "楼梯平台上部及下部过道处的净高不应小于2.0m，梯段净高不应小于2.2m。->检测通过");
                return false;
            }
        }
        public static List<Element> GetIntersectsElements(Element e, Document document)
        {
            //对于solid的boundingbox属性，API文档给出提示信息：
            //solid的boundingbox是局部坐标系中的属性，需要将其转换为世界坐标系
            var transform = e.get_BoundingBox(document.ActiveView).Transform;

            var minSolid = e.get_BoundingBox(document.ActiveView).Min;
            var maxSolid = e.get_BoundingBox(document.ActiveView).Max;

            var acturalMin = transform.OfPoint(minSolid);
            var acturalMax = transform.OfPoint(maxSolid);

            var outline = new Outline(acturalMin, acturalMax);
            var boxFilter = new BoundingBoxIntersectsFilter(outline);
            var collector = new FilteredElementCollector(document).OfClass(typeof(FamilyInstance));
            var intersectElements = collector
                .WherePasses(boxFilter)
                .ToList();

            return intersectElements;
        }
        public List<Element> GetInstanceListInSpace(Element elem, int hight, FilteredElementCollector collector)//1、当前元素  2、高度 3.碰撞检测目标
        {

            ElementIntersectsElementFilter filter = new ElementIntersectsElementFilter(elem);//用API过滤相交元素
            List<Element> lstElem = collector.WherePasses(filter).ToElements().ToList();
            return lstElem;
        }
        public List<XYZ> GetFaceCenter(Face temFace)
        {
            Mesh mesh = temFace.Triangulate();
            List<XYZ> pointList = new List<XYZ>();
            int num = 4;
            if (temFace != null)
            {          
                foreach (XYZ ii in mesh.Vertices)
                {     
                        pointList.Add(ii);
                }
            }
            //TaskDialog.Show("点", pointList.Count.ToString());
            
            //while (pointList.Count > num)
            //{
            //    pointList.RemoveAt(0);
            //}
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
            //创建一个长方体拉伸模型
            using (Transaction tran = new Transaction(doc, "拉伸"))
            {
                XYZ up = new XYZ(0,0,0.001);
                tran.Start();
                XYZ midOfP0 = (p[0] + (p[2] - p[0]) / 4)  + up;
                XYZ midOfP1 = (p[1] + (p[3] - p[1]) / 4)  + up;
                XYZ midOfP2 = (p[2] - (p[2] - p[0]) / 4) + up;
                XYZ midOfP3 = (p[3] - (p[3] - p[1]) / 4) + up;
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
                Solid solid = GeometryCreationUtilities.CreateExtrusionGeometry(loops, direction, height/0.3048);
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









        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            ElementCategoryFilter elementCategoryFilter2 = new ElementCategoryFilter(BuiltInCategory.OST_Stairs);

            collector2.WherePasses(elementCategoryFilter2);
            //TaskDialog.Show("楼梯数量", collector2.GetElementCount().ToString());
            List<ElementId> errElements = new List<ElementId>();
            StringBuilder stringBuilder_notpass_stairs = new StringBuilder(); 
            StringBuilder stringBuilder_notpass_stairslanding = new StringBuilder(); 
            int passNum = 0;
            foreach (var item in collector2)
            {
                Stairs stairs = item as Stairs;
                if (stairs != null)
                {
                    IList<ElementId> stairsRun = stairs.GetStairsRuns().ToList<ElementId>();

                    if (stairsRun.Count != 0)
                    {
                        //HighLint(uiDoc, stairsRun);
                        foreach (ElementId stairsRunId in stairsRun)
                        {
                            if (!IsIndivied(doc.GetElement(stairsRunId), doc, uiDoc, 2))
                            {
                                //errElements.Add(stairsRunId);
                                stringBuilder_notpass_stairs.AppendLine(stairsRunId.ToString());
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
                            if (!IsIndivied(doc.GetElement(stairsLandId), doc, uiDoc, 2.2))
                            {
                                errElements.Add(stairsLandId);
                                stringBuilder_notpass_stairslanding.AppendLine(stairsLandId.ToString());
                            }
                            else
                            {
                                passNum++;
                            }
                        }
                    }
                }
            }


            //输出不符合
            if (errElements.Count != 0)
            {
                TaskDialog.Show("楼梯检测结果", "楼梯平台上部及下部过道处的净高不应小于2.0m，梯段净高不应小于2.2m。->检测不通过");
                String err = "通过数为" + passNum + "\n";
                int num = 0;
                foreach (ElementId s in errElements)
                {
                    if (num == 5)
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

            //IsIndivied(doc.GetElement(new ElementId(2602683)), doc, uiDoc, 2.2);
            return Result.Succeeded;
        }









    }
    
}
