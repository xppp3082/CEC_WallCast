﻿#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using System.Windows.Forms;
#endregion

namespace CEC_WallCast
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class WallCastUpdate : IExternalCommand
    {
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
        public static List<Transform> usefulLinkTrans = new List<Transform>();
        public static string errorOutput = "";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //程式要做的事情
            //1.蒐集所有的穿牆套管
            //2.蒐集所有的外參牆
            //3.更新穿牆套管與牆之間的關係
            //4.更新穿牆套管與管線間的關係
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();//引用stopwatch物件
            sw.Reset();//碼表歸零
            sw.Start();//碼表開始計時
            try
            {
                //準備東西
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                Document doc = uidoc.Document;

                //找到所有穿牆套管元件
                List<FamilyInstance> famList = findTargetElements(doc);
                List<string> usefulParaName = new List<string> { "BTOP", "BCOP", "BBOP", "TTOP", "TCOP", "TBOP",
                    "干涉管數量", "系統別","【原則檢討】是否穿牆"};
                //檢查參數
                foreach (FamilyInstance famInst in famList)
                {
                    foreach (string item in usefulParaName)
                    {
                        if (!checkPara(famInst, item))
                        {
                            MessageBox.Show($"執行失敗，請檢查{famInst.Symbol.FamilyName}元件中是否缺少{item}參數欄位");
                            return Result.Failed;
                        }
                    }
                }

                //子功能測試
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("更新穿牆套管資訊");
                    Dictionary<ElementId, List<Element>> wallCastDict = getCastWallDict(doc);
                    int transCount = 0;
                    foreach (ElementId id in wallCastDict.Keys)
                    {
                        updateCastWithWall(doc.GetElement(id), wallCastDict[id].First());
                        updateCastContent(doc, doc.GetElement(id));
                        //用來記錄transform的方法有點詭異，日後可能會出bug
                        modifyCastLen(doc.GetElement(id), wallCastDict[id].First(), usefulLinkTrans[transCount]);
                        transCount++;
                    }

                    //檢查所有的實做套管，如果不再Dictionary中，則表示其沒有穿牆
                    foreach (FamilyInstance inst in famList)
                    {
                        if (!wallCastDict.Keys.Contains(inst.Id))
                        {
                            inst.LookupParameter("干涉管數量").Set(0);
                            inst.LookupParameter("系統別").Set("SP");
                            inst.LookupParameter("【原則檢討】是否穿牆").Set("不符合");
                        }
                    }

                    trans.Commit();
                }
            }
            catch
            {
                MessageBox.Show("執行失敗");
                return Result.Failed;
            }
            sw.Stop();//碼錶停止
            double sec = Math.Round(sw.Elapsed.TotalMilliseconds / 1000, 2);
            string output = $"套管資訊更新完成，共花費 {sec} 秒\n";
            MessageBox.Show(output + errorOutput);
            return Result.Succeeded;
        }
        public static XYZ TransformPoint(XYZ point, Transform transform)
        {
            double x = point.X;
            double y = point.Y;
            double z = point.Z;

            //transform basis of the old coordinate system in the new coordinate // system
            XYZ b0 = transform.get_Basis(0);
            XYZ b1 = transform.get_Basis(1);
            XYZ b2 = transform.get_Basis(2);
            XYZ origin = transform.Origin;

            //transform the origin of the old coordinate system in the new 
            //coordinate system
            double xTemp = x * b0.X + y * b1.X + z * b2.X + origin.X;
            double yTemp = x * b0.Y + y * b1.Y + z * b2.Y + origin.Y;
            double zTemp = x * b0.Z + y * b1.Z + z * b2.Z + origin.Z;

            return new XYZ(xTemp, yTemp, zTemp);
        }
        public IList<Solid> GetTargetSolids(Element element)
        {
            List<Solid> solids = new List<Solid>();
            Options options = new Options();
            //預設為不包含不可見元件，因此改成true
            options.ComputeReferences = true;
            options.DetailLevel = ViewDetailLevel.Fine;
            options.IncludeNonVisibleObjects = true;
            GeometryElement geomElem = element.get_Geometry(options);
            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid)
                {
                    Solid solid = (Solid)geomObj;
                    if (solid.Faces.Size > 0 && solid.Volume > 0.0)
                    {
                        solids.Add(solid);
                    }
                }
                else if (geomObj is GeometryInstance)//一些特殊狀況可能會用到，like樓梯
                {
                    GeometryInstance geomInst = (GeometryInstance)geomObj;
                    GeometryElement instGeomElem = geomInst.GetInstanceGeometry();
                    foreach (GeometryObject instGeomObj in instGeomElem)
                    {
                        if (instGeomObj is Solid)
                        {
                            Solid solid = (Solid)instGeomObj;
                            if (solid.Faces.Size > 0 && solid.Volume > 0.0)
                            {
                                solids.Add(solid);
                            }
                        }
                    }
                }
            }
            return solids;
        }
        public Solid singleSolidFromElement(Element inputElement)
        {
            Document doc = inputElement.Document;
            Autodesk.Revit.ApplicationServices.Application app = doc.Application;
            // create solid from Element:
            IList<Solid> fromElement = GetTargetSolids(inputElement);
            int solidCount = fromElement.Count;
            // MessageBox.Show(solidCount.ToString());
            // Merge all found solids into single one
            Solid solidResult = null;
            //XYZ checkheight = new XYZ(0, 0, 6.88976);
            //Transform tr = Transform.CreateTranslation(checkheight);
            if (solidCount == 1)
            {
                solidResult = fromElement[0];
            }
            else if (solidCount > 1)
            {
                solidResult =
                    BooleanOperationsUtils.ExecuteBooleanOperation(fromElement[0], fromElement[1], BooleanOperationsType.Union);
            }

            if (solidCount > 2)
            {
                for (int i = 2; i < solidCount; i++)
                {
                    solidResult = BooleanOperationsUtils.ExecuteBooleanOperation(solidResult, fromElement[i], BooleanOperationsType.Union);
                }
            }
            return solidResult;
        }
        public bool checkPara(Element elem, string paraName)
        {
            bool result = false;
            foreach (Parameter parameter in elem.Parameters)
            {
                Parameter val = parameter;
                if (val.Definition.Name == paraName)
                {
                    result = true;
                }
            }
            return result;
        }
        public List<FamilyInstance> findTargetElements(Document doc)
        {
            //RC套管跟SC開口的內部名稱是不同的
            string internalName = "CEC-穿牆";
            List<FamilyInstance> castInstances = new List<FamilyInstance>();
            try
            {
                FilteredElementCollector coll = new FilteredElementCollector(doc);
                ElementCategoryFilter castCate_Filter = new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory);
                ElementClassFilter castInst_Filter = new ElementClassFilter(typeof(Instance));
                LogicalAndFilter andFilter = new LogicalAndFilter(castCate_Filter, castInst_Filter);
                coll.WherePasses(andFilter).WhereElementIsNotElementType().ToElements(); //找出模型中實做的穿牆套管
                if (coll != null)
                {
                    foreach (FamilyInstance e in coll)
                    {
                        Parameter p = e.Symbol.LookupParameter("API識別名稱");
                        if (p != null && p.AsString().Contains(internalName))
                        {
                            castInstances.Add(e);
                        }
                    }
                }
                else if (castInstances.Count() == 0)
                {
                    {
                        MessageBox.Show("尚未匯入套管元件，或模型中沒有實做的套管元件");
                    }
                }
            }
            catch
            {
                MessageBox.Show("蒐集套管發生問題!");
            }
            return castInstances;
        }
        public List<RevitLinkInstance> findWallLinkInstance(Document doc)
        {
            //找到擁有牆元件的外參
            List<RevitLinkInstance> linkInstanceList = new List<RevitLinkInstance>();
            FilteredElementCollector linkCollector = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));
            foreach (RevitLinkInstance linkInst in linkCollector)
            {
                Document linkDoc = linkInst.GetLinkDocument();
                if (linkDoc != null)
                {
                    FilteredElementCollector coll = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType();
                    if (coll.Count() > 0) linkInstanceList.Add(linkInst);
                }
            }
            return linkInstanceList;
        }
        public FilteredElementCollector getAllLinkedWall(Document linkedDoc)
        {
            FilteredElementCollector beamCollector = new FilteredElementCollector(linkedDoc).OfClass(typeof(Wall));
            beamCollector.WhereElementIsNotElementType();
            return beamCollector;
        }
        public Dictionary<ElementId, List<Element>> getCastWallDict(Document doc)
        {
            //蒐集套管與牆的關係
            Dictionary<ElementId, List<Element>> castWallDict = new Dictionary<ElementId, List<Element>>();
            try
            {
                //蒐集檔案中實做的牆套管
                List<FamilyInstance> familyInstances = findTargetElements(doc);
                List<RevitLinkInstance> linkWallInstances = findWallLinkInstance(doc);
                Transform totalTransform = null;
                Transform inverseTransform = null;
                if (linkWallInstances.Count != 0)
                {
                    foreach (FamilyInstance inst in familyInstances)
                    {
                        foreach (RevitLinkInstance linkInst in linkWallInstances)
                        {
                            //dictionary中尚未包含這個套管的ID則進行計算
                            //if (!castWallDict.Keys.Contains(inst.Id))
                            //{
                            totalTransform = linkInst.GetTotalTransform();
                            inverseTransform = totalTransform.Inverse;
                            FilteredElementCollector collectorWall = getAllLinkedWall(linkInst.GetLinkDocument());
                            Solid castSolid = singleSolidFromElement(inst);
                            if (castSolid == null) continue;

                            //座標轉換，因為抓法關係，要轉換成外參檔「原本」的座標
                            castSolid = SolidUtils.CreateTransformed(castSolid, inverseTransform);
                            BoundingBoxXYZ solidBounding = castSolid.GetBoundingBox();
                            XYZ solidCenter = castSolid.ComputeCentroid();
                            Transform newTrans = Transform.Identity;
                            newTrans.Origin = solidCenter;
                            Outline outLine = new Outline(newTrans.OfPoint(solidBounding.Min), newTrans.OfPoint(solidBounding.Max));

                            BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
                            ElementIntersectsSolidFilter elementIntersectsSolidFilter = new ElementIntersectsSolidFilter(castSolid);
                            collectorWall.WherePasses(boundingBoxIntersectsFilter).WherePasses(elementIntersectsSolidFilter);

                            List<Element> tempList = collectorWall.ToList();
                            //將蒐集到的東西放進dictionary中
                            if (tempList.Count > 0 && !castWallDict.Keys.Contains(inst.Id))
                            {
                                castWallDict.Add(inst.Id, tempList);
                                usefulLinkTrans.Add(totalTransform);
                            }
                            else if (tempList.Count > 0 && castWallDict.Keys.Contains(inst.Id))
                            {
                                foreach (Element e in tempList)
                                {
                                    castWallDict[inst.Id].Add(e);
                                }
                            }
                            ////在將東西放進Dictionary前，先確定長度與位置，調整一下
                            //modifyCastLen(inst, tempList.First(), totalTransform);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("無法判斷套管與牆的關係!");
            }
            return castWallDict;
        }
        public Element updateCastWithWall(Element elem, Element linkedWall)
        {
            FamilyInstance updateCast = null;
            try
            {
                FamilyInstance inst = elem as FamilyInstance;
                Document document = elem.Document;
                Wall wall = linkedWall as Wall;

                //調整高度與長度
                double castDiameter = inst.Symbol.LookupParameter("管外直徑").AsDouble() / 2;
                double castHeight = inst.LookupParameter("偏移").AsDouble() - castDiameter; //COP

                //設定BBOP、BCOP、BTOP (牆只需要設定從底部開始的參數)
                Parameter BBOP = inst.LookupParameter("BBOP");
                Parameter BCOP = inst.LookupParameter("BCOP");
                Parameter BTOP = inst.LookupParameter("BTOP");
                Parameter TTOP = inst.LookupParameter("TTOP");
                Parameter TCOP = inst.LookupParameter("TCOP");
                Parameter TBOP = inst.LookupParameter("TBOP");
                double outterDiameter = inst.Symbol.LookupParameter("管外直徑").AsDouble();
                double wallBaseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
                double BBOP_toSet = castHeight - wallBaseOffset - outterDiameter / 2;
                double BCOP_toSet = castHeight - wallBaseOffset;
                double BTOP_toSet = castHeight - wallBaseOffset + outterDiameter / 2;
                double wallHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
                double TCOP_toSet = wallHeight - BCOP_toSet;
                double TBOP_toSet = TCOP_toSet + outterDiameter / 2;
                double TTOP_toSet = TCOP_toSet - outterDiameter / 2;
                BBOP.Set(BBOP_toSet);
                BCOP.Set(BCOP_toSet);
                BTOP.Set(BTOP_toSet);
                TBOP.Set(TBOP_toSet);
                TCOP.Set(TCOP_toSet);
                TTOP.Set(TTOP_toSet);
                inst.LookupParameter("【原則檢討】是否穿牆").Set("OK");
                updateCast = inst;
            }
            catch
            {
                errorOutput += $"更新套管資訊失敗，ID為 {elem.Id} 的套管無法更新資訊!\n";
            }
            return updateCast;
        }
        public List<RevitLinkInstance> getMEPLinkInstance(Document doc)
        {
            List<RevitLinkInstance> linkInstanceList = new List<RevitLinkInstance>();
            FilteredElementCollector linkCollector = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));
            //製作針對管、風管、電管的過濾器
            BuiltInCategory[] builts =
            {
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_DuctCurves
            };
            List<ElementFilter> filters = new List<ElementFilter>();
            foreach (BuiltInCategory built in builts)
            {
                ElementCategoryFilter filter = new ElementCategoryFilter(built);
                filters.Add(filter);
            }
            LogicalOrFilter categoryFilter = new LogicalOrFilter(filters);
            foreach (RevitLinkInstance linkInst in linkCollector)
            {
                Document linkDoc = linkInst.GetLinkDocument();
                if (linkDoc != null)
                {
                    FilteredElementCollector coll = new FilteredElementCollector(linkDoc).WherePasses(categoryFilter).WhereElementIsNotElementType();
                    if (coll.Count() > 0) linkInstanceList.Add(linkInst);
                }
            }
            return linkInstanceList;
        }
        public Element updateCastContent(Document doc, Element elem)
        {
            FamilyInstance updateCast = null;
            try
            {
                List<string> systemName = new List<string>() { "E", "T", "W", "P", "F", "A", "G" };
                FamilyInstance inst = elem as FamilyInstance;
                BuiltInCategory[] builts =
                {
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_DuctCurves
            };
                List<ElementFilter> filters = new List<ElementFilter>();
                foreach (BuiltInCategory built in builts)
                {
                    ElementCategoryFilter filter = new ElementCategoryFilter(built);
                    filters.Add(filter);
                }
                LogicalOrFilter categoryFilter = new LogicalOrFilter(filters);
                FilteredElementCollector pipeCollector = new FilteredElementCollector(doc).WherePasses(categoryFilter);
                BoundingBoxXYZ castBounding = inst.get_BoundingBox(null);
                Outline castOutline = new Outline(castBounding.Min, castBounding.Max);
                BoundingBoxIntersectsFilter boxIntersectsFilter = new BoundingBoxIntersectsFilter(castOutline);
                Solid castSolid = singleSolidFromElement(inst);
                ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(castSolid);
                pipeCollector.WherePasses(boxIntersectsFilter).WherePasses(solidFilter);
                Parameter systemType = inst.LookupParameter("系統別");

                List<RevitLinkInstance> otherMEP = getMEPLinkInstance(doc);
                List<Element> pipeCollector_final = new List<Element>();
                //將本機端蒐集到的管放進list
                foreach (Element e in pipeCollector)
                {
                    pipeCollector_final.Add(e);
                }
                //foreach(RevitLinkInstance link in otherMEP)
                //{
                //    MessageBox.Show(link.GetLinkDocument().Title);
                //}

                //針對每一個實做的外參，蒐集管段後加入list
                foreach (RevitLinkInstance linkInst in otherMEP)
                {
                    //座標變換，因為抓法的關係，要轉換成外參檔「原本」的座標
                    Transform trans = linkInst.GetTotalTransform();
                    Transform inverseTransform = trans.Inverse;
                    Document linkdoc = linkInst.GetLinkDocument();
                    if (linkdoc != null)
                    {
                        Solid CSDsolid = SolidUtils.CreateTransformed(castSolid, inverseTransform);
                        ElementIntersectsSolidFilter solidFilter_CSD = new ElementIntersectsSolidFilter(CSDsolid);
                        BoundingBoxXYZ solidBounding = CSDsolid.GetBoundingBox();
                        XYZ solidCenter = CSDsolid.ComputeCentroid();
                        Transform newTrans = Transform.Identity;
                        newTrans.Origin = solidCenter;
                        Outline outLine = new Outline(newTrans.OfPoint(solidBounding.Min), newTrans.OfPoint(solidBounding.Max));
                        BoundingBoxIntersectsFilter boxIntersectsFilter_CSD = new BoundingBoxIntersectsFilter(outLine);
                        FilteredElementCollector CSDcollector = new FilteredElementCollector(linkdoc).WherePasses(categoryFilter).WherePasses(boxIntersectsFilter_CSD).WherePasses(solidFilter_CSD);

                        if (CSDcollector.Count() > 0)
                        {
                            foreach (Element e in CSDcollector)
                            {
                                pipeCollector_final.Add(e);
                            }
                        }
                    }
                }
                if (pipeCollector_final.Count() == 0)
                {
                    systemType.Set("SP");
                }
                inst.LookupParameter("干涉管數量").Set(pipeCollector_final.Count());
                //針對蒐集到的管去做系統別的更新，因為電管沒有系統類型，要和管分開處理
                if (pipeCollector_final.Count() == 1)
                {
                    if (pipeCollector_final.First().Category.Name == "電管")
                    {
                        systemType.Set("E");
                    }
                    else
                    {
                        Element targetPipe = pipeCollector_final.First();
                        string pipeSystem = null;
                        if (targetPipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM) != null)
                        {
                            pipeSystem = targetPipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                        }
                        else if (targetPipe.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM) != null)
                        {
                            pipeSystem = targetPipe.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString();
                        }
                        string shortSystemName = pipeSystem.Substring(0, 1);//以開頭的前綴字抓系統縮寫
                        if (systemName.Contains(shortSystemName))
                        {
                            systemType.Set(shortSystemName);
                        }
                        else if (pipeSystem.Contains("空調"))
                        {
                            systemType.Set("A");
                        }
                        else
                        {
                            systemType.Set("SP");
                        }
                    }
                }
                //如果有共管的狀況
                else if (pipeCollector_final.Count() >= 2)
                {
                    List<int> isPipe = new List<int>();
                    List<string> shortNameList = new List<string>();
                    foreach (Element pipe in pipeCollector_final)
                    {
                        if (pipe.Category.Name == "電管")
                        {
                            isPipe.Add(0);
                            shortNameList.Add("E");
                        }
                        else if (pipe.Category.Name == "管")
                        {
                            isPipe.Add(1);
                            string pipeSystem = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                            string shortSystemName = pipeSystem.Substring(0, 1);//以開頭的前綴字抓系統縮寫
                            shortNameList.Add(shortSystemName);
                        }
                        else if (pipe.Category.Name == "風管")
                        {
                            isPipe.Add(2);
                            string pipeSystem = pipe.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString();
                            string shortSystemName = pipeSystem.Substring(0, 1);//以開頭的前綴字抓系統縮寫
                            //shortNameList.Add(shortSystemName);
                            shortNameList.Add(shortSystemName);
                        }
                    }
                    List<int> newTypeList = isPipe.Distinct().ToList();
                    List<string> newList = shortNameList.Distinct().ToList();
                    //先用類別判斷，在用是否共管判斷，就算共管，如果同系統還是得寫一樣的縮寫名稱
                    if (newTypeList.Count() == 1 && newTypeList.First() == 0)
                    {
                        systemType.Set("E");
                    }
                    else if (newTypeList.Count() >= 1 /*&& newTypeList.First() == true*/)
                    {
                        //因為剛好空調也叫M，才需要特別處理
                        if (newList.Count() == 1 && newList.Contains("M"))
                        {
                            //systemType.Set(newList.First());
                            systemType.Set("A");
                        }
                        else if (newList.Count() == 1 && newList.First() != "M")
                        {
                            systemType.Set(newList.First());
                        }
                        //如果為不同系統共管，則設為M
                        else if (newList.Count() > 1)
                        {
                            systemType.Set("M");
                        }
                    }
                    //else if (newTypeList.Count() > 1)
                    //{
                    //    systemType.Set("M");
                    //}
                }
                updateCast = inst;
            }
            catch
            {
                errorOutput += $"套管干涉檢查失敗，ID為 {elem.Id} 的套管無法順利檢查\n";
                //MessageBox.Show($"穿樑套管干涉檢查失敗，ID為{elem.Id}的套管無法順利檢查!");
            }
            return updateCast;
        }
        public  Element modifyCastLen(Element elem , Element linkedWall,Transform toTrans)
        {
            FamilyInstance updateCast = null;
            try
            {
                Document document = elem.Document;
                FamilyInstance inst = elem as FamilyInstance;
                Wall wall = linkedWall as Wall;
                LocationCurve wallLocate = linkedWall.Location as LocationCurve;
                Curve beamCurve = wallLocate.Curve;
                XYZ startPoint = beamCurve.GetEndPoint(0);
                XYZ endPoint = beamCurve.GetEndPoint(1);
                startPoint = TransformPoint(startPoint, toTrans);
                endPoint = TransformPoint(endPoint, toTrans);
                endPoint = new XYZ(endPoint.X, endPoint.Y, startPoint.Z);
                Line tempCrv = Line.CreateBound(startPoint, endPoint);

                LocationPoint castLocate = elem.Location as LocationPoint;
                XYZ castPt = castLocate.Point;
                XYZ tempPt = new XYZ(castPt.X, castPt.Y, startPoint.Z);
                IntersectionResult intersectResult = tempCrv.Project(castPt);
                XYZ targetPoint = intersectResult.XYZPoint;
                targetPoint = new XYZ(targetPoint.X, targetPoint.Y, castPt.Z);
                XYZ positionChange = targetPoint - castPt;
                double castLength = UnitUtils.ConvertFromInternalUnits(wall.Width, unitType) + 20;
                double castLength_toSet = UnitUtils.ConvertToInternalUnits(castLength ,unitType);
                Parameter instLenPara = inst.LookupParameter("L");

                //先調整位置
                if (!castPt.IsAlmostEqualTo(targetPoint))
                {
                    ElementTransformUtils.MoveElement(document, inst.Id, positionChange);
                }
                //在調整長度
                if (instLenPara.AsDouble() < wall.Width)
                {
                    instLenPara.Set(castLength_toSet);
                }
                updateCast = inst;
            }
            catch
            {
                errorOutput += $"更新套管長度失敗，ID為 {elem.Id} 的套管無法更新長度!\n";
            }
            return updateCast;
        }
    }
}
