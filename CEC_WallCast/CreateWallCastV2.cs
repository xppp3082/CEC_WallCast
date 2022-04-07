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
    class CreateWallCastV2 : IExternalCommand
    {
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            while (true)
            {
                try
                {
                    //準備東西
                    UIApplication uiapp = commandData.Application;
                    UIDocument uidoc = uiapp.ActiveUIDocument;
                    Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                    Document doc = uidoc.Document;

                    //拿到管元件
                    ISelectionFilter pipeFilter = new PipeSelectionFilter(doc);
                    Reference pickPipeRef = uidoc.Selection.PickObject(ObjectType.Element, pipeFilter, "請選擇貫穿牆的管");
                    Element pickPipe = doc.GetElement(pickPipeRef.ElementId);
                    //拿到整份牆外參檔&Transform
                    ISelectionFilter linkedWallFilter = new WallSelectionFilter(doc);
                    Reference pickWallRef = uidoc.Selection.PickObject(ObjectType.LinkedElement, linkedWallFilter, "請選擇貫穿的牆");
                    RevitLinkInstance pickWall = doc.GetElement(pickWallRef) as RevitLinkInstance;
                    Transform linkTransform = pickWall.GetTotalTransform();

                    //拿到實際要用的那道牆元件
                    Element linkedWall = pickWall.GetLinkDocument().GetElement(pickWallRef.LinkedElementId);
                    Wall wall = linkedWall as Wall;
                    double holeLength = UnitUtils.ConvertFromInternalUnits(wall.Width, unitType) + 20;
                    LocationCurve wallLocate = linkedWall.Location as LocationCurve;
                    Line wallLine = wallLocate.Curve as Line;
                    XYZ wallDir = wallLine.Direction;
                    XYZ wallNorDir = wallDir.CrossProduct(XYZ.BasisZ).Normalize().Negate();
                    XYZ holeDir = XYZ.BasisY;
                    double angle = holeDir.AngleTo(wallNorDir);


                    MEPCurve pipeCrv = pickPipe as MEPCurve;
                    Level level = pipeCrv.ReferenceLevel; //取得管線的參考樓層
                    double elevation = level.Elevation;
                    XYZ HoleLocation = GetHoleLocation(linkedWall, pickPipe, linkTransform);
                    if (HoleLocation == null)
                    {
                        MessageBox.Show("選中的管並未和牆交集");
                    }

                    Family Wall_Cast;
                    FamilyInstance instance = null;
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("載入檔案測試");
                        Wall_Cast = new WallCast().WallCastSymbol(doc);
                        tx.Commit();
                    }

                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("放置穿樑套管");
                        FamilySymbol CastSymbol2 = new WallCast().findWall_CastSymbol(doc, Wall_Cast, pickPipe);
                        instance = doc.Create.NewFamilyInstance(HoleLocation, CastSymbol2, level, StructuralType.NonStructural);
                        //參數檢查
                        List<string> paraNameToCheck = new List<string>()
                                {
                                   "L","系統別","TTOP","TCOP","TBOP","BBOP","BCOP","BTOP"
                                };

                        foreach (string item in paraNameToCheck)
                        {
                            if (!checkPara(instance, item))
                            {
                                MessageBox.Show($"執行失敗，請檢查{instance.Symbol.FamilyName}元件中是否缺少{item}參數欄位");
                                return Result.Failed;
                            }
                        }

                        //調整高度與長度
                        double castDiameter = instance.Symbol.LookupParameter("管外直徑").AsDouble() / 2;
                        double pipeHeight = pickPipe.LookupParameter("偏移").AsDouble();
                        instance.LookupParameter("偏移").Set(pipeHeight + castDiameter);
                        double castLength = UnitUtils.ConvertToInternalUnits(holeLength, unitType);
                        instance.LookupParameter("L").Set(castLength);
                        LocationPoint loPoint = instance.Location as LocationPoint;
                        XYZ newPoint = new XYZ(HoleLocation.X, HoleLocation.Y, HoleLocation.Z + 10);
                        Line axis = Line.CreateBound(HoleLocation, newPoint);
                        loPoint.Rotate(axis, angle);


                        //設定BBOP、BCOP、BTOP (牆只需要設定從底部開始的參數)
                        Parameter BBOP = instance.LookupParameter("BBOP");
                        Parameter BCOP = instance.LookupParameter("BCOP");
                        Parameter BTOP = instance.LookupParameter("BTOP");
                        Parameter TTOP = instance.LookupParameter("TTOP");
                        Parameter TCOP = instance.LookupParameter("TCOP");
                        Parameter TBOP = instance.LookupParameter("TBOP");
                        double outterDiameter = instance.Symbol.LookupParameter("管外直徑").AsDouble();
                        double wallBaseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
                        double BBOP_toSet = pipeHeight - wallBaseOffset - outterDiameter / 2;
                        double BCOP_toSet = pipeHeight - wallBaseOffset;
                        double BTOP_toSet = pipeHeight - wallBaseOffset + outterDiameter / 2;
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
                        trans.Commit();
                    }

                }
                catch
                {
                    //MessageBox.Show("執行失敗");
                    //return Result.Failed;
                    break;
                }
            }
            return Result.Succeeded;
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

        public Solid singleSolidFromElement(Element element)
        {
            Options options = new Options();
            GeometryElement geomElem = element.get_Geometry(options);
            Solid solidResult = null;
            foreach (GeometryObject geomObj in geomElem)
            {
                solidResult = geomObj as Solid;
            }
            return solidResult;
        }
        public XYZ GetHoleLocation(Element wallElem, Element pipeElem, Transform trans)
        {
     
            //取得牆的solid
            Solid solid_wall = singleSolidFromElement(wallElem);
            solid_wall = SolidUtils.CreateTransformed(solid_wall, trans);
            //取得管的curve
            LocationCurve pipe_locationcurve = pipeElem.Location as LocationCurve;
            Curve pipe_Line = pipe_locationcurve.Curve;
            //取得管與牆相交的curve
            SolidCurveIntersectionOptions intersectOptions = new SolidCurveIntersectionOptions();
            SolidCurveIntersection intersection = solid_wall.IntersectWithCurve(pipe_Line, intersectOptions);
            XYZ point_Center = null;
            if (intersection.SegmentCount > 0)
            {
                Curve curveInside = intersection.GetCurveSegment(0);
                //取得curve中點
                XYZ point_Start = curveInside.GetEndPoint(0);
                XYZ point_End = curveInside.GetEndPoint(1);
                XYZ tempCenter = curveInside.Evaluate(0.5, true);
                //point_Center = new XYZ((point_Start.X + point_End.X) / 2, (point_Start.Y + point_End.Y) / 2, ((point_Start.Z + point_End.Z) / 2));
                point_Center = tempCenter;

            }
            return point_Center;

        }
        public class WallCast
        {
            #region
            //將穿牆套管的功能做成class管理
            //1.先找到套管的Family
            //2.用Family反查Symbol
            #endregion
            public Family WallCastSymbol(Document doc)
            {
                string internalNameWall = "CEC-穿牆套管";
                Family Wall_CastType = null;
                ElementFilter Wall_CastCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory);
                ElementFilter Wall_CastSymbolFilter = new ElementClassFilter(typeof(FamilySymbol));

                LogicalAndFilter andFilter = new LogicalAndFilter(Wall_CastCategoryFilter, Wall_CastSymbolFilter);
                FilteredElementCollector RC_CastSymbol = new FilteredElementCollector(doc);
                RC_CastSymbol = RC_CastSymbol.WherePasses(andFilter);//這地方有點怪，無法使用andFilter RC_CastSymbolFilter
                bool symbolFound = false;
                foreach (FamilySymbol e in RC_CastSymbol)
                {
                    Parameter p = e.LookupParameter("API識別名稱");
                    if (p != null && p.AsString().Contains(internalNameWall))
                    {
                        symbolFound = true;
                        Wall_CastType = e.Family;
                        break;
                    }
                }
                if (!symbolFound)
                {
                    MessageBox.Show("尚未載入指定的穿牆套管元件!");
                }
                return Wall_CastType;
            }

            public FamilySymbol findWall_CastSymbol(Document doc, Family CastFamily, Element element)
            {
                FamilySymbol targetFamilySymbol = null; //用來找目標familySymbol
                                                        //如果確定找到family後，針對不同得管選取不同的穿樑套管大小，以大兩吋為規則，如果有坡度則大三吋
                Parameter targetPara = null;
                //Pipe
                if (element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM) != null)
                {
                    targetPara = element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                }
                //Conduit
                else if (element.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM) != null)
                {
                    targetPara = element.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                }
                //Duct
                else if (element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM) != null)
                {
                    targetPara = element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                }
                else if (element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM) != null)
                {
                    targetPara = element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                }
                //利用管徑(doubleType)來判斷
                var covertUnit = UnitUtils.ConvertFromInternalUnits(targetPara.AsDouble(), unitType);
                if (CastFamily != null)
                {
                    foreach (ElementId castId in CastFamily.GetFamilySymbolIds())
                    {
                        FamilySymbol tempSymbol = doc.GetElement(castId) as FamilySymbol;
                        //if (targetPara.AsValueString() == "50 mm")
                        if (covertUnit >= 50 && covertUnit < 65)
                        {
                            if (tempSymbol.Name == "80mm")
                            {
                                targetFamilySymbol = tempSymbol;
                            }
                        }
                        //else if (targetPara.AsValueString() == "65 mm")
                        else if (covertUnit >= 65 && covertUnit < 75)
                        {
                            if (tempSymbol.Name == "100mm")
                            {
                                targetFamilySymbol = tempSymbol;
                            }
                        }
                        //多出關於電管的判斷
                        //else if (targetPara.AsValueString() == "80 mm" || targetPara.AsValueString() == "82 mm" || targetPara.AsValueString() == "92 mm" || targetPara.AsValueString() == "75 mm")
                        else if (covertUnit >= 75 && covertUnit <= 95)
                        {
                            if (tempSymbol.Name == "125mm")
                            {
                                targetFamilySymbol = tempSymbol;
                            }
                        }
                        //else if (targetPara.AsValueString() == "100 mm" || targetPara.AsValueString() == "88 mm" || targetPara.AsValueString() == "104 mm")
                        else if (covertUnit >= 100 && covertUnit <= 125)
                        {
                            if (tempSymbol.Name == "150mm")
                            {
                                targetFamilySymbol = tempSymbol;
                            }
                        }
                        //else if (targetPara.AsValueString() == "125 mm")
                        else if (covertUnit > 125 && covertUnit <= 150)
                        {
                            if (tempSymbol.Name == "200mm")
                            {
                                targetFamilySymbol = tempSymbol;
                            }
                        }
                        else if (covertUnit > 150 && covertUnit <= 200)
                        {
                            if (tempSymbol.Name == "250mm")
                            {
                                targetFamilySymbol = tempSymbol;
                            }
                        }
                        else if (covertUnit > 200)
                        {
                            if (tempSymbol.Name == "300mm")
                            {
                                targetFamilySymbol = tempSymbol;
                            }
                        }
                        else
                        {
                            if (tempSymbol.Name == "80mm")
                            {
                                targetFamilySymbol = tempSymbol;
                            }
                        }
                    }
                }
                targetFamilySymbol.Activate();
                return targetFamilySymbol;
            }
        }

    }

    //過濾器區域
    public class PipeSelectionFilter : ISelectionFilter
    {
        private Document _doc;
        public PipeSelectionFilter(Document doc)
        {
            this._doc = doc;
        }
        public bool AllowElement(Element element)
        {
            Category pipe = Category.GetCategory(_doc, BuiltInCategory.OST_PipeCurves);
            Category duct = Category.GetCategory(_doc, BuiltInCategory.OST_DuctCurves);
            Category conduit = Category.GetCategory(_doc, BuiltInCategory.OST_Conduit);
            if (element.Category.Id == pipe.Id)
            {
                return true;
            }
            else if (element.Category.Id == duct.Id && element.LookupParameter("直徑") != null)
            {
                return true;
            }
            else if (element.Category.Id == conduit.Id )
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }
    //限制選擇牆時能選到的品類
    public class WallSelectionFilter : ISelectionFilter
    {
        private Document _doc;

        public WallSelectionFilter(Document doc)
        {
            this._doc = doc;
        }
        public bool AllowElement(Element element)
        {
            return true;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            var link = this._doc.GetElement(refer) as RevitLinkInstance;
            var elem = link.GetLinkDocument().GetElement(refer.LinkedElementId);
            if (elem is Wall)
            {
                return true;
            }
            return false;
        }
    }
}