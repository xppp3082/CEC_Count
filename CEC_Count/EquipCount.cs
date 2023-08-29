#region Namespaces
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
using System.Text;
using System.IO;
using System.Threading;
using System.Windows.Threading;
#endregion

namespace CEC_Count
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class EquipCount : IExternalCommand
    {
        private Thread _uiThread;
        // ModelessForm instance
        private CountingUI _mMyForm;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Method m = new Method(uiapp);
            List<BuiltInCategory> builts = new List<BuiltInCategory>() {
                BuiltInCategory.OST_PipeCurves, 
                BuiltInCategory.OST_MechanicalEquipment, 
                BuiltInCategory.OST_Sprinklers ,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_FireAlarmDevices
            };
            //using (TransactionGroup transGroup = new TransactionGroup(doc))
            //{
            //    transGroup.Start("共用參數調整");
            //    List<string> paraNames = new List<string>() { "MEP用途", "MEP區域" };
            //    m.loadSharedParmeter(builts, paraNames);
            //    MessageBox.Show("共用參數調整成功");
            //    transGroup.Assimilate();
            //}

            //using (Transaction tt = new Transaction(doc))
            //{
            //    tt.Start("共用參數調整");


            //    tt.Commit();
            //}

            //#region
            ////試著用本機視圖去抓外參中的量體
            //FilteredElementCollector linkInstCollector = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));
            //Document targetDoc = null;
            //RevitLinkInstance targetLinkInst = null;
            //foreach (Element e in linkInstCollector)
            //{
            //    RevitLinkInstance linkInst = e as RevitLinkInstance;
            //    if (linkInst != null)
            //    {
            //        Document linkedDoc = linkInst.GetLinkDocument();
            //        if (linkedDoc == null) continue;
            //        FilteredElementCollector tempColl = new FilteredElementCollector(linkedDoc).OfCategory(BuiltInCategory.OST_Mass).WhereElementIsNotElementType();
            //        if (tempColl.Count() > 0)
            //        {
            //            targetDoc = linkedDoc;
            //            targetLinkInst = linkInst;
            //            break;
            //        }
            //    }
            //}

            //////複製可以嘗試複製視圖，但感覺不是好方法，改用BouningBoxUV去進行過濾
            ////VisibleInViewFilter viewFilter = new VisibleInViewFilter(doc, doc.ActiveView.Id, false);
            //Transform trans = targetLinkInst.GetTotalTransform();
            //Autodesk.Revit.DB.View av = doc.ActiveView;

            ////ViewPlan vp = ((ViewPlan)(doc.ActiveView));
            ////if (vp == null) MessageBox.Show("請在平面視圖中使用此功能");
            //BoundingBoxXYZ bounding = av.get_BoundingBox(null);
            ////Outline outline = new Outline(trans.OfPoint(bounding.Min), trans.OfPoint(bounding.Max));
            //Outline outline = new Outline(bounding.Min, bounding.Max);
            //BoundingBoxIntersectsFilter boxIntersectFilter = new BoundingBoxIntersectsFilter(outline);
            //using (Transaction ttrans = new Transaction(doc))
            //{
            //    ttrans.Start("試圖量體創建測試");
            //    Solid tempSolid = getSolidFromBBox(doc.ActiveView);
            //    ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(tempSolid);
            //    BoundingBoxXYZ bBox = tempSolid.GetBoundingBox();
            //    Outline outline1 = new Outline(bBox.Min, bBox.Max);
            //    BoundingBoxIntersectsFilter bBoxIntersectFilter = new BoundingBoxIntersectsFilter(outline1);
            //    FilteredElementCollector coll = new FilteredElementCollector(targetDoc).OfCategory(BuiltInCategory.OST_Mass).WherePasses(bBoxIntersectFilter).WhereElementIsNotElementType();
            //    MessageBox.Show(coll.Count().ToString());       
            //    DirectShape ds = createSolidFromBBox(av);
            //    MessageBox.Show(ds.Id.ToString());

            //    ttrans.Commit();
            //}
            ////PlanViewRange PVR = vp.GetViewRange();
            ////double CutOffset = PVR.GetOffset(PlanViewPlane.CutPlane)*2;
            //////MessageBox.Show((CutOffset * 30.48).ToString());
            ////ViewCropRegionShapeManager CR = vp.GetCropRegionShapeManager();
            ////IList<CurveLoop> Crops = CR.GetCropShape();
            ////MessageBox.Show(Crops.Count().ToString());
            ////Solid VirtualSolid = GeometryCreationUtilities.CreateExtrusionGeometry(new CurveLoop[] { Crops.First() }, XYZ.BasisZ, CutOffset);
            ////Solid VirtualLinkSolid = SolidUtils.CreateTransformed(VirtualSolid, targetLinkInst.GetTotalTransform().Inverse);
            ////ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(VirtualLinkSolid);

            ////FilteredElementCollector massColl = new FilteredElementCollector(targetDoc).OfCategory(BuiltInCategory.OST_Mass).WherePasses(boxIntersectFilter)./*OwnedByView(doc.ActiveView.Id).*/WhereElementIsNotElementType();
            ////MessageBox.Show(massColl.Count().ToString());
            //#endregion

            this.ShowForm(commandData);
            //this.ShowFormSeparateThread(commandData);
            return Result.Succeeded;
        }
        public DirectShape createSolidFromBBox(Autodesk.Revit.DB.View view)
        {
            Document doc = view.Document;

            BoundingBoxXYZ inputBb = null;
            double cutPlaneHeight = 0.0;
            XYZ pt0 = null;
            XYZ pt1 = null;
            XYZ pt2 = null;
            XYZ pt3 = null;
            Solid preTransformBox = null;
            if (view.ViewType == ViewType.FloorPlan)
            {
                inputBb = view.get_BoundingBox(null);
                Autodesk.Revit.DB.Plane planePlanView = view.SketchPlane.GetPlane();
                Autodesk.Revit.DB.PlanViewRange viewRange = (view as Autodesk.Revit.DB.ViewPlan).GetViewRange();
                cutPlaneHeight = viewRange.GetOffset(Autodesk.Revit.DB.PlanViewPlane.CutPlane);
                //XYZ pt0 = inputBb.Min;
                //XYZ pt1 = new XYZ(inputBb.Max.X, inputBb.Min.Y, inputBb.Min.Z);
                //XYZ pt2 = new XYZ(inputBb.Max.X, inputBb.Max.Y, inputBb.Min.Z);
                //XYZ pt3 = new XYZ(inputBb.Min.X, inputBb.Max.Y, inputBb.Min.Z);
                double level = view.GenLevel.ProjectElevation;
                pt0 = new XYZ(inputBb.Min.X, inputBb.Min.Y, level);
                pt1 = new XYZ(inputBb.Max.X, inputBb.Min.Y, level);
                pt2 = new XYZ(inputBb.Max.X, inputBb.Max.Y, level);
                pt3 = new XYZ(inputBb.Min.X, inputBb.Max.Y, level);

                Line edge0 = Line.CreateBound(pt0, pt1);
                Line edge1 = Line.CreateBound(pt1, pt2);
                Line edge2 = Line.CreateBound(pt2, pt3);
                Line edge3 = Line.CreateBound(pt3, pt0);
                List<Curve> edges = new List<Curve>();
                edges.Add(edge0);
                edges.Add(edge1);
                edges.Add(edge2);
                edges.Add(edge3);
                CurveLoop baseLoop = CurveLoop.Create(edges);
                List<CurveLoop> loopList = new List<CurveLoop>();
                loopList.Add(baseLoop);
                preTransformBox = GeometryCreationUtilities.CreateExtrusionGeometry(loopList, XYZ.BasisZ, cutPlaneHeight);
                //Solid 
                double solidheight = inputBb.Max.Z - inputBb.Min.Z;
            }
            else if (view.ViewType == ViewType.ThreeD)
            {
                View3D view3D = (View3D)view;
                //inputBb = view3D.GetSectionBox();
                inputBb = view.CropBox;
                if (inputBb == null) MessageBox.Show("請確認剖面框是否開啟");
                pt0 = inputBb.Min;
                pt1 = new XYZ(inputBb.Max.X, inputBb.Min.Y, inputBb.Min.Z);
                pt2 = new XYZ(inputBb.Max.X, inputBb.Max.Y, inputBb.Min.Z);
                pt3 = new XYZ(inputBb.Min.X, inputBb.Max.Y, inputBb.Min.Z);
                Line edge0 = Line.CreateBound(pt0, pt1);
                Line edge1 = Line.CreateBound(pt1, pt2);
                Line edge2 = Line.CreateBound(pt2, pt3);
                Line edge3 = Line.CreateBound(pt3, pt0);
                List<Curve> edges = new List<Curve>();
                edges.Add(edge0);
                edges.Add(edge1);
                edges.Add(edge2);
                edges.Add(edge3);
                CurveLoop baseLoop = CurveLoop.Create(edges);
                List<CurveLoop> loopList = new List<CurveLoop>();
                loopList.Add(baseLoop);
                double solidheight = inputBb.Max.Z - inputBb.Min.Z;
                preTransformBox = GeometryCreationUtilities.CreateExtrusionGeometry(loopList, XYZ.BasisZ, solidheight);
            }
            // Put this inside a transaction!


            DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
            ds.ApplicationId = "Test";
            ds.ApplicationDataId = "testBox";
            List<GeometryObject> GeoList = new List<GeometryObject>();
            GeoList.Add(preTransformBox); // <-- the solid created for the intersection can be used here
            ds.SetShape(GeoList);
            ds.SetName("ID_testBox");

            return ds;
        }

        public void ShowForm(ExternalCommandData commandData)
        {
            UIApplication uiapp = commandData.Application;
            // If we do not have a dialog yet, create and show it
            Document doc = uiapp.ActiveUIDocument.Document;
            if (_mMyForm != null && _mMyForm == null) return;
            //EXTERNAL EVENTS WITH ARGUMENTS
            EventHandlerWithStringArg evStr = new EventHandlerWithStringArg();
            EventHandlerWithWpfArg evWpf = new EventHandlerWithWpfArg();

            #region

            // The dialog becomes the owner responsible for disposing the objects given to it.
            #endregion
            _mMyForm = new CountingUI(commandData, evStr, evWpf);
            _mMyForm.Show();
        }

        public Solid getSolidFromBBox(Autodesk.Revit.DB.View view)
        {
            Document doc = view.Document;

            BoundingBoxXYZ inputBb = null;
            double cutPlaneHeight = 0.0;
            XYZ pt0 = null;
            XYZ pt1 = null;
            XYZ pt2 = null;
            XYZ pt3 = null;
            Solid preTransformBox = null;
            if (view.ViewType == ViewType.FloorPlan)
            {
                inputBb = view.get_BoundingBox(null);
                Autodesk.Revit.DB.Plane planePlanView = view.SketchPlane.GetPlane();
                Autodesk.Revit.DB.PlanViewRange viewRange = (view as Autodesk.Revit.DB.ViewPlan).GetViewRange();
                cutPlaneHeight = viewRange.GetOffset(Autodesk.Revit.DB.PlanViewPlane.CutPlane);
                //XYZ pt0 = inputBb.Min;
                //XYZ pt1 = new XYZ(inputBb.Max.X, inputBb.Min.Y, inputBb.Min.Z);
                //XYZ pt2 = new XYZ(inputBb.Max.X, inputBb.Max.Y, inputBb.Min.Z);
                //XYZ pt3 = new XYZ(inputBb.Min.X, inputBb.Max.Y, inputBb.Min.Z);
                double level = view.GenLevel.ProjectElevation;
                pt0 = new XYZ(inputBb.Min.X, inputBb.Min.Y, level);
                pt1 = new XYZ(inputBb.Max.X, inputBb.Min.Y, level);
                pt2 = new XYZ(inputBb.Max.X, inputBb.Max.Y, level);
                pt3 = new XYZ(inputBb.Min.X, inputBb.Max.Y, level);

                Line edge0 = Line.CreateBound(pt0, pt1);
                Line edge1 = Line.CreateBound(pt1, pt2);
                Line edge2 = Line.CreateBound(pt2, pt3);
                Line edge3 = Line.CreateBound(pt3, pt0);
                List<Curve> edges = new List<Curve>();
                edges.Add(edge0);
                edges.Add(edge1);
                edges.Add(edge2);
                edges.Add(edge3);
                CurveLoop baseLoop = CurveLoop.Create(edges);
                List<CurveLoop> loopList = new List<CurveLoop>();
                loopList.Add(baseLoop);
                preTransformBox = GeometryCreationUtilities.CreateExtrusionGeometry(loopList, XYZ.BasisZ, cutPlaneHeight);
                //Solid 
                double solidheight = inputBb.Max.Z - inputBb.Min.Z;
            }
            else if (view.ViewType == ViewType.ThreeD)
            {
                View3D view3D = (View3D)view;
                inputBb = view3D.GetSectionBox();
                if (inputBb == null) MessageBox.Show("請確認剖面框是否開啟");
                pt0 = inputBb.Min;
                pt1 = new XYZ(inputBb.Max.X, inputBb.Min.Y, inputBb.Min.Z);
                pt2 = new XYZ(inputBb.Max.X, inputBb.Max.Y, inputBb.Min.Z);
                pt3 = new XYZ(inputBb.Min.X, inputBb.Max.Y, inputBb.Min.Z);
                Line edge0 = Line.CreateBound(pt0, pt1);
                Line edge1 = Line.CreateBound(pt1, pt2);
                Line edge2 = Line.CreateBound(pt2, pt3);
                Line edge3 = Line.CreateBound(pt3, pt0);
                List<Curve> edges = new List<Curve>();
                edges.Add(edge0);
                edges.Add(edge1);
                edges.Add(edge2);
                edges.Add(edge3);
                CurveLoop baseLoop = CurveLoop.Create(edges);
                List<CurveLoop> loopList = new List<CurveLoop>();
                loopList.Add(baseLoop);
                double solidheight = inputBb.Max.Z - inputBb.Min.Z;
                preTransformBox = GeometryCreationUtilities.CreateExtrusionGeometry(loopList, XYZ.BasisZ, solidheight);
            }
            return preTransformBox;
        }
        public void ShowFormSeparateThread(ExternalCommandData commandData)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            // If we do not have a thread started or has been terminated start a new one
            if (!(_uiThread is null) && _uiThread.IsAlive) return;
            //EXTERNAL EVENTS WITH ARGUMENTS
            EventHandlerWithStringArg evStr = new EventHandlerWithStringArg();
            EventHandlerWithWpfArg evWpf = new EventHandlerWithWpfArg();


            //新增執行序
            _uiThread = new Thread(() =>
            {
                SynchronizationContext.SetSynchronizationContext(
                    new DispatcherSynchronizationContext(
                        Dispatcher.CurrentDispatcher));
                // The dialog becomes the owner responsible for disposing the objects given to it.

                _mMyForm = new CountingUI(commandData, evStr, evWpf);
                _mMyForm.Closed += (s, e) => Dispatcher.CurrentDispatcher.InvokeShutdown();
                _mMyForm.Show();
                Dispatcher.Run();
            });

            _uiThread.SetApartmentState(ApartmentState.STA);
            _uiThread.IsBackground = true;
            _uiThread.Start();
        }
    }
}
