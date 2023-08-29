#region Namespaces
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
#endregion

namespace CEC_Count
{
    class Method
    {
        private UIApplication uiapp;
        private UIDocument uidoc;
        private Autodesk.Revit.ApplicationServices.Application app;
        private Document doc;
        public Method(UIApplication uiapp)
        {
            this.uiapp = uiapp;
            this.uidoc = uiapp.ActiveUIDocument; ;
            this.app = uiapp.Application;
            this.doc = uidoc.Document;
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

        public LogicalOrFilter categoryFilter_MEP(List<BuiltInCategory> builts)
        {
            List<ElementFilter> filters = new List<ElementFilter>();
            foreach (BuiltInCategory built in builts)
            {
                ElementCategoryFilter filter = new ElementCategoryFilter(built);
                filters.Add(filter);
            }
            LogicalOrFilter categoryFilter = new LogicalOrFilter(filters);
            return categoryFilter;
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
        public void loadSharedParmeter(List<BuiltInCategory> builts, List<string> sharedParaNames)
        {
            string checkString = "";
            foreach (string st in sharedParaNames)
            {
                CategorySet catSet = app.Create.NewCategorySet();
                BindingMap bm = doc.ParameterBindings;
                DefinitionBindingMapIterator itor = bm.ForwardIterator();
                itor.Reset();
                Definition d = null;
                ElementBinding elemBind = null;
                //如果現在的專案中已經載入該參數欄位，則不需重新載入
                while (itor.MoveNext())
                {
                    d = itor.Key;
                    if (d.Name == st)
                    {
                        elemBind = (ElementBinding)itor.Current;
                        catSet = elemBind.Categories;
                        break;
                    }
                }
                foreach (BuiltInCategory builtCate in builts)
                {
                    Category tempCate = Category.GetCategory(doc, builtCate);
                    //defaultCateList.Add(tempCate);
                    if (!catSet.Contains(tempCate))
                    {
                        catSet.Insert(tempCate);
                    }
                }
                //如果該共用參數已經載入成為專案參數，重新加入binding
                if (d.Name == st && catSet.Size > 0)
                {

                    using (Transaction tx = new Transaction(doc, "Add Binding"))
                    {
                        tx.Start("重新調整共用參數");
                        InstanceBinding ib = doc.Application.Create.NewInstanceBinding(catSet);
                        bool result = doc.ParameterBindings.ReInsert(d, ib, BuiltInParameterGroup.PG_SEGMENTS_FITTINGS);
                        tx.Commit();
                    }
                }
                //如果該專案參數還沒被載入，則載入之
                else if (d.Name != st)
                {
                    checkString += $"專案尚未載入「 {st}」 參數，將自動載入\n";
                    var infoPath = @"Dropbox\info.json";
                    var jsonPath = Path.Combine(Environment.GetEnvironmentVariable("LocalAppData"), infoPath);
                    if (!File.Exists(jsonPath)) jsonPath = Path.Combine(Environment.GetEnvironmentVariable("AppData"), infoPath);
                    if (!File.Exists(jsonPath)) throw new Exception("請安裝並登入Dropbox桌面應用程式!");
                    var dropboxPath = File.ReadAllText(jsonPath).Split('\"')[5];
                    var spFilePath = dropboxPath + @"\BIM-Share\BIM共用參數.txt";
                    app.SharedParametersFilename = spFilePath;
                    DefinitionFile spFile = app.OpenSharedParameterFile();
                    ExternalDefinition targetDefinition = null;
                    foreach (DefinitionGroup dG in spFile.Groups)
                    {
                        if (dG.Name == "機電_共同")
                        {
                            foreach (ExternalDefinition def in dG.Definitions)
                            {
                                if (def.Name == st) targetDefinition = def;
                            }
                        }
                    }
                    //在此之前要建立一個審核該參數是否已經被載入的機制，如果已被載入則不載入
                    if (targetDefinition != null)
                    {
                        using (Transaction trans = new Transaction(doc))
                        {
                            trans.Start("載入共用參數");
                            InstanceBinding newIB = app.Create.NewInstanceBinding(catSet);
                            doc.ParameterBindings.Insert(targetDefinition, newIB, BuiltInParameterGroup.PG_SEGMENTS_FITTINGS);
                            trans.Commit();
                        }
                    }
                    else if (targetDefinition == null)
                    {
                        MessageBox.Show($"共用參數中沒有找到 {st} 參數");
                    }
                }
            }
        }
        public void getDocFromRevitLinkInst(CountingUI ui, Document doc)
        {
            List<LinkDoc> linkedDocs = new List<LinkDoc>();
            LinkDoc localDoc = new LinkDoc() { Name = doc.Title + "(本機)", Document = doc };
            //localDoc.linkedInstList = new List<RevitLinkInstance>();
            localDoc.Trans = Transform.Identity;
            linkedDocs.Add(localDoc);
            FilteredElementCollector linkInstCollector = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));
            foreach (Element e in linkInstCollector)
            {
                RevitLinkInstance linkInst = e as RevitLinkInstance;
                if (linkInst != null)
                {
                    Document linkedDoc = linkInst.GetLinkDocument();
                    if (linkedDoc == null) continue;
                    LinkDoc tempDoc = new LinkDoc { Name = linkedDoc.Title, Document = linkedDoc };
                    tempDoc.Trans = linkInst.GetTotalTransform();
                    //tempDoc.linkedInstList = new List<RevitLinkInstance>();
                    //tempDoc.linkedInstList.Add(linkInst);
                    linkedDocs.Add(tempDoc);
                }
            }
            ui.rvtLinkInstCombo.ItemsSource = linkedDocs;
        }
        //public void getTargetCategory(CountingUI ui, bool isMEP, List<BuiltInCategory> builts)
        public List<CustomCate> getTargetCategory(CountingUI ui, bool isMEP, List<BuiltInCategory> builts)
        //public ObservableCollection<CustomCate> getTargetCategory(CountingUI ui, bool isMEP, List<BuiltInCategory> builts)
        {
            Categories defaultCates = doc.Settings.Categories;
            //ObservableCollection<CustomCate> targetCates = new ObservableCollection<CustomCate>();
            List<CustomCate> targetCates = new List<CustomCate>();
            foreach (BuiltInCategory built in builts)
            {
                Category tempCate = Category.GetCategory(doc, built);
                CustomCate cusCat = new CustomCate()
                {
                    Name = tempCate.Name,
                    Id = tempCate.Id,
                    Cate = tempCate,
                    Selected = false,
                    BuiltCate = built
                };
                targetCates.Add(cusCat);
            }
            if (isMEP == true) ui.mepCateList.ItemsSource = targetCates;
            else ui.civilCateList.ItemsSource = targetCates;
            return targetCates;
        }
        public List<BuiltInCategory> getBuiltinCatesFromCusCate(List<CustomCate> cusCateList)
        {
            List<BuiltInCategory> builts = new List<BuiltInCategory>();
            foreach (CustomCate cusCate in cusCateList)
            {
                BuiltInCategory builtInCate = cusCate.BuiltCate;
                builts.Add(builtInCate);
            }
            return builts;
        }
        //public List<Element> getMassFromLinkDoc(RevitLinkInstance LinkedInst)
        public Solid getSolidFromActiveView(Document doc, Autodesk.Revit.DB.View view)
        {
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

        public List<Element> getMassFromLinkDoc(Document linkDoc, Transform transForm, Solid solid)
        {
            Transform inverseTrans = transForm.Inverse;
            Solid newSolid = SolidUtils.CreateTransformed(solid, inverseTrans);
            ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(newSolid);
            BoundingBoxXYZ bBox = newSolid.GetBoundingBox();
            XYZ solidCenter = newSolid.ComputeCentroid();
            Transform newTrans = Transform.Identity;
            newTrans.Origin = solidCenter;
            Outline outline = new Outline(newTrans.OfPoint( bBox.Min), newTrans.OfPoint( bBox.Max));
            BoundingBoxIntersectsFilter bBoxFilter = new BoundingBoxIntersectsFilter(outline);
            FilteredElementCollector massCollector = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Mass)
                .WherePasses(bBoxFilter)/*.WherePasses(solidFilter)*//*.WhereElementIsNotElementType()*/;
            List<Element> massList = massCollector.ToList();
            return massList;
        }
        public List<Element> getMassFromLinkDoc(Document linkDoc, Transform transForm)
        {
            FilteredElementCollector massCollector = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Mass).WhereElementIsNotElementType();
            List<Element> massList = massCollector.ToList();
            return massList;
        }
        //蒐集Mass並進行碰撞
        public void countByMass(List<CustomCate> customCates, Element mass, Transform transform)
        //public List<Element> countByMass(List<CustomCate> customCates, Element mass, Transform transform)
        {
            List<Element> targetList = null;
            try
            {
                //Transform transform = linkedInst.GetTotalTransform();
                //Document linkDoc = linkedInst.GetLinkDocument();
                //FilteredElementCollector massCollect_Link = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Mass).WhereElementIsNotElementType();
                List<BuiltInCategory> targetCateLst = getBuiltinCatesFromCusCate(customCates);
                LogicalOrFilter orFilter = categoryFilter_MEP(targetCateLst);

                //載入共用參數
                List<string> paraNames = new List<string>() { "MEP用途", "MEP區域" };

                foreach (string st in paraNames)
                {
                    if (!checkPara(mass, st))
                    {
                        MessageBox.Show($"請確認模型中的量體是否存在「{st}」參數");
                        break;
                    }
                }
                string MEPUtility = mass.LookupParameter("MEP用途").AsString();
                string MEPRegion = mass.LookupParameter("MEP區域").AsString();
                loadSharedParmeter(targetCateLst, paraNames);
                BoundingBoxXYZ massBounding = mass.get_BoundingBox(null);
                Outline massOutline = new Outline(transform.OfPoint(massBounding.Min), transform.OfPoint(massBounding.Max));
                BoundingBoxIntersectsFilter boxIntersectFilter = new BoundingBoxIntersectsFilter(massOutline);
                Solid castSolid = singleSolidFromElement(mass);
                if (castSolid != null)
                {
                    castSolid = SolidUtils.CreateTransformed(castSolid, transform);
                    ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(castSolid);
                    FilteredElementCollector MepCollector = new FilteredElementCollector(doc).WherePasses(orFilter).WhereElementIsNotElementType();
                    //再創建一個Collector來比對沒有被量體干涉到的元件
                    MepCollector.WherePasses(boxIntersectFilter).WherePasses(solidFilter);
                    if (MepCollector.Count() > 0)
                    {
                        targetList = MepCollector.ToList();
                        using (Transaction trans = new Transaction(doc))
                        {
                            trans.Start("寫入分區參數");
                            foreach (Element ee in MepCollector)
                            {
                                //Parameter targetPara = ee.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                                Parameter targetPara = ee.LookupParameter("MEP用途");
                                Parameter targetPara2 = ee.LookupParameter("MEP區域");
                                if (targetPara != null && targetPara2 != null)
                                {
                                    targetPara.Set(MEPUtility);
                                    targetPara2.Set(MEPRegion);
                                }
                            }
                            trans.Commit();
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show($"量體干涉發生錯誤");
            }
            //usefulElements.Union(targetList);
            //return targetList;
        }

        //public void removeUnuseElementPara(List<CustomCate>customCates,List<Element>usefulElement)
        public void removeUnuseElementPara(List<CustomCate> customCates)
        {
            List<string> paraNames = new List<string>() { "MEP用途", "MEP區域" };
            List<BuiltInCategory> targetCateLst = getBuiltinCatesFromCusCate(customCates);
            LogicalOrFilter orFilter = categoryFilter_MEP(targetCateLst);
            FilteredElementCollector roughCollector = new FilteredElementCollector(doc).WherePasses(orFilter).WhereElementIsNotElementType();
            foreach (Element e in roughCollector)
            {
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("清除未干涉參數");
                    Parameter targetPara = e.LookupParameter(paraNames[0]);
                    Parameter targetPara2 = e.LookupParameter(paraNames[1]);
                    if (targetPara != null && targetPara2 != null)
                    {
                        targetPara.Set("");
                        targetPara2.Set("");
                    }
                    trans.Commit();
                }
            }

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


    }
    public class LinkDoc
    {
        public string Name { get; set; }
        public Document Document { get; set; }
        public Transform Trans { get; set; }

        //public bool Selected { get; set; }
        //public List<RevitLinkInstance> linkedInstList { get; set; }
    }
    //public class CustomCate
    public class CustomCate /*: INotifyPropertyChanged*/
    {
        public string Name { get; set; }
        public ElementId Id { get; set; }
        public Category Cate { get; set; }
        public BuiltInCategory BuiltCate { get; set; }
        public bool Selected { get; set; }
        //public bool Selected
        //{
        //    get => Selected;
        //    set
        //    {
        //        if (value == Selected) return;
        //        Selected = value;
        //        OnPropertyChanged();
        //    }
        //}

        //public event PropertyChangedEventHandler PropertyChanged;
        //protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        //{
        //    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        //}
    }

}
