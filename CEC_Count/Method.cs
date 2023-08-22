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
using System.ComponentModel;
using System.Collections.ObjectModel;
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
            CategorySet catSet = app.Create.NewCategorySet();
            //List<Category> defaultCateList = new List<Category>() { };
            foreach (BuiltInCategory builtCate in builts)
            {
                Category tempCate = Category.GetCategory(doc, builtCate);
                //defaultCateList.Add(tempCate);
                if (!catSet.Contains(tempCate))
                {
                    catSet.Insert(tempCate);
                }
            }
            foreach (string st in sharedParaNames)
            {
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
                        break;
                    }
                }
                //如果該共用參數已經載入成為專案參數，重新加入binding
                if (d.Name == st && catSet.Size > 0)
                {
                    using (Transaction tx = new Transaction(doc, "Add Binding"))
                    {
                        tx.Start();
                        InstanceBinding ib = doc.Application.Create.NewInstanceBinding(catSet);
                        bool result = doc.ParameterBindings.ReInsert(d, ib, BuiltInParameterGroup.PG_SEGMENTS_FITTINGS);
                        tx.Commit();
                    }
                }
                //如果該專案參數還沒被載入，則載入之
                else if (d.Name != st)
                {
                    checkString += $"專案尚未載入「 {st}」 參數，將自動載入\n";
                    //MessageBox.Show($"專案尚未載入「 {st}」 參數，將自動載入");
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
        public void getTargetCategory(CountingUI ui, bool isMEP, List<BuiltInCategory> builts)
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
            //return targetCates;
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
        public List<Element> getMassFromLinkDoc(Document linkDoc, Transform transForm)
        {
            //Transform transForm = LinkedInst.GetTotalTransform();
            //Document linkDoc = LinkedInst.GetLinkDocument();
            FilteredElementCollector massCollector = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Mass).WhereElementIsNotElementType();
            List<Element> massList = massCollector.ToList();
            return massList;
        }
        //蒐集Mass並進行碰撞
        //public void getMassByTagerLinkInst(RevitLinkInstance linkedInst, List<CustomCate> customCates, Element mass)
        public void countByMass(List<CustomCate> customCates, Element mass, Transform transform)
        {
            try
            {
                //Transform transform = linkedInst.GetTotalTransform();
                //Document linkDoc = linkedInst.GetLinkDocument();
                //FilteredElementCollector massCollect_Link = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Mass).WhereElementIsNotElementType();
                List<BuiltInCategory> targetCateLst = getBuiltinCatesFromCusCate(customCates);
                LogicalOrFilter orFilter = categoryFilter_MEP(targetCateLst);

                //載入共用參數
                List<string> paraNames = new List<string>() { "MEP用途", "MEP區域" };

                //if (massCollect_Link.Count() == 0) MessageBox.Show("選中的量體來源檔案中並無量體!");
                //foreach (Element e in massCollect_Link)
                //{
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
                    ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(castSolid);
                    FilteredElementCollector MepCollector = new FilteredElementCollector(doc).WherePasses(orFilter).WhereElementIsNotElementType();

                    MepCollector.WherePasses(boxIntersectFilter).WherePasses(solidFilter);
                    if (MepCollector.Count() > 0)
                    {
                        using (Transaction trans = new Transaction(doc))
                        {
                            trans.Start("寫入分區參數");
                            foreach (Element ee in MepCollector)
                            {
                                //MessageBox.Show("YA");
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
                MessageBox.Show("量體干涉失敗");
            }

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
    public class CustomCate
    {
        public string Name { get; set; }
        public ElementId Id { get; set; }
        public Category Cate { get; set; }
        public bool Selected { get; set; }
        public BuiltInCategory BuiltCate { get; set; }
    }

}
