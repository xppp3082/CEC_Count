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
            //string testOutput = "";

            //Method m = new Method(uiapp);
            ////製作過濾器，此功能後續可供選擇，要更新參數的品類有哪些
            //List<string> paraNames = new List<string>() { "MEP用途", "MEP區域" };
            ////BuiltInCategory[] builts =
            //List<BuiltInCategory>builts=new List<BuiltInCategory>()
            //{
            //BuiltInCategory.OST_PipeCurves,
            //BuiltInCategory.OST_PipeAccessory,
            //BuiltInCategory.OST_PipeFitting,
            //BuiltInCategory.OST_Conduit,
            //BuiltInCategory.OST_ConduitFitting,
            //BuiltInCategory.OST_DuctCurves,
            //BuiltInCategory.OST_MechanicalEquipment,
            //BuiltInCategory.OST_FurnitureSystems,
            //BuiltInCategory.OST_Sprinklers,
            //BuiltInCategory.OST_FireAlarmDevices
            //};
            //LogicalOrFilter orFilter = m.categoryFilter_MEP(builts);
            //m.loadSharedParmeter(builts,paraNames);//載入共用參數
            //this.ShowForm(commandData);
            this.ShowFormSeparateThread(commandData);
            ////CountingUI ui = new CountingUI(commandData);
            ////ui.ShowDialog();

            ////選擇具有量體的外參模型，並將其量體中boundingbox取出來
            //FilteredElementCollector rvtLinkCollector = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));

            //#region 取得量體文件-->取法要在改變
            //RevitLinkInstance linkInst = rvtLinkCollector.First() as RevitLinkInstance;
            //Transform transform = linkInst.GetTotalTransform();
            //Document linkDoc = linkInst.GetLinkDocument();
            //FilteredElementCollector massCollect_link = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Mass).WhereElementIsNotElementType();
            //MessageBox.Show($"{linkDoc.Title}中，共有{massCollect_link.Count()}個實作量體");

            ////取得量體後，利用量體去和「每一個」本地端的元件進行碰撞
            //if (massCollect_link.Count() == 0) { MessageBox.Show("選取的外參檔案中無實作的量體"); }
            //foreach (Element e in massCollect_link)
            //{
            //    foreach (string st in paraNames)
            //    {
            //        if (!m.checkPara(e, st))
            //        {
            //            MessageBox.Show($"請確認模型中的量體是否存在「{st}」參數");
            //            break;
            //        }
            //    }
            //    string MEPUtility = e.LookupParameter("MEP用途").AsString();
            //    string MEPRegion = e.LookupParameter("MEP區域").AsString();

            //    BoundingBoxXYZ massBounding = e.get_BoundingBox(null);
            //    Outline massOutline = new Outline(transform.OfPoint(massBounding.Min), transform.OfPoint(massBounding.Max));
            //    BoundingBoxIntersectsFilter boxIntersectFilter = new BoundingBoxIntersectsFilter(massOutline);
            //    Solid castSolid = m.singleSolidFromElement(e);
            //    if (castSolid == null) continue;
            //    ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(castSolid);
            //    FilteredElementCollector MepCollector = new FilteredElementCollector(doc).WherePasses(orFilter).WhereElementIsNotElementType();
            //    MepCollector.WherePasses(boxIntersectFilter).WherePasses(solidFilter);
            //    using (Transaction trans = new Transaction(doc))
            //    {
            //        trans.Start("寫入分區參數");
            //        foreach (Element ee in MepCollector)
            //        {
            //            Parameter targetPara = ee.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            //            targetPara.Set(MEPRegion);
            //        }
            //        trans.Commit();

            //        Categories cates = doc.Settings.Categories;
            //        foreach (Category cat in cates)
            //        {
            //            testOutput += $"{ cat.Id}";
            //        }

            //    }
            //}
            //MessageBox.Show(testOutput);
            //#endregion
            return Result.Succeeded;
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

        public void ShowFormSeparateThread(ExternalCommandData commandData)
        {
            // If we do not have a thread started or has been terminated start a new one
            if (!(_uiThread is null) && _uiThread.IsAlive) return;
            //EXTERNAL EVENTS WITH ARGUMENTS
            EventHandlerWithStringArg evStr = new EventHandlerWithStringArg();
            EventHandlerWithWpfArg evWpf = new EventHandlerWithWpfArg();


            //新增執行敘
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
