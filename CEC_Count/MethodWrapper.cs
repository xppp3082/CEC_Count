﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace CEC_Count
{
    public class EventHandlerWithStringArg : RevitEventWrapper<string>
    {
        /// <summary>
        /// The Execute override void must be present in all methods wrapped by the RevitEventWrapper.
        /// This defines what the method will do when raised externally.
        /// </summary>
        public override void Execute(UIApplication uiApp, string args)
        {
            // Do your processing here with "args"
            TaskDialog.Show("External Event", args);
        }
    }
    /// <summary>
    /// This is an example of of wrapping a method with an ExternalEventHandler using an instance of WPF
    /// as an argument. Any type of argument can be passed to the RevitEventWrapper, and therefore be used in
    /// the execution of a method which has to take place within a "Valid Revit API Context". This specific
    /// pattern can be useful for smaller applications, where it is convenient to access the WPF properties
    /// directly, but can become cumbersome in larger application architectures. At that point, it is suggested
    /// to use more "low-level" wrapping, as with the string-argument-wrapped method above.
    /// </summary>
    /// 
    public class EventHandlerWithWpfArg : RevitEventWrapper<CountingUI>
    {
        public override void Execute(UIApplication uiApp, CountingUI ui)
        {
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            List<CustomCate> cusCateLst = new List<CustomCate>();
            Method m = new Method(uiApp);
            //以勾選的品類作為計數單位
            int count = 0;

            ui.Dispatcher.Invoke(() => ui.pBar.Value = 0);
            //ui.Dispatcher.Invoke(() => ui.pBar.Maximum = count);
            int selectedINdex = ui.rvtLinkInstCombo.SelectedIndex;
            //排序以本機端檔案為第一個
            if (selectedINdex == 0)
            {
                Category massCate = Category.GetCategory(doc, BuiltInCategory.OST_Mass);
                CustomCate tempCate = new CustomCate()
                {
                    Name = massCate.Name,
                    Id = massCate.Id,
                    Cate = massCate,
                    Selected = true,
                    BuiltCate = BuiltInCategory.OST_Mass
                };
                cusCateLst.Add(tempCate);
            }
            foreach (CustomCate cate in ui.mepCateList.Items)
            {
                if (cate.Selected == true)
                {
                    count++;
                    cusCateLst.Add(cate);
                }
            }
            string countMessage = $"共選擇了 {count} 個品類，正在進行數量計算...";
            ui.Dispatcher.Invoke(() => ui.selectHint.Text = countMessage);
            ui.Dispatcher.Invoke(() => ui.selectHint.Foreground = Brushes.Black);
            LinkDoc targetDoc = ui.rvtLinkInstCombo.SelectedItem as LinkDoc;
            if (targetDoc != null)
            {
                Autodesk.Revit.DB.Transform targetTrans = targetDoc.Trans;
                //Autodesk.Revit.DB.Transform targetTrans = targetDoc.linkedInstList.First().GetTotalTransform();
                //先取得量體總數，更新進度條上限
                List<Element> massLIst = m.getMassFromLinkDoc(targetDoc.Document, targetTrans);
                ui.Dispatcher.Invoke(() => ui.pBar.Maximum = massLIst.Count());

                //針對量體總數進行干涉與進度更新
                using (TransactionGroup transGroup = new TransactionGroup(doc))
                {
                    transGroup.Start("分區數量計算");
                    foreach (Element e in massLIst)
                    {
                        m.countByMass(cusCateLst, e, targetTrans);
                        ui.Dispatcher.Invoke(() => ui.pBar.Value += 1, System.Windows.Threading.DispatcherPriority.Background);
                    }
                    //foreach(CustomCate cusCate in cusCateLst)
                    //{
                    //    ui.Dispatcher.Invoke(() =>ui.pBar.Value+=1,System.Windows.Threading.DispatcherPriority.Background);
                    //}
                    //排到最後才執行的部分
                    transGroup.Assimilate();
                }

                Task.Run(() =>
                    {
                        string completeMessage = $"計算完成!!";
                        ui.Dispatcher.Invoke(() => ui.selectHint.Text = completeMessage);
                        ui.Dispatcher.Invoke(() => ui.selectHint.Foreground = Brushes.Blue);
                    });
                ui.Activate();
            }
            else if (targetDoc == null)
            {
                Task.Run(() =>
                {
                    string errorMessage = $"計算失敗，請確認是否選擇明確的量體來源模型!!";
                    //Color errorColor = new Color(255, 0, 0);
                    ui.Dispatcher.Invoke(() => ui.selectHint.Text = errorMessage);
                    ui.Dispatcher.Invoke(() => ui.selectHint.Foreground = Brushes.Red);
                });
                ui.Activate();
            }
        }
    }
}
