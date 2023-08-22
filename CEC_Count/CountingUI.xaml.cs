﻿using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using System.Windows.Forms;
using System.Text;
using System.IO;
using System.Threading;
using System.Collections.ObjectModel;

namespace CEC_Count
{
    /// <summary>
    /// CountingUI.xaml 的互動邏輯
    /// </summary>
    public partial class CountingUI : Window
    {
        private readonly UIApplication _uiApp;
        private readonly Autodesk.Revit.ApplicationServices.Application _app;
        private readonly UIDocument _uiDoc;
        private readonly Document _doc;

        private readonly EventHandlerWithStringArg _mExternalMethodStringArg;
        private readonly EventHandlerWithWpfArg _mExternalMethodWpfArg;
        public ObservableCollection<CustomCate> mepCusCateList;
        public ObservableCollection<CustomCate> civilCusCateList;


        public CountingUI(ExternalCommandData commandData, EventHandlerWithStringArg evExternalMethodStringArg,
            EventHandlerWithWpfArg eExternalMethodWpfArg)
        {
            _uiApp = commandData.Application;
            _app = _uiApp.Application;
            _uiDoc = commandData.Application.ActiveUIDocument;
            _doc = _uiDoc.Document;

            InitializeComponent();

            //UI初始設定
            _mExternalMethodStringArg = evExternalMethodStringArg;
            _mExternalMethodWpfArg = eExternalMethodWpfArg;
            Method m = new Method(_uiApp);
            m.getDocFromRevitLinkInst(this, _doc);
            List<BuiltInCategory> MEPcates = new List<BuiltInCategory>()
            {
                BuiltInCategory.OST_PipeCurves,//管
                BuiltInCategory.OST_PipeFitting,//管配件
                BuiltInCategory.OST_PipeAccessory,//管附件
                BuiltInCategory.OST_DuctCurves,//風管
                BuiltInCategory.OST_DuctFitting,//風管配件
                BuiltInCategory.OST_DuctAccessory,//風管附件
                BuiltInCategory.OST_Conduit,//電管
                BuiltInCategory.OST_ConduitFitting,//電管配件
                BuiltInCategory.OST_CableTray,//電纜架
                BuiltInCategory.OST_CableTrayFitting,//電纜架配件
                BuiltInCategory.OST_Sprinklers,
                BuiltInCategory.OST_FlexPipeCurves,
                BuiltInCategory.OST_FlexDuctCurves,
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_FireAlarmDevices,
                BuiltInCategory.OST_LightingDevices,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_CommunicationDevices,
                BuiltInCategory.OST_SecurityDevices,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_ElectricalEquipment
            };
            List<BuiltInCategory> Civilcates = new List<BuiltInCategory>()
            {
                BuiltInCategory.OST_Ceilings,
                BuiltInCategory.OST_Roofs,
                BuiltInCategory.OST_Columns,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralTruss,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_Stairs,
                BuiltInCategory.OST_Windows,
                BuiltInCategory.OST_Doors
            };
            m.getTargetCategory(this, true, MEPcates);
            m.getTargetCategory(this, false, Civilcates);
        }

        private void continueButton_Click(object sender, RoutedEventArgs e)
        {

            //用來執行methodWrapper裡面的方法
            _mExternalMethodWpfArg.Raise(this);
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("Cancel button was clicked.");
            Close();
            return;
        }

        private void mepCheckAll_check(object sender, RoutedEventArgs e)
        {
            bool allCheck = (mepCheckAll.IsChecked == true);
            //ObservableCollection<CustomCate> cusCateList = new ObservableCollection<CustomCate>();
            int count = 0;
            //this.mepCateList.Items
            foreach (CustomCate cusCate in this.mepCateList.ItemsSource)
            {
                cusCate.Selected = allCheck;
                count++;
                //cusCateList.Add(cusCate);
              //cusCate.Selected = true;
            }
            //this.Dispatcher.Invoke(() => this.mepCateList.DataContext = cusCateList);
            //InitializeComponent();
        }

        private void civilCheckAll_check(object sender, RoutedEventArgs e)
        {
            //bool allCheck = (civilCheckAll.IsChecked == true);
            //this.mepCheckAll.IsChecked = true;
        }
        //private void changePropertyofCusCate
    }
}