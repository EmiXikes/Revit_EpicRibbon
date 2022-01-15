using Autodesk.Revit.UI;
using System;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Drawing;
using System.Activities.Presentation;
using System.IO;
using System.Collections.Generic;
using adWin = Autodesk.Windows;

namespace Revit_EpicRibbon
{
    public class EpicApp : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
            // throw new NotImplementedException();
        }


        public Result OnStartup(UIControlledApplication application)
        {
            // Create a custom ribbon tab
            String tabName = "Epic Tools";
            application.CreateRibbonTab(tabName);





            #region System Changer

            String dllPathSystem = @"C:\Epic\EpicToolsAddinForRevit\Libs\Revit_SystemSetter.dll";
            //String dllPathSystem = @"C:\Users\User\source\repos\_Revit\Revit_SystemSetter\bin\Debug\Revit_SystemSetter.dll";


            // EL
            PushButtonData btn_H2 = new PushButtonData("btn_H2", "EL", dllPathSystem, "Revit_SystemSetter.SetSystem_H2");
            PushButtonData btn_H201 = new PushButtonData("btn_H201", "EL EI", dllPathSystem, "Revit_SystemSetter.SetSystem_H201");
            PushButtonData btn_H202 = new PushButtonData("btn_H202", "EL EL+", dllPathSystem, "Revit_SystemSetter.SetSystem_H202");
            PushButtonData btn_H203 = new PushButtonData("btn_H203", "METAL", dllPathSystem, "Revit_SystemSetter.SetSystem_H203");

            //Voids
            PushButtonData btn_H2V1 = new PushButtonData("btn_H2V1", "SAME", dllPathSystem, "Revit_SystemSetter.SetSystem_H2V1");
            PushButtonData btn_H2V2 = new PushButtonData("btn_H2V2", "NEW", dllPathSystem, "Revit_SystemSetter.SetSystem_H2V2");
            PushButtonData btn_H2V3 = new PushButtonData("btn_H2V3", "OLD", dllPathSystem, "Revit_SystemSetter.SetSystem_H2V3");

            // EL + ESS
            PushButtonData btn_HJ201 = new PushButtonData("btn_HJ201", "EL+ESS", dllPathSystem, "Revit_SystemSetter.SetSystem_HJ201");


            // ESS
            PushButtonData btn_J2 = new PushButtonData("btn_J2", "ESS", dllPathSystem, "Revit_SystemSetter.SetSystem_J2");
            PushButtonData btn_J201 = new PushButtonData("btn_J201", "ESS EI" ,dllPathSystem, "Revit_SystemSetter.SetSystem_J201");
            PushButtonData btn_J201a = new PushButtonData("btn_J201a", "ESS EI", dllPathSystem, "Revit_SystemSetter.SetSystem_J201");

            // Lights
            PushButtonData btn_H501 = new PushButtonData("btn_H501", "G-LIGHT" ,dllPathSystem, "Revit_SystemSetter.SetSystem_H501");
            PushButtonData btn_H502 = new PushButtonData("btn_H502", "E-LIGHT" ,dllPathSystem, "Revit_SystemSetter.SetSystem_H502");

            //Lightning
            PushButtonData btn_H701 = new PushButtonData("btn_H701", "LGTNG" ,dllPathSystem, "Revit_SystemSetter.SetSystem_H701");
            PushButtonData btn_H702 = new PushButtonData("btn_H702", "LGTNG-EX", dllPathSystem, "Revit_SystemSetter.SetSystem_H702");



            // Create a ribbon panel
            RibbonPanel ribbonPanelSystem = application.CreateRibbonPanel(tabName, "System Changer");

            // Add the buttons to the panel
            //List<RibbonItem> projectButtons = new List<RibbonItem>();
            //projectButtons.AddRange(m_projectPanel.AddStackedItems(button1, button2));

            PushButton btn;

            // EL buttons
            btn = ribbonPanelSystem.AddItem(btn_H2) as PushButton;
            btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.SysChanger_EL_32.png");

            ribbonPanelSystem.AddSeparator();

            var stEL = ribbonPanelSystem.AddStackedItems(btn_H201, btn_H202, btn_H203);
            List<ImageSource> imageSourcesEL = new List<ImageSource>() 
            {
                PngImageSource("Revit_EpicRibbon.ICONS.SysChanger_EL_EI_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.SysChanger_EL_EIPlus_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.SysChanger_EL_M_16.png")
            };
            for(int i = 0; i < stEL.Count; i++)
            {
                btn = stEL[i] as PushButton;
                btn.Image = imageSourcesEL[i];
            }

            ribbonPanelSystem.AddSeparator();

            // ESS buttons

            btn = ribbonPanelSystem.AddItem(btn_J2) as PushButton;
            btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.SysChanger_ESS_32.png");

            ribbonPanelSystem.AddSeparator();

            var stESS = ribbonPanelSystem.AddStackedItems(btn_J201, btn_J201a);
            List<ImageSource> imageSourcesESS = new List<ImageSource>()
            {
                PngImageSource("Revit_EpicRibbon.ICONS.SysChanger_ESS_EI_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.SysChanger_ESS_EI_16.png"),
            };

            for (int i = 0; i < stESS.Count; i++)
            {
                btn = stESS[i] as PushButton;
                btn.Image = imageSourcesESS[i];
            }

            ribbonPanelSystem.AddSeparator();

            // Light Buttons
            var stL = ribbonPanelSystem.AddStackedItems(btn_H501, btn_H502);
            List<ImageSource> imageSourcesLight = new List<ImageSource>()
            {
                PngImageSource("Revit_EpicRibbon.ICONS.SysChanger_LG_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.SysChanger_LE_16.png"),
            };
            for (int i = 0; i < stL.Count; i++)
            {
                btn = stL[i] as PushButton;
                btn.Image = imageSourcesLight[i];
            }

            ribbonPanelSystem.AddSeparator();

            // Lightning
            ribbonPanelSystem.AddStackedItems(btn_H701, btn_H702);
            ribbonPanelSystem.AddSeparator();
            ribbonPanelSystem.AddStackedItems(btn_H2V1, btn_H2V2, btn_H2V3);

            #endregion
            #region Void Enummerator

            String dllPathVoids = @"C:\Epic\EpicToolsAddinForRevit\Libs\Revit_VoidEnumerator.dll";
            //String dllPathVoids = @"C:\Users\User\source\repos\_Revit\Revit_VoidEnumerator\bin\Debug\Revit_VoidEnumerator.dll";

            PushButtonData btn_VoidEnum = new PushButtonData("btn_VoidEnum", "Void ENUM", dllPathVoids, "Revit_VoidEnumerator.VoidEnum");

            // Create a ribbon panel
            RibbonPanel ribbonPanelVoids = application.CreateRibbonPanel(tabName, "Void Enum");

            ribbonPanelVoids.AddItem(btn_VoidEnum);


            #endregion

            #region LumiSnap

            String dllPathLumiSnap = @"C:\Users\xikes\source\repos\LumiSnap\bin\Debug\LumiSnap.dll";

            PushButtonData btn_LumiSnap = new PushButtonData("btn_lumisnap", "Lumi Snap", dllPathLumiSnap, "EpicLumiSnap.LumiSnap");
            PushButtonData btn_LumiSnapSettings = new PushButtonData("btn_lumisnapsettings", "Settings", dllPathLumiSnap, "EpicLumiSnap.LumiSnapSettings");

            // Create a ribbon panel
            RibbonPanel ribbonPanelLumiSnap = application.CreateRibbonPanel(tabName, "Lumi Snap");


            btn = ribbonPanelLumiSnap.AddItem(btn_LumiSnap) as PushButton;
            btn.Image = PngImageSource("Revit_EpicRibbon.ICONS.LumiSnap_16.png");
            btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.LumiSnap_32.png");

            ribbonPanelLumiSnap.AddSeparator();

            btn = ribbonPanelLumiSnap.AddItem(btn_LumiSnapSettings) as PushButton;
            btn.Image = PngImageSource("Revit_EpicRibbon.ICONS.SettingsGear_16.png");
            btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.SettingsGear_32.png");

            #endregion

            //String ModifyTabName = "";
            //adWin.RibbonControl MyRibbon;
            //MyRibbon = Autodesk.Windows.ComponentManager.Ribbon;
            //foreach (var t in MyRibbon.Tabs)
            //{
            //    if (t.AutomationName.Contains("Modify"))
            //    {
            //        t.
            //        ModifyTabName = t.AutomationName;
            //    }
            //}


            return Result.Succeeded;
        }


        private System.Windows.Media.ImageSource BmpImageSource(string embeddedPath)
        {
            Stream stream = this.GetType().Assembly.GetManifestResourceStream(embeddedPath);
            var decoder = new System.Windows.Media.Imaging.BmpBitmapDecoder(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }

        private System.Windows.Media.ImageSource PngImageSource(string embeddedPath)
        {
            Stream stream = this.GetType().Assembly.GetManifestResourceStream(embeddedPath);
            var decoder = new System.Windows.Media.Imaging.PngBitmapDecoder(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }

    }
}
