using Autodesk.Revit.UI;
using System;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Drawing;
using System.Activities.Presentation;
using System.IO;
using System.Collections.Generic;
using System.Reflection;
using Autodesk.Revit.DB.Events;
using adWin = Autodesk.Windows;
using System.Diagnostics;

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

            var BaseLocation = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            BaseLocation = System.IO.Path.Combine(BaseLocation, "..", "Apps");

            #region System Changer

            String dllPathSystem;
            //String dllPathSystem = @"C:\Users\User\source\repos\_Revit\Revit_SystemSetter\bin\Debug\Revit_SystemSetter.dll";
            //dllPathSystem = @"C:\Epic\EpicToolsAddinForRevit\Libs\Revit_SystemSetter.dll";
            dllPathSystem = System.IO.Path.Combine(BaseLocation, "SystemChanger", "Revit_SystemSetter.dll");

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
            PushButtonData btn_H703 = new PushButtonData("btn_H703", "EARTH", dllPathSystem, "Revit_SystemSetter.SetSystem_H703");



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
            ribbonPanelSystem.AddStackedItems(btn_H701, btn_H702, btn_H703);
            ribbonPanelSystem.AddSeparator();
            ribbonPanelSystem.AddStackedItems(btn_H2V1, btn_H2V2, btn_H2V3);

            #endregion

            #region Void Enummerator

            String dllPathVoids;
            //String dllPathVoids = @"C:\Users\User\source\repos\_Revit\Revit_VoidEnumerator\bin\Debug\Revit_VoidEnumerator.dll";
            //dllPathVoids = @"C:\Epic\EpicToolsAddinForRevit\Libs\Revit_VoidEnumerator.dll";
            dllPathVoids = System.IO.Path.Combine(BaseLocation, "VoidEnum", "Revit_VoidEnumerator.dll");

            PushButtonData btn_VoidEnum = new PushButtonData("btn_VoidEnum", "Void ENUM", dllPathVoids, "Revit_VoidEnumerator.VoidEnum");

            // Create a ribbon panel
            RibbonPanel ribbonPanelVoids = application.CreateRibbonPanel(tabName, "Void Enum");

            ribbonPanelVoids.AddItem(btn_VoidEnum);


            #endregion

            #region EpicLumi

            String dllPathLumi;
            //dllPathLumi = @"C:\Users\User\source\repos\_Revit\Revit_EpicLumImporter\bin\Debug\EpicLumi.dll";
            dllPathLumi = System.IO.Path.Combine(BaseLocation, "Lumi", "EpicLumi.dll");

            PushButtonData btn_LumiMain = new PushButtonData("btn_lumiMain", "Lumi", dllPathLumi, "EpicLumi.Lumi");
            PushButtonData btn_LumiInfoBlocks = new PushButtonData("btn_lumiInfoBlocks", "nfoImp", dllPathLumi, "EpicLumi.LumiInfoImport");
            PushButtonData btn_LumiAnno = new PushButtonData("btn_lumiAnno", "Tag", dllPathLumi, "EpicLumi.LumiTag");
            PushButtonData btn_LumiAnnoUpdate = new PushButtonData("btn_lumiAnnoUpdate", "Update", dllPathLumi, "EpicLumi.LumiTagUpdate");

            // Create a ribbon panel
            RibbonPanel ribbonPanelLumi = application.CreateRibbonPanel(tabName, "Epic Lumi");

            btn = ribbonPanelLumi.AddItem(btn_LumiMain) as PushButton;
            btn.Image = PngImageSource("Revit_EpicRibbon.ICONS.LumiSnap_16.png");
            btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.LumiMain.png");

            btn = ribbonPanelLumi.AddItem(btn_LumiInfoBlocks) as PushButton;
            btn.Image = PngImageSource("Revit_EpicRibbon.ICONS.ImportInfo2_16.png");
            btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.ImportInfo2_32.png");

            btn = ribbonPanelLumi.AddItem(btn_LumiAnno) as PushButton;
            btn.Image = PngImageSource("Revit_EpicRibbon.ICONS.LumiTag2_16.png");
            btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.LumiTag2_32.png");

            btn = ribbonPanelLumi.AddItem(btn_LumiAnnoUpdate) as PushButton;
            btn.Image = PngImageSource("Revit_EpicRibbon.ICONS.LumiTagUpdate_16.png");
            btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.LumiTagUpdate_32.png");

            ribbonPanelLumi.AddSeparator();

            #endregion

            #region LumiSnap

            String dllPathLumiSnap;
            dllPathLumiSnap = dllPathLumi;

            PushButtonData btn_LumiSnap = new PushButtonData("btn_lumisnap", "Lumi Snap", dllPathLumiSnap, "EpicLumi.LumiSnap");
            PushButtonData btn_LumiSnapSettings = new PushButtonData("btn_lumisnapsettings", "Settings", dllPathLumiSnap, "EpicLumi.LumiSnapSettings");
            PushButtonData btn_LumiSnapInfo = new PushButtonData("btn_lumisnapInfo", "Info", dllPathLumiSnap, "EpicLumi.LumiSnapInfo");

            PushButtonData btn_LumiElevChange = new PushButtonData("btn_lumiElevChange", "Elv.Ch", dllPathLumiSnap, "EpicLumi.LumiElevationChange");
            PushButtonData btn_LumiFlip = new PushButtonData("btn_lumiFlip", "Flip", dllPathLumiSnap, "EpicLumi.LumiFlip");

            // Create a ribbon panel
            //RibbonPanel ribbonPanelLumiSnap = application.CreateRibbonPanel(tabName, "Lumi Snap");


            btn = ribbonPanelLumi.AddItem(btn_LumiSnap) as PushButton;
            btn.Image = PngImageSource("Revit_EpicRibbon.ICONS.LumiSnap_16.png");
            btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.LumiSnap_32.png");

            btn = ribbonPanelLumi.AddItem(btn_LumiElevChange) as PushButton;
            btn = ribbonPanelLumi.AddItem(btn_LumiFlip) as PushButton;

            //ribbonPanelLumiSnap.AddSeparator();

            var stackLumiSnap = ribbonPanelLumi.AddStackedItems(btn_LumiSnapSettings, btn_LumiSnapInfo);
            List<ImageSource> imageSourcesLumiSnap = new List<ImageSource>()
            {
                PngImageSource("Revit_EpicRibbon.ICONS.SettingsGear_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.LumiSnapInfo_16.png"),
            };
            for (int i = 0; i < stackLumiSnap.Count; i++)
            {
                btn = stackLumiSnap[i] as PushButton;
                btn.Image = imageSourcesLumiSnap[i];
            }

            //btn = ribbonPanelLumiSnap.AddItem(btn_LumiSnapSettings) as PushButton;
            //btn.Image = PngImageSource("Revit_EpicRibbon.ICONS.SettingsGear_16.png");
            //btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.SettingsGear_32.png");

            //btn = ribbonPanelLumiSnap.AddItem(btn_LumiSnapInfo) as PushButton;
            //btn.Image = PngImageSource("Revit_EpicRibbon.ICONS.LumiSnapInfo_16.png");
            //btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.LumiSnapInfo_32.png");

            #endregion

            #region TrayTagger

            String dllPathTrayTagger;
            //dllPathLumi = @"C:\Users\User\source\repos\_Revit\Revit_EpicLumImporter\bin\Debug\EpicLumi.dll";
            dllPathTrayTagger = System.IO.Path.Combine(BaseLocation, "TrayTagger", "TrayTagger.dll");

            PushButtonData btn_TrayTag = new PushButtonData("btn_traytagger", "TrayTag", dllPathTrayTagger, "TrayTagger.TrayTag");
            PushButtonData btn_TrayParams = new PushButtonData("btn_trayparams", "TrayParams", dllPathTrayTagger, "TrayTagger.TrayParamAdd");
            //PushButtonData btn_LumiInfoBlocks = new PushButtonData("btn_lumiInfoBlocks", "nfoImp", dllPathTrayTagger, "EpicLumi.LumiInfoImport");
            //PushButtonData btn_LumiAnno = new PushButtonData("btn_lumiAnno", "Tag", dllPathTrayTagger, "EpicLumi.LumiTag");
            //PushButtonData btn_LumiAnnoUpdate = new PushButtonData("btn_lumiAnnoUpdate", "Update", dllPathTrayTagger, "EpicLumi.LumiTagUpdate");

            // Create a ribbon panel
            RibbonPanel ribbonPanelTrayT = application.CreateRibbonPanel(tabName, "Tray Tagger");

            btn = ribbonPanelTrayT.AddItem(btn_TrayTag) as PushButton;
            btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.TrayTag_32.png");
            ribbonPanelTrayT.AddSeparator();
            btn = ribbonPanelTrayT.AddItem(btn_TrayParams) as PushButton;
            //btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.TrayTag_32.png");
            //btn.LargeImage = PngImageSource("Revit_EpicRibbon.ICONS.LumiMain.png");

            #endregion

            #region WallBox
            String dllPathWallBox;
            //dllPathLumi = @"C:\Users\User\source\repos\_Revit\Revit_EpicLumImporter\bin\Debug\EpicLumi.dll";
            dllPathWallBox = System.IO.Path.Combine(BaseLocation, "EpicWallBox", "EpicWallBox.dll");
            

            // Bottom
            PushButtonData btn_scBxManual_ConduitBotCenter = new PushButtonData("btn_scBxManual_BotCenter", ".", dllPathWallBox, "EpicWallBox.ManualConduitBottomCenter");
            PushButtonData btn_scBxManual_ConduitBotLeft = new PushButtonData("btn_scBxManual_BotLeft", ".", dllPathWallBox, "EpicWallBox.ManualConduitBottomLeft");
            PushButtonData btn_scBxManual_ConduitBotRight = new PushButtonData("btn_scBxManual_BotRight", ".", dllPathWallBox, "EpicWallBox.ManualConduitBottomRight");

            // Top
            PushButtonData btn_scBxManual_ConduitTopCenter = new PushButtonData("btn_scBxManual_TopCenter", ".", dllPathWallBox, "EpicWallBox.ManualConduitTopCenter");
            PushButtonData btn_scBxManual_ConduitTopLeft = new PushButtonData("btn_scBxManual_TopLeft", ".", dllPathWallBox,     "EpicWallBox.ManualConduitTopLeft");
            PushButtonData btn_scBxManual_ConduitTopRight = new PushButtonData("btn_scBxManual_TopRight", ".", dllPathWallBox,   "EpicWallBox.ManualConduitTopRight");

            // Middle
            PushButtonData btn_scBxManual_AddFromPicked = new PushButtonData("btn_scBxManual_AddSocketFromPicked", ".", dllPathWallBox, "EpicWallBox.ManualAddSocketBoxFromPicked");
            PushButtonData btn_scBxManual_AddRight = new PushButtonData("btn_scBxManual_AddSocketRight", ".", dllPathWallBox, "EpicWallBox.ManualAddSocketBoxRight");
            PushButtonData btn_scBxManual_AddLeft = new PushButtonData("btn_scBxManual_AddSocketLeft", ".", dllPathWallBox, "EpicWallBox.ManualAddSocketBoxLeft");

            // Additional
            PushButtonData btn_scBxManual_AddTop = new PushButtonData("btn_scBxManual_AddSocketTop", "TOP", dllPathWallBox, "EpicWallBox.ManualAddSocketBoxTop");
            PushButtonData btn_scBxManual_AddBot = new PushButtonData("btn_scBxManual_AddSocketBot", "BOT", dllPathWallBox, "EpicWallBox.ManualAddSocketBoxBot");


            // Other
            PushButtonData btn_scBxManual_DeleteConnected = new PushButtonData("btn_scBxManual_DeleteConnected", "Delete all Connected", dllPathWallBox, "EpicWallBox.DeleteConnected");
            PushButtonData btn_scBxManual_Settings = new PushButtonData("btn_scBxManual_Settings", "SettingsX", dllPathWallBox, "EpicWallBox.WallSnapSettings");
            PushButtonData btn_scBxManual_SnapToSelected = new PushButtonData("btn_scBxManual_SnapToSelected", "Snap To Selected", dllPathWallBox, "EpicWallBox.SnapToSelected");
            PushButtonData btn_scBxManual_ConduitToSelected = new PushButtonData("btn_scBxManual_ConduitToSelected", "Conduit To Selected", dllPathWallBox, "EpicWallBox.ConduitToSelected");


            PushButtonData btn_scBxManual_Empty01 = new PushButtonData("btn_scBxManual_Nothing01", "Nothing", dllPathWallBox, "EpicWallBox.EmptyPlaceholderButton");
            PushButtonData btn_scBxManual_Empty02 = new PushButtonData("btn_scBxManual_Nothing02", "Nothing", dllPathWallBox, "EpicWallBox.EmptyPlaceholderButton");
            PushButtonData btn_scBxManual_Empty03 = new PushButtonData("btn_scBxManual_Nothing03", "Nothing", dllPathWallBox, "EpicWallBox.EmptyPlaceholderButton");
            PushButtonData btn_scBxManual_Empty04 = new PushButtonData("btn_scBxManual_Nothing04", "Nothing", dllPathWallBox, "EpicWallBox.EmptyPlaceholderButton");
            PushButtonData btn_scBxManual_Empty05 = new PushButtonData("btn_scBxManual_Nothing05", "Nothing", dllPathWallBox, "EpicWallBox.EmptyPlaceholderButton");
            PushButtonData btn_scBxManual_Empty06 = new PushButtonData("btn_scBxManual_Nothing06", "Nothing", dllPathWallBox, "EpicWallBox.EmptyPlaceholderButton");
            PushButtonData btn_scBxManual_Empty07 = new PushButtonData("btn_scBxManual_Nothing07", "Nothing", dllPathWallBox, "EpicWallBox.EmptyPlaceholderButton");
            PushButtonData btn_scBxManual_Empty08 = new PushButtonData("btn_scBxManual_Nothing08", "Nothing", dllPathWallBox, "EpicWallBox.EmptyPlaceholderButton");
            //PushButtonData btn_scBxManual_Empty09 = new PushButtonData("btn_scBxManual_Nothing09", "Nothing", dllPathWallBox, "EpicWallBox.EmptyPlaceholderButton");



            // Create a ribbon panel
            RibbonPanel ribbonPanelWallbox = application.CreateRibbonPanel(tabName, "Epic Tools - Wall Boxes");
            
            PushButton wBbtn;

            #region Stack1
            var manualButtons4 = ribbonPanelWallbox.AddStackedItems(
                btn_scBxManual_Empty03,
                btn_scBxManual_AddLeft,
                btn_scBxManual_Empty04);
            List<ImageSource> imageSourcesManualButtons4 = new List<ImageSource>()
            {
                PngImageSource("Revit_EpicRibbon.ICONS.Empty.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_AddSocketLeft_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.Empty.png")
            };
            for (int i = 0; i < manualButtons4.Count; i++)
            {
                var rbi = manualButtons4[i];
                wBbtn = rbi as PushButton;
                wBbtn.Image = imageSourcesManualButtons4[i];
            }

            #endregion

            #region Stack2
            var manualButtons5 = ribbonPanelWallbox.AddStackedItems(
                btn_scBxManual_AddTop,
                btn_scBxManual_AddFromPicked,
                btn_scBxManual_AddBot);
            List<ImageSource> imageSourcesManualButtons5 = new List<ImageSource>()
            {
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_AddSocketUp_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_Smiley_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_AddSocketBot_16.png")
            };
            for (int i = 0; i < manualButtons5.Count; i++)
            {
                var rbi = manualButtons5[i];
                wBbtn = rbi as PushButton;
                wBbtn.Image = imageSourcesManualButtons5[i];
            }

            #endregion

            #region Stack3
            var manualButtons6 = ribbonPanelWallbox.AddStackedItems(
                btn_scBxManual_Empty05,
                btn_scBxManual_AddRight,
                btn_scBxManual_Empty06);
            List<ImageSource> imageSourcesManualButtons6 = new List<ImageSource>()
            {
                PngImageSource("Revit_EpicRibbon.ICONS.Empty.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_AddSocketRight_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.Empty.png")
            };
            for (int i = 0; i < manualButtons6.Count; i++)
            {
                var rbi = manualButtons6[i];
                wBbtn = rbi as PushButton;
                wBbtn.Image = imageSourcesManualButtons6[i];
            }

            #endregion

            ribbonPanelWallbox.AddSeparator();

            #region Stack35
            var manualButtons65 = ribbonPanelWallbox.AddStackedItems(
                btn_scBxManual_SnapToSelected,
                btn_scBxManual_Empty08,
                btn_scBxManual_ConduitToSelected
                );
            List<ImageSource> imageSourcesManualButtons65 = new List<ImageSource>()
            {
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_Smiley_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.Empty.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_Smiley_16.png")
            };
            for (int i = 0; i < manualButtons65.Count; i++)
            {
                var rbi = manualButtons65[i];
                wBbtn = rbi as PushButton;
                wBbtn.Image = imageSourcesManualButtons65[i];
            }

            #endregion

            ribbonPanelWallbox.AddSeparator();

            #region Stack4
            var manualButtonsLeft = ribbonPanelWallbox.AddStackedItems(
                btn_scBxManual_ConduitTopLeft,
                btn_scBxManual_Empty01,
                btn_scBxManual_ConduitBotLeft);

            List<ImageSource> imageSourcesManualButtonsLeft = new List<ImageSource>()
            {
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_TopLeft_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.Empty.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_BotLeft_16.png")
            };

            for (int i = 0; i < manualButtonsLeft.Count; i++)
            {
                var rbi = manualButtonsLeft[i];
                wBbtn = rbi as PushButton;
                wBbtn.Image = imageSourcesManualButtonsLeft[i];
            }
            #endregion

            #region Stack5
            var manualButtonsCenter = ribbonPanelWallbox.AddStackedItems(
                btn_scBxManual_ConduitTopCenter,
                btn_scBxManual_DeleteConnected,
                btn_scBxManual_ConduitBotCenter);
            List<ImageSource> imageSourcesManualCenterBot = new List<ImageSource>()
            {
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_TopCenter_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_DeleteConnected_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_BotCenter_16.png")
            };
            for (int i = 0; i < manualButtonsCenter.Count; i++)
            {
                var rbi = manualButtonsCenter[i];
                wBbtn = rbi as PushButton;
                wBbtn.Image = imageSourcesManualCenterBot[i];
            }

            #endregion

            #region Stack6
            var manualButtonsRight = ribbonPanelWallbox.AddStackedItems(
                btn_scBxManual_ConduitTopRight,
                btn_scBxManual_Empty02,
                btn_scBxManual_ConduitBotRight);
            List<ImageSource> imageSourcesManualButtonsRight = new List<ImageSource>()
            {
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_TopRight_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.Empty.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.SocBoxLine_BotRight_16.png")
            };
            for (int i = 0; i < manualButtonsRight.Count; i++)
            {
                var rbi = manualButtonsRight[i];
                wBbtn = rbi as PushButton;
                wBbtn.Image = imageSourcesManualButtonsRight[i];
            }

            #endregion

            ribbonPanelWallbox.AddSeparator();

            #region Stack7
            var manualButtonsAdditional = ribbonPanelWallbox.AddStackedItems(
                btn_scBxManual_Settings,
                btn_scBxManual_Empty07
                );

            List<ImageSource> imageSourcesManualButtonsAdditional = new List<ImageSource>()
            {
                PngImageSource("Revit_EpicRibbon.ICONS.SettingsGear_16.png"),
                PngImageSource("Revit_EpicRibbon.ICONS.Empty.png"),

            };
            for (int i = 0; i < manualButtonsAdditional.Count; i++)
            {
                var rbi = manualButtonsAdditional[i];
                wBbtn = rbi as PushButton;
                wBbtn.Image = imageSourcesManualButtonsAdditional[i];
            }

            #endregion

            #endregion

            application.ControlledApplication.ApplicationInitialized += OnApplicationInitialized;
            return Result.Succeeded;
        }

        #region adWin stuff
        void OnApplicationInitialized(
              object sender,
              Autodesk.Revit.DB.Events.ApplicationInitializedEventArgs e)
        {
            // Find a system ribbon tab and panel to house 
            // our API items
            // Also find our API tab, panel and button within 
            // the Autodesk.Windows.RibbonControl context

            adWin.RibbonControl adWinRibbon
              = adWin.ComponentManager.Ribbon;

            adWin.RibbonTab adWinSysTab = null;
            adWin.RibbonPanel adWinSysPanel = null;

            adWin.RibbonTab adWinApiTab = null;
            adWin.RibbonPanel adWinApiPanel = null;
            adWin.RibbonItem adWinApiItem = null;

            foreach (adWin.RibbonTab ribbonTab
              in adWinRibbon.Tabs)
            {
                // Look for the specified system tab

                if (ribbonTab.Id == "Epic Tools")
                {
                    adWinSysTab = ribbonTab;

                    foreach (adWin.RibbonPanel ribbonPanel
                      in ribbonTab.Panels)
                    {
                        // Look for the specified panel 
                        // within the system tab

                        Debug.WriteLine("Found panel: " + ribbonPanel.Source.Id);

                        if (ribbonPanel.Source.Id.Contains("Epic Tools - Wall Boxes"))
                        {
                            foreach (var item in ribbonPanel.Source.Items)
                            {
                                try
                                {
                                    adWin.RibbonRowPanel rRowPanel = (adWin.RibbonRowPanel)item;
                                    Debug.WriteLine("Found rowPanel with: " + rRowPanel.Items.Count + " items.");

                                    foreach (var rItem in rRowPanel.Items)
                                    {
                                        try
                                        {
                                            Debug.WriteLine("Found item: " + rItem.Text);
                                            rItem.ShowText = false;
                                            //var btn = (adWin.RibbonButton)rItem;
                                            //btn.Size = adWin.RibbonItemSize.Large;
                                        }
                                        catch (Exception)
                                        {
                                            Debug.WriteLine("Panel item processing error");
                                            continue;
                                        }
                                    }
                                }
                                catch (Exception)
                                {
                                    Debug.WriteLine("Panel item processing error");
                                    continue;
                                }
                                



                                
                            }
                            adWinSysPanel = ribbonPanel;
                        }
                    }
                }
                else
                {
                    //// Look for our API tab

                    //if (ribbonTab.Id == ApiTabName)
                    //{
                    //    adWinApiTab = ribbonTab;

                    //    foreach (adWin.RibbonPanel ribbonPanel
                    //      in ribbonTab.Panels)
                    //    {
                    //        // Look for our API panel.              

                    //        // The Source.Id property of an API created 
                    //        // ribbon panel has the following format: 
                    //        // CustomCtrl_%[TabName]%[PanelName]
                    //        // Where PanelName correlates with the string 
                    //        // entered as the name of the panel at creation
                    //        // The Source.AutomationName property can also 
                    //        // be used as it is also a direct correlation 
                    //        // of the panel name, but without all the cruft
                    //        // Be sure to include any new line characters 
                    //        // (\n) used for the panel name at creation as 
                    //        // they still form part of the Id & AutomationName

                    //        //if(ribbonPanel.Source.AutomationName 
                    //        //  == ApiPanelName) // Alternative method

                    //        if (ribbonPanel.Source.Id ==
                    //          "CustomCtrl_%" + ApiTabName + "%" + ApiPanelName)
                    //        {
                    //            adWinApiPanel = ribbonPanel;

                    //            foreach (adWin.RibbonItem ribbonItem
                    //              in ribbonPanel.Source.Items)
                    //            {
                    //                // Look for our command button

                    //                // The Id property of an API created ribbon 
                    //                // item has the following format: 
                    //                // CustomCtrl_%CustomCtrl_%[TabName]%[PanelName]%[ItemName]
                    //                // Where ItemName correlates with the string 
                    //                // entered as the first parameter (name) 
                    //                // of the PushButtonData() constructor
                    //                // While AutomationName correlates with 
                    //                // the string entered as the second 
                    //                // parameter (text) of the PushButtonData() 
                    //                // constructor
                    //                // Be sure to include any new line 
                    //                // characters (\n) used for the button 
                    //                // name and text at creation as they 
                    //                // still form part of the ItemName 
                    //                // & AutomationName

                    //                //if(ribbonItem.AutomationName 
                    //                //  == ApiButtonText) // alternative method

                    //                if (ribbonItem.Id
                    //                  == "CustomCtrl_%CustomCtrl_%"
                    //                  + ApiTabName + "%" + ApiPanelName
                    //                  + "%" + ApiButtonName)
                    //                {
                    //                    adWinApiItem = ribbonItem;
                    //                }
                    //            }
                    //        }
                    //    }
                    //}
                }
            }

            // Make sure we got everything we need

            if (adWinSysTab != null
              && adWinSysPanel != null
              && adWinApiTab != null
              && adWinApiPanel != null
              && adWinApiItem != null)
            {
                // First we'll add the whole panel including 
                // the button to the system tab

                adWinSysTab.Panels.Add(adWinApiPanel);

                // now lets also add the button itself 
                // to a system panel

                adWinSysPanel.Source.Items.Add(adWinApiItem);

                // Remove panel from original API tab
                // It can also be left there if needed, 
                // there doesn't seem to be any problems with
                // duplicate panels / buttons on seperate tabs 
                // / panels respectively

                adWinApiTab.Panels.Remove(adWinApiPanel);

                // Remove our original API tab from the ribbon

                adWinRibbon.Tabs.Remove(adWinApiTab);
            }

            // A new panel should now be added to the 
            // specified system tab. Its command buttons 
            // will behave as they normally would, including 
            // API access and ExternalCommandAvailability tests. 
            // There will also be a second copy of the command 
            // button from the panel added to the specified 
            // system panel.

        }
        #endregion

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
