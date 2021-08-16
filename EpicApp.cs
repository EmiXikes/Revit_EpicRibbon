using Autodesk.Revit.UI;
using System;
using System.Windows.Media;
using System.Windows.Media.Imaging;

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

            String dllPath = @"C:\Users\User\source\repos\_Revit\Revit_SystemSetter\bin\Debug\Revit_SystemSetter.dll";

            // EL
            PushButtonData btn_H2 = new PushButtonData("btn_H2", "H2", dllPath, "Revit_SystemSetter.SetSystem_H2");
            PushButtonData btn_H201 = new PushButtonData("btn_H201", "H201", dllPath, "Revit_SystemSetter.SetSystem_H201");
            PushButtonData btn_H202 = new PushButtonData("btn_H202", "H202", dllPath, "Revit_SystemSetter.SetSystem_H202");
            PushButtonData btn_H203 = new PushButtonData("btn_H203", "H203", dllPath, "Revit_SystemSetter.SetSystem_H203");
            btn_H2.Image = new BitmapImage(new Uri(@"C:\_x\source\Refs\EpicStyle\IMG\1_16.ico"));
            btn_H2.LargeImage = new BitmapImage(new Uri(@"C:\_x\source\Refs\EpicStyle\IMG\1_32.ico"));

            //Voids
            PushButtonData btn_H2V1 = new PushButtonData("btn_H2V1", "SAME", dllPath, "Revit_SystemSetter.SetSystem_H2V1");
            PushButtonData btn_H2V2 = new PushButtonData("btn_H2V2", "NEW", dllPath, "Revit_SystemSetter.SetSystem_H2V2");
            PushButtonData btn_H2V3 = new PushButtonData("btn_H2V3", "OLD", dllPath, "Revit_SystemSetter.SetSystem_H2V3");

            // EL + ESS
            PushButtonData btn_HJ201 = new PushButtonData("btn_HJ201", "HJ201", dllPath, "Revit_SystemSetter.SetSystem_HJ201");


            // ESS
            PushButtonData btn_J2 = new PushButtonData("btn_J2", "J2", dllPath, "Revit_SystemSetter.SetSystem_J2");
            PushButtonData btn_J201 = new PushButtonData("btn_J201", "J201" ,dllPath, "Revit_SystemSetter.SetSystem_J201");

            // Lights
            PushButtonData btn_H501 = new PushButtonData("btn_H501", "H501" ,dllPath, "Revit_SystemSetter.SetSystem_H501");
            PushButtonData btn_H502 = new PushButtonData("btn_H502", "H502" ,dllPath, "Revit_SystemSetter.SetSystem_H502");

            //Lightning
            PushButtonData btn_H701 = new PushButtonData("btn_H701", "H701" ,dllPath, "Revit_SystemSetter.SetSystem_H701");
            PushButtonData btn_H702 = new PushButtonData("btn_H702", "H702" ,dllPath, "Revit_SystemSetter.SetSystem_H702");



            // Create a ribbon panel
            RibbonPanel m_projectPanel = application.CreateRibbonPanel(tabName, "System Changer");

            // Add the buttons to the panel
            //List<RibbonItem> projectButtons = new List<RibbonItem>();
            //projectButtons.AddRange(m_projectPanel.AddStackedItems(button1, button2));

            m_projectPanel.AddItem(btn_H2);
            m_projectPanel.AddStackedItems(btn_H201, btn_H202, btn_H203);
            m_projectPanel.AddSeparator();
            m_projectPanel.AddStackedItems(btn_J2, btn_J201);
            m_projectPanel.AddSeparator();
            m_projectPanel.AddStackedItems(btn_H501, btn_H502);
            m_projectPanel.AddSeparator();
            m_projectPanel.AddStackedItems(btn_H701, btn_H702);
            m_projectPanel.AddSeparator();
            m_projectPanel.AddStackedItems(btn_H2V1, btn_H2V2, btn_H2V3);

            return Result.Succeeded;
        }
    }
}
