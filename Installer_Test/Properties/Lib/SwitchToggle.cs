using System;
using System.Threading;
using System.Windows.Forms;

using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;

using Installer_Test.Lib;
using Installer_Test.Properties.Lib;

using TestStack.White.UIItems.WindowItems;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Installer_Test.Lib
{

    public class SwitchToggle
    {
        public static string[] arrEdition;
        public static string currEdition;

        public static Property conf = Property.GetPropertyInstance();
        public static string exe = conf.get("QBExePath");

        public static void SwitchEdition(string SKU)
        {
            Logger.logMessage("Switch Edition - Started");
            Logger.logMessage("-----------------------------------------------------------");
 
            if (SKU == "Enterprise")
            {
                arrEdition = new string[] {"Enterprise Solutions General Business" , "Enterprise Solutions Contractor","Enterprise Solutions Manufacturing & Wholesale   ", "Enterprise Solutions Nonprofit",
                "Enterprise Solutions Professional Services", "Enterprise Solutions Retail"};
            }

            if (SKU == "Premier")
            {
                arrEdition = new string[] {"Premier Edition (General Business)" , "Premier Contractor Edition", "Premier Manufacturing & Wholesale Edition   ", "Premier Nonprofit Edition",
                "Premier Professional Services Edition", "Premier Retail Edition"};
            }

            for (int i = 1; i < arrEdition.Length; i++)
            {
                Perform_Switch(arrEdition[i], SKU);
            }

            Perform_Switch(arrEdition[0], SKU);

            Logger.logMessage("Switch Edition - Completed");
            Logger.logMessage("-----------------------------------------------------------");

        }

        public static void Perform_Switch(string currEdition, string SKU)
        {
            
            TestStack.White.Application qbApp = null;
            TestStack.White.UIItems.WindowItems.Window qbWindow = null;
            //qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            //qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);

            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");

            try
            {

                Actions.SelectMenu(qbApp, qbWindow, "Help", "Manage My License", "Change to a Different Industry Edition...");
                Thread.Sleep(500);
                Window editionWindow = Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition");

            
                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition"), currEdition);
                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition"), "Next >");
                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition"), "Finish");

                Thread.Sleep(2000);

                if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Automatic Backup") == true)
                {
                    SendKeys.SendWait("%N");
                }
                
                Install_Functions.Select_Edition(SKU);

                //qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
                //qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);

                qbApp = QuickBooks.GetApp("QuickBooks");
                qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");
                Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");

                //QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
                Thread.Sleep(10000);

                Logger.logMessage("Switch Edition - Successful");
                Logger.logMessage("-----------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Switch Edition - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("-----------------------------------------------------------");
            }
        }

        public static void ToggleEdition(string SKU)
        {
            Logger.logMessage("Toggle Edition - Started");
            Logger.logMessage("-----------------------------------------------------------");

            if (SKU == "Enterprise")
            {
                arrEdition = new string[] {"Enterprise Solutions General Business" , "Enterprise Solutions Accountant - Home  ","Enterprise Solutions Contractor", "Enterprise Solutions Manufacturing & Wholesale   ",
                "Enterprise Solutions Nonprofit", "Enterprise Solutions Professional Services", "Enterprise Solutions Retail"};
            }

            if (SKU == "Premier")
            {
                arrEdition = new string[] {"Premier Edition (General Business)" , "Premier Accountant Edition - Home  ", "Premier Contractor Edition", "Premier Manufacturing & Wholesale Edition  ",
                "Premier Nonprofit Edition", "Premier Professional Services Edition", "Premier Retail Edition" , "QuickBooks Pro"};
            }

            for (int i = 2; i < arrEdition.Length; i++)
            {
                Perform_Toggle(arrEdition[i], SKU);
            }

            Perform_Toggle(arrEdition[0], SKU);
            Perform_Toggle(arrEdition[1], SKU);

            Logger.logMessage("Toggle Edition - Completed");
            Logger.logMessage("-----------------------------------------------------------");
        }

        public static void Perform_Toggle(string currEdition, string SKU)
        {
            TestStack.White.Application qbApp = null;
            TestStack.White.UIItems.WindowItems.Window qbWindow = null;

            //qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            //qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);

            qbApp = QuickBooks.GetApp("QuickBooks");
            qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");

            try
            {
                Actions.SelectMenu(qbApp, qbWindow, "File", "Toggle to Another Edition...");
                Thread.Sleep(500);
                Window editionWindow = Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition");
         
                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition"), currEdition);
                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition"), "Next >");
                Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition"), "Toggle");

                Thread.Sleep(2000);

                if (Actions.CheckWindowExists(Actions.GetDesktopWindow("QuickBooks"), "Automatic Backup") == true)
                {
                  SendKeys.SendWait("%N");
                }
                
                Install_Functions.Select_Edition(SKU);
                Thread.Sleep(20000);

                qbApp = QuickBooks.GetApp("QuickBooks");
                qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");

                Actions.SelectMenu(qbApp, qbWindow, "Window", "Close All");

                //QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
                Thread.Sleep(10000);

                Logger.logMessage("Toggle Edition - Successful");
                Logger.logMessage("-----------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Toggle Edition - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("-----------------------------------------------------------");
            }
        }
    }
}
