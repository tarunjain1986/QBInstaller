using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading;
using FrameworkLibraries.Utils;
using TestStack.White.UIItems.WindowItems;
using TestStack.White;
using FrameworkLibraries.ActionLibs;
using TestStack.White.UIItems;
using System.Windows.Forms;
using System.Windows.Automation;
using TestStack.White.UIItems.Finders;
using FrameworkLibraries.ObjMaps.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using System.IO;
using System.Reflection;

namespace FrameworkLibraries.AppLibs.MayaConnected
{
    public class Maya
    {
        public static Property conf = Property.GetPropertyInstance();
        public static string Execution_Speed = conf.get("ExecutionSpeed");
        public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        public static int ResetWindow_Timeout = int.Parse(conf.get("ResetWindowTimeOut"));
        public static string UserName = conf.get("MayaLoginUserName");
        public static string Password = conf.get("MayaLoginPassword");
        public static string DefaultCompanyFile = conf.get("DefaultCompanyFile");
        public static string DefaultCompanyFilePath = DefaultCompanyFile;
        public static string TestDataSourceDirectory = conf.get("TestDataSourceDirectory");
        public static string TestDataLocalDirectory = conf.get("TestDataLocalDirectory");

        //**************************************************************************************************************************************************************

        public static void LoginIntoMaya(Window window, bool selectCompanyFile)
        {
            try
            {
                Logger.logMessage("Function call @ :" + DateTime.Now);
                Logger.logMessage("LoginIntoMaya " + window);
                var allPanes = Actions.GetAllPanesInWindow(window);
                var allEditBoxes = Actions.GetAllEditBoxesInsideAPane(window, allPanes[2]);
                Actions.WaitForTextVisibleInsidePane(window, allPanes[2], "Sign In", Sync_Timeout);
                Actions.ClickElement(allEditBoxes[0]);
                Actions.SendKeysToWindow(window, UserName);
                Actions.ClickElement(allEditBoxes[1]);
                Actions.SendKeysToWindow(window, Password);
                Actions.ClickButtonInsidePanelByName(window, allPanes[2], "Sign In");

                if(selectCompanyFile)
                {
                    Actions.WaitForElementVisible(window, DefaultCompanyFile, Sync_Timeout);
                    Actions.ClickTextInsidePanel(window, allPanes[2], DefaultCompanyFile);
                }

                Actions.WaitForTextVisibleInsidePane(window, allPanes[2], "Home", Sync_Timeout);

                try { Actions.ClickButtonInsidePanelByName(window, allPanes[2], "Cancel"); }
                catch { }

                Thread.Sleep(int.Parse(Execution_Speed) * 10);
                Logger.logMessage("LoginIntoMaya " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("LoginIntoMaya " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static void LoginIntoMaya(Window window, bool selectCompanyFile, string companyFile)
        {
            try
            {
                Logger.logMessage("Function call @ :" + DateTime.Now);
                Logger.logMessage("LoginIntoMaya " + window + "->" + companyFile);

                var allPanes = Actions.GetAllPanesInWindow(window);
                var allEditBoxes = Actions.GetAllEditBoxesInsideAPane(window, allPanes[2]);
                Actions.WaitForTextVisibleInsidePane(window, allPanes[2], "Sign In", Sync_Timeout);
                Actions.ClickElement(allEditBoxes[0]);
                Actions.SendKeysToWindow(window, UserName);
                Actions.ClickElement(allEditBoxes[1]);
                Actions.SendKeysToWindow(window, Password);
                Actions.ClickButtonInsidePanelByName(window, allPanes[2], "Sign In");

                if (selectCompanyFile)
                {
                    Actions.WaitForElementVisible(window, companyFile, Sync_Timeout);
                    Actions.ClickTextInsidePanel(window, allPanes[2], companyFile);
                }

                Actions.WaitForTextVisibleInsidePane(window, allPanes[2], "Home", Sync_Timeout);
                
                try { Actions.ClickButtonInsidePanelByName(window, allPanes[2], "Cancel"); }
                catch { }
                Logger.logMessage("LoginIntoMaya " + window + "->" + companyFile + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("LoginIntoMaya " + window + "->" + companyFile + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static void PrepareMayaBaseState(Window window)
        {
            try
            {
                Logger.logMessage("Function call @ :" + DateTime.Now);
                Logger.logMessage("PrepareMayaBaseState " + window);

                var allPanes = Actions.GetAllPanesInWindow(window);

                //"Subscribe to QuickBooks Online" handler
                if (Actions.CheckTextExistsInsidePane(allPanes[2], "Subscribe to QuickBooks Online"))
                {
                    Actions.ClickButtonInsidePanelByName(window, allPanes[2], "Cancel");
                }


                Logger.logMessage("LoginIntoMaya " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("LoginIntoMaya " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static void SingOutMaya(Window window, string identifier)
        {
            try
            {
                Logger.logMessage("Function call @ :" + DateTime.Now);
                Logger.logMessage("SingOutMaya " + window);

                var allPanes = Actions.GetAllPanesInWindow(window);
                var menuItems = Actions.GetAllMenuItemsInsideAPane(window, allPanes[2]);

                foreach (var item in menuItems)
                {
                    if (item.Name.Contains(identifier))
                        item.Click();
                }

                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);
                Actions.SendTABToWindow(window);

                Actions.SendENTERoWindow(window);

                Logger.logMessage("SingOutMaya " + window + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("SingOutMaya " + window + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************


    }

}
