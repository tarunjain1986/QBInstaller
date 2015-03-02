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

namespace FrameworkLibraries.AppLibs.QBDT
{
    public class QuickBooks
    {
        public static Property conf = Property.GetPropertyInstance();
        public static string Execution_Speed = conf.get("ExecutionSpeed");
        public static string Sync_Timeout = conf.get("SyncTimeOut");
        public static string ResetWindow_Timeout = conf.get("ResetWindowTimeOut");
        public static string UserName = conf.get("QBLoginUserName");
        public static string Password = conf.get("QBLoginPassword");
        public static string DefaultCompanyFile = conf.get("DefaultCompanyFile");
        public static string DefaultCompanyFilePath = DefaultCompanyFile;
        public static string TestDataSourceDirectory = conf.get("TestDataSourceDirectory");
        public static string TestDataLocalDirectory = conf.get("TestDataLocalDirectory");
        public static string QbwINI = conf.get("QBW.ini");

        //**************************************************************************************************************************************************************

        public static TestStack.White.Application Initialize(String exePath)
        {
            ///////////////////////////////////////////
            conf.reload();
            QbwINI = conf.get("QBW.ini");
            ///////////////////////////////////////////

            Logger.logMessage("Initialize " + exePath);
            
            var accessiblity = FrameworkLibraries.Utils.FileOperations.CheckForStringInFile(QbwINI, "QBACCESSIBILITY=1");
            if (!accessiblity)
            {
                Logger.logMessage("QBAccessiblity settings not availble in - " + QbwINI);
                Logger.logMessage("Trying to set QBACCESSIBILITY=1 and kill any existing QBW32 process..");
                FileOperations.AppendStringToFile(QbwINI, "[ACCESSIBILITY]");
                FileOperations.AppendStringToFile(QbwINI, "QBACCESSIBILITY=1");
                OSOperations.KillProcess("QBW32");
                Thread.Sleep(5000);
            }

            int processID = 0;
            TestStack.White.Application app = null;

            try
            {
                List<Window> allWin = Desktop.Instance.Windows();
                foreach (Window item in allWin)
                {
                    if (item.Name.Contains("QuickBooks"))
                    {
                        foreach (Process p in Process.GetProcesses("."))
                        {
                            if (p.ProcessName.Contains("QBW32") || p.ProcessName.Contains("qbw32"))
                            {
                                processID = p.Id;
                                app = TestStack.White.Application.Attach(processID);
                                app.WaitWhileBusy();
                                Actions.WaitForAppWindow("QuickBooks", int.Parse(Sync_Timeout));
                                Logger.logMessage("Existing QB instance found..!!");
                                return app;
                            }
                        }
                    }
                }

                Logger.logMessage("No existing QB instance, so launching - " + exePath);
                Process proc = new Process();
                proc.StartInfo.FileName = exePath;
                proc.Start();
                Thread.Sleep(7500);

                //Alert window handler
                if (Actions.CheckDesktopWindowExists("Alert"))
                    Actions.CheckForAlertAndClose("Alert");

                //Crash handler
                if (Actions.CheckDesktopWindowExists("QuickBooks - Unrecoverable Error"))
                    Actions.QBCrashHandler();
                
                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.WaitForAppWindow("QuickBooks", int.Parse(Sync_Timeout));
                }
                catch (Exception) { }
                Thread.Sleep(int.Parse(Execution_Speed));
                foreach (Process p in Process.GetProcesses("."))
                {
                    if (p.ProcessName.Contains("QBW32") || p.ProcessName.Contains("qbw32"))
                    {
                        processID = p.Id;
                    }
                }
                app = TestStack.White.Application.Attach(processID);
                app.WaitWhileBusy();
                Thread.Sleep(int.Parse(Execution_Speed));

                Logger.logMessage("Initialize " + exePath + " - Sucessful");
                Logger.logMessage("------------------------------------------------------------------------------");

                return app;
            }
            catch (Exception e)
            {
                Logger.logMessage("Initialize " + exePath + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static Window PrepareBaseState(TestStack.White.Application app)
        {
            Window qbWin = null;

            try
            {
                List<Window> windows = app.GetWindows();
                foreach (Window item in windows)
                {
                    if (item.Name.Contains("QuickBooks"))
                    {
                        qbWin = item;
                        Thread.Sleep(int.Parse(Execution_Speed));
                        break;
                    }
                }

                Logger.logMessage("PrepareBaseState " + app + " - Sucessful");
                Logger.logMessage(qbWin.Title);
                Logger.logMessage("------------------------------------------------------------------------------");

                return qbWin;

            }
            catch (Exception e)
            {
                Logger.logMessage("PrepareBaseState " + app + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static Window MaximizeQB(TestStack.White.Application app)
        {
            Window qbWin = null;

            try
            {
                List<Window> windows = app.GetWindows();
                foreach (Window item in windows)
                {
                    if (item.Name.Contains("QuickBooks"))
                    {
                        qbWin = item;
                        qbWin.DisplayState = TestStack.White.UIItems.WindowItems.DisplayState.Maximized;
                        Thread.Sleep(int.Parse(Execution_Speed));
                        break;
                    }
                }

                Logger.logMessage("Maximized " + app + " - Successful");
                Logger.logMessage(qbWin.Title);
                Logger.logMessage("------------------------------------------------------------------------------");

                return qbWin;

            }
            catch (Exception e)
            {
                Logger.logMessage("Maximized " + app + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static void RepairOrUnInstallQB(string qbVersion, bool repair, bool remove)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("Repair/UnInstallQB " + qbVersion + " - Started..");

            try
            {
                FrameworkLibraries.Utils.OSOperations.CommandLineExecute("control appwiz.cpl");

                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.WaitForWindow("Programs and Features", int.Parse(Sync_Timeout));
                }
                catch { }


                Actions.SetFocusOnWindow(Actions.GetDesktopWindow("Programs and Features"));
                var controlPanelWindow = Actions.GetDesktopWindow("Programs and Features");
                var uiaWindow = Actions.UIA_GetAppWindow("Programs and Features");

                if (controlPanelWindow.DisplayState != DisplayState.Maximized)
                   controlPanelWindow.DisplayState = TestStack.White.UIItems.WindowItems.DisplayState.Maximized;
                Thread.Sleep(int.Parse(Execution_Speed));
                

                // If the uiaWindow is not created for some reason.
                if (uiaWindow == null)
                {
                    controlPanelWindow.Enter(qbVersion);                    
                    Thread.Sleep(2000);
                    Actions.ClickElementByName(controlPanelWindow, "Uninstall/Change");
                    Thread.Sleep(2000);
                }
                // else if the uiaWindow is created successfully.
                else
                {
                     Actions.UIA_SetTextByName(uiaWindow, Actions.GetDesktopWindow("Programs and Features"), "Search Box", qbVersion);
                     Thread.Sleep(int.Parse(Execution_Speed));
                     try
                     {
                        Logger.logMessage("---------------Try-Catch Block------------------------");
                        Actions.WaitForElementEnabled(Actions.GetDesktopWindow("Programs and Features"), qbVersion, int.Parse(Sync_Timeout));
                     }
                     catch (Exception e)
                     {
                        Logger.logMessage("---------------------------------------------------------");
                        Logger.logMessage("Element not enabled " + qbVersion);
                        Logger.logMessage(e.Message);
                        Logger.logMessage("---------------------------------------------------------");
                     }

                     Actions.UIA_ClickEditControlByName(uiaWindow, Actions.GetDesktopWindow("Programs and Features"), qbVersion);
                     try
                     {
                         Logger.logMessage("---------------Try-Catch Block------------------------");
                         Actions.WaitForElementEnabled(Actions.GetDesktopWindow("Programs and Features"), "Uninstall/Change", int.Parse(Sync_Timeout));
                     }
                     catch { }

                     Actions.UIA_ClickItemByName(uiaWindow, Actions.GetDesktopWindow("Programs and Features"), "Uninstall/Change");

                 }
   
                   
                // Wait for the Uninstall flow to trigger.
                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.WaitForWindow("QuickBooks Installation", int.Parse(Sync_Timeout));
                }
                catch { }

                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.WaitForElementEnabled(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >", int.Parse(Sync_Timeout));
                }
                catch { }

                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");

                
                if(remove)
                {
                    Logger.logMessage("Remove " + qbVersion + " - Started..");

                    Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Remove");
                    Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                    Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Remove");
                }
                
                
                if (repair)
                {
                    Logger.logMessage("Repair " + qbVersion + " - Started..");

                    Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Repair");
                    Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                    Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Repair");
                }

                    try
                    {
                        Logger.logMessage("---------------Try-Catch Block------------------------");
                        if (Actions.CheckElementExistsByName(Actions.GetDesktopWindow("QuickBooks Installation"), "OK"))
                            Actions.WaitForElementEnabled(Actions.GetDesktopWindow("QuickBooks Installation"), "OK", int.Parse(Sync_Timeout));
                    }
                    catch { }

                    try
                    {
                        Logger.logMessage("---------------Try-Catch Block------------------------");
                        if (Actions.CheckElementExistsByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >"))
                            Actions.WaitForElementEnabledOrTransformed(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >", "Finish", int.Parse(Sync_Timeout));
                    }
                    catch { }

                    try
                    {
                        Logger.logMessage("---------------Try-Catch Block------------------------");
                        if (Actions.CheckElementExistsByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Files in Use"))
                        {
                            Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Do not close applications. (A reboot will be required.)");
                            Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "OK");
                        }
                    }
                    catch { }

                    try
                    {
                        Logger.logMessage("---------------Try-Catch Block------------------------");
                        if (Actions.CheckElementExistsByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >"))
                            Actions.WaitForElementEnabledOrTransformed(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >", "Finish", int.Parse(Sync_Timeout));
                    }
                    catch { }

                    Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Finish");


                    if (Actions.CheckDesktopWindowExists("QuickBooks Installation Information"))
                        Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation Information"), "No");

                    Logger.logMessage("Repair " + qbVersion + " - Successful");
                    Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Repair/UnInstallQB " + qbVersion + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************

        public static void InstallQB(string installDir, string exeName, string licenseNumber, string productNumber)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("InstallQB " + installDir + " - Started..");
            Logger.logMessage("License Number: " + licenseNumber);
            Logger.logMessage("Product Number " + productNumber);

            try
            {
                OSOperations.InvokeInstaller(installDir, exeName);
                try { Actions.WaitForWindow("QuickBooks Installation", int.Parse(Sync_Timeout)); }
                catch { }
                try { Actions.WaitForElementEnabled(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >", int.Parse(Sync_Timeout)); }
                catch { }
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "I accept the terms of the license agreement");
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Express (recommended)");
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");

                var license = StringFunctions.SplitString(licenseNumber);
                var product = StringFunctions.SplitString(productNumber);

                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1054", license[0]);
                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1055", license[1]);
                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1056", license[2]);
                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1057", license[3]);
                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1059", product[0]);
                Actions.SetTextByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1060", product[1]);
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Next >");
                Actions.ClickElementByAutomationID(Actions.GetDesktopWindow("QuickBooks Installation"), "1");
                try { Actions.WaitForElementVisible(Actions.GetDesktopWindow("QuickBooks Installation"), "Finish", int.Parse(Sync_Timeout)); }
                catch { }
                Actions.ClickElementByName(Actions.GetDesktopWindow("QuickBooks Installation"), "Finish");

                try
                {
                    if (Actions.DesktopInstance_CheckElementExistsByAutomationID("1"))
                        Actions.DesktopInstance_ClickElementByAutomationID("1");
                }
                catch { }

                Logger.logMessage("InstallQB " + installDir + " - Successful");
            }
            catch (Exception e)
            {
                Logger.logMessage("InstallQB " + installDir + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }


        //**************************************************************************************************************************************************************


        public static bool CreateInvoice(TestStack.White.Application qbApp, Window qbWindow, String customer, String cls, String account, String template, int invoiceNumber, int poNumber, String terms, String via, String fob, String quatity, String item, String itemDesc, bool markPending)
        {
            try
            {
                Actions.SelectMenu(qbApp, qbWindow, "Customers", "Create Invoices");
                Thread.Sleep(int.Parse(Execution_Speed));
                Window invoiceWindow = Actions.GetWindow(qbWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Invoice.Objects.CreateInvoice_Window_Name);

                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.ClickElementByAutomationID(invoiceWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Invoice.Objects.MaximizeWindow_Button_AutoID);
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
                catch (Exception)
                { }

                Actions.ClickButtonByAutomationID(invoiceWindow, "PrevBtn");
                Actions.SendKeysToWindow(invoiceWindow, customer);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendKeysToWindow(invoiceWindow, cls);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendKeysToWindow(invoiceWindow, account);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendKeysToWindow(invoiceWindow, template);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendBCKSPACEToWindow(invoiceWindow);
                Actions.SendKeysToWindow(invoiceWindow, Convert.ToString(invoiceNumber));
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendBCKSPACEToWindow(invoiceWindow);
                Actions.SendKeysToWindow(invoiceWindow, Convert.ToString(poNumber));
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendKeysToWindow(invoiceWindow, terms);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendKeysToWindow(invoiceWindow, via);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendKeysToWindow(invoiceWindow, fob);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendKeysToWindow(invoiceWindow, quatity);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendKeysToWindow(invoiceWindow, item);
                Actions.SendTABToWindow(invoiceWindow);
                Actions.SendSHIFT_ENDToWindow(invoiceWindow);

                if (markPending)
                { Actions.ClickButtonByAutomationID(invoiceWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Invoice.Objects.MarkPending_Button_AutoID); }

                Actions.ClickElementByName(invoiceWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Invoice.Objects.SaveClsoe_Button_Name);

                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Recording Transaction"), "Yes");
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
                catch { }

                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Enter Memorized Transactions Later"), "Ok");
                }
                catch { }

                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Information Changed"), "No");
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
                catch { }

                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Past Transactions"), "No");
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
                catch { }

                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Available Credits"), "No");
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
                catch { }

                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Transaction Cleared"), "No");
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
                catch { }


                return true;

            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static TestStack.White.Application GetApp(string appName)
        {
            int processID = 0;
            TestStack.White.Application app = null;

            try
            {
                List<Window> allWin = Desktop.Instance.Windows();
                foreach (Window item in allWin)
                {
                    if (item.Name.Contains(appName))
                    {
                        foreach (Process p in Process.GetProcesses("."))
                        {
                            if (p.ProcessName.Contains("QBW32") || p.ProcessName.Contains("qbw32"))
                            {
                                processID = p.Id;
                                app = TestStack.White.Application.Attach(processID);
                                app.WaitWhileBusy();
                                Thread.Sleep(int.Parse(Execution_Speed));
                                break;
                            }
                        }
                    }
                }

                return app;
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static Window GetAppWindow(TestStack.White.Application app, string winName)
        {
            Window win = null;

            try
            {
                List<Window> allWin = app.GetWindows();

                foreach (Window item in allWin)
                {
                    if (item.Name.Contains(winName))
                    {
                        win = item;
                        break;
                    }
                }

                return win;
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static void OpenOrUpgradeCompanyFile(string companyFilePath, TestStack.White.Application qbApp, Window qbWindow, bool backupcopy, bool portalcopy)
        {
            Logger.logMessage("OpenOrUpgradeCompanyFile " + companyFilePath + "->" + " - Begin");
            try
            {
                Thread.Sleep(int.Parse(Execution_Speed));
                Actions.SelectMenu(qbApp, qbWindow, "File", "Open or Restore Company...");
                Thread.Sleep(int.Parse(Execution_Speed));

                if (backupcopy)
                {
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open or Restore Company"), "Restore a backup copy");
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open or Restore Company"), "Next");
                    Thread.Sleep(int.Parse(Execution_Speed));
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open or Restore Company"), "Local backup");
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open or Restore Company"), "Next");
                    Actions.WaitForAnyChildWindow(qbWindow, "Open or Restore Company", int.Parse(Sync_Timeout));

                    try
                    {
                        Logger.logMessage("---------------Try-Catch Block------------------------");
                        Actions.SetTextOnElementByName(Actions.GetChildWindow(qbWindow, "Open Backup Copy"), "File name:", companyFilePath);
                        Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open Backup Copy"), "Open");
                        Actions.WaitForChildWindow(qbWindow, "Open or Restore Company", int.Parse(Sync_Timeout));
                    }
                    catch (Exception) { }

                    try
                    {
                        Logger.logMessage("---------------Try-Catch Block------------------------");
                        Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open or Restore Company"), "Next");
                        Thread.Sleep(int.Parse(Execution_Speed));
                    }
                    catch (Exception) { }

                    try
                    {

                        Logger.logMessage("---------------Try-Catch Block------------------------");
                        Actions.SetTextOnElementByName(Actions.GetChildWindow(qbWindow, "Save Company File as"), "File name:", Utils.StringFunctions.RandomString(5));
                        Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Save Company File as"), "Save");
                        Actions.WaitForAnyChildWindow(qbWindow, "Save Company File as", int.Parse(Sync_Timeout));
                    }
                    catch (Exception) { }

                }
                else if (portalcopy)
                {
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open or Restore Company"), "Restore a portable file");
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open or Restore Company"), "Next");
                    Actions.WaitForAnyChildWindow(qbWindow, "Open or Restore Company", int.Parse(Sync_Timeout));

                    try
                    {
                        Logger.logMessage("---------------Try-Catch Block------------------------");
                        Actions.SetTextOnElementByName(Actions.GetChildWindow(qbWindow, "Open Portable Company File"), "File name:", companyFilePath);
                        Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open Portable Company File"), "Open");
                        Actions.WaitForChildWindow(qbWindow, "Open or Restore Company", int.Parse(Sync_Timeout));
                    }
                    catch (Exception) { }

                    try
                    {
                        Logger.logMessage("---------------Try-Catch Block------------------------");
                        Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open or Restore Company"), "Next");
                        Thread.Sleep(int.Parse(Execution_Speed));
                    }
                    catch (Exception) { }

                    try
                    {
                        Logger.logMessage("---------------Try-Catch Block------------------------");
                        Actions.SetTextOnElementByName(Actions.GetChildWindow(qbWindow, "Save Company File as"), "File name:", Utils.StringFunctions.RandomString(5));
                        Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Save Company File as"), "Save");
                        Actions.WaitForAnyChildWindow(qbWindow, "Save Company File as", int.Parse(Sync_Timeout));
                    }
                    catch (Exception) { }

                }
                else
                {
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open or Restore Company"), "Open a company file");
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open or Restore Company"), "Next");
                    Actions.WaitForAnyChildWindow(qbWindow, "Open or Restore Company", int.Parse(Sync_Timeout));
                    Actions.SetTextOnElementByName(Actions.GetChildWindow(qbWindow, "Open a Company"), "File name:", companyFilePath);
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Open a Company"), "Open");
                    Actions.WaitForAnyChildWindow(qbWindow, "Open a Company", int.Parse(Sync_Timeout));
                }

                List<Window> modalWin = null;
                int iteration = 0;

                do
                {
                    modalWin = qbWindow.ModalWindows();
                    iteration = iteration + 1;

                    if (iteration <= 7)
                    {
                        foreach (Window item in modalWin)
                        {

                            //QB Login window handler
                            if (item.Name.Contains("QuickBooks Login"))
                            {
                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.SetFocusOnWindow(item);
                                    Actions.SendBCKSPACEToWindow(item);
                                    Actions.SetTextByAutomationID(item, "15922", UserName);
                                    Actions.SendTABToWindow(item);
                                    Actions.SendKeysToWindow(item, Password);
                                    Actions.ClickElementByAutomationID(item, "51");
                                    Actions.WaitForAnyChildWindow(qbWindow, "QuickBooks Login", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }
                            }

                            //Register quickbooks window handler
                            else if (item.Name.Contains("Register QuickBooks"))
                            {
                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Remind Me Later");
                                    Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                }
                                catch { }
                            }

                            //Update to new version window handler - I agree
                            else if (item.Name.Contains("Update Company File for New Version") || item.Name.Contains("Update Company File to New Version"))
                            {
                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "I understand that my company file will be updated to this new version of QuickBooks.");
                                }
                                catch (Exception) { }
                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Update Now");
                                    Actions.WaitForAnyChildWindow(qbWindow, "Update Company File", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }
                            }

                            //QB Backup
                            else if (item.Name.Contains("QuickBooks Backup"))
                            {
                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "OK");
                                    Actions.WaitForChildWindow(qbWindow, "Backup", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }

                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Backup"), "Yes");
                                    Actions.WaitForAnyChildWindow(qbWindow, "Backup", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }

                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Backup Incompatible"), "OK");
                                    Thread.Sleep(int.Parse(Execution_Speed));
                                }
                                catch (Exception) { }

                            }

                            //Backup incompatible window handler
                            else if (item.Name.Contains("Backup Incompatible"))
                            {
                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "OK");
                                }
                                catch (Exception) { }
                            }

                            //Sync company file window handler
                            else if (item.Name.Contains("Sync Company File"))
                            {
                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Continue");
                                    Actions.WaitForAnyChildWindow(qbWindow, "Sync Company File", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }
                            }

                            //QB Information window handler
                            else if (item.Name.Contains("QuickBooks Information"))
                            {
                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "OK");
                                    Actions.WaitForAnyChildWindow(qbWindow, "QuickBooks Information", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }
                            }

                            //Create backup copy window handler
                            else if (item.Name.Contains("Create Backup"))
                            {
                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Next");
                                    Thread.Sleep(int.Parse(Execution_Speed));
                                }
                                catch (Exception) { }
                            }

                            //Backup options window handler - file path
                            else if (item.Name.Equals("Backup Options"))
                            {
                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.SetTextByAutomationID(item, "2002", TestDataLocalDirectory + "Test");
                                    MessageBox.Show("....");
                                }
                                catch (Exception) { }

                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Yes");
                                }
                                catch (Exception) { }

                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "OK");
                                }
                                catch (Exception) { }

                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks"), "Use this Location");
                                }
                                catch (Exception) { }

                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Save Backup Copy"), "Save");
                                    Actions.WaitForAnyChildWindow(qbWindow, "Save Backup Copy", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }
                            }

                            //Quickbooks use this location window handler
                            else if (item.Name.Contains("QuickBooks"))
                            {
                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks"), "Use this Location");
                                    Actions.WaitForAnyChildWindow(qbWindow, "QuickBooks", int.Parse(Sync_Timeout));
                                }
                                catch (Exception)
                                {
                                }
                            }

                            //Save backup copy window handler
                            else if (item.Name.Contains("Save Backup Copy"))
                            {
                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Save Backup Copy"), "Save");
                                    Actions.WaitForAnyChildWindow(qbWindow, "Save Backup Copy", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }

                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Update Company"), "Yes");
                                    Actions.WaitForAnyChildWindow(qbWindow, "Update Company", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }
                            }

                            //Update company window handler
                            else if (item.Name.Contains("Update Company"))
                            {
                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Yes");
                                    Actions.WaitForAnyChildWindow(qbWindow, "Update Company", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }

                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Continue");
                                    Actions.WaitForAnyChildWindow(qbWindow, "Update Company", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }
                            }

                            //Enter email address window handler
                            else if (item.Name.Contains("Enter your email address"))
                            {
                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Encountered a Problem"), "Skip");
                                }
                                catch (Exception) { }

                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Close");
                                }
                                catch (Exception) { }
                            }

                            else if (item.Name.Contains("Encountered a Problem"))
                            {
                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Skip");
                                }
                                catch (Exception) { }

                            }

                            //Warning window handler
                            else if (item.Name.Contains("Warning"))
                            {

                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "OK");
                                    Actions.WaitForAnyChildWindow(qbWindow, "Update Now", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }

                                try
                                {

                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "OK");
                                    Actions.WaitForAnyChildWindow(qbWindow, "Warning", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }

                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Continue");
                                    Actions.WaitForAnyChildWindow(qbWindow, "Warning", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }

                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Cancel");
                                    Actions.WaitForAnyChildWindow(qbWindow, "Warning", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }

                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Start");
                                    Actions.WaitForAnyChildWindow(qbWindow, "Warning", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }

                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks File Doctor"), "Continue");
                                }
                                catch (Exception) { }

                            }

                            //QuickBooks File Doctor window handler
                            else if (item.Name.Contains("QuickBooks File Doctor"))
                            {
                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Continue");
                                    Actions.WaitForAnyChildWindow(qbWindow, "QuickBooks File Doctor", int.Parse(Sync_Timeout));
                                }
                                catch (Exception) { }
                            }


                            //Home window handler
                            else if (item.Name.Contains("Home"))
                            {
                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Close");
                                    Thread.Sleep(int.Parse(Execution_Speed));
                                }
                                catch (Exception) { }
                            }

                            //Enter memorized transaction window handler
                            else if (item.Name.Contains("Enter Memorized Transactions"))
                            {
                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "Enter All Later");
                                    Thread.Sleep(int.Parse(Execution_Speed));
                                }
                                catch { }

                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "OK");
                                    Thread.Sleep(int.Parse(Execution_Speed));
                                }
                                catch { }
                            }

                            //Enter memorized transaction window handler
                            else if (item.Name.Contains("Enter Memorized Transactions"))
                            {
                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "OK");
                                    Thread.Sleep(int.Parse(Execution_Speed));
                                }
                                catch { }
                            }

                            //Insights works on accural basis window handler
                            else if (item.Name.Contains("Insights works on the accrual basis only"))
                            {
                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    Actions.ClickElementByName(item, "OK");
                                    Thread.Sleep(int.Parse(Execution_Speed));
                                }
                                catch { }
                            }

                            //Alert window handler
                            else
                            {
                                //Alert window handler
                                if (Actions.CheckDesktopWindowExists("Alert"))
                                    Actions.CheckForAlertAndClose("Alert");

                                //Crash handler
                                if (Actions.CheckDesktopWindowExists("QuickBooks - Unrecoverable Error"))
                                {
                                    Actions.QBCrashHandler();
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        ResetQBWindows(qbApp, qbWindow, false);
                        break;
                    }
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
                while (modalWin.Count != 0);
                Thread.Sleep(int.Parse(Execution_Speed));

                Logger.logMessage("OpenOrUpgradeCompanyFile " + companyFilePath + "->" + " - End");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("OpenOrUpgradeCompanyFile " + companyFilePath + "->" + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");

                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static void ResetQBWindows(TestStack.White.Application qbApp, Window qbWin, bool openFileOnNoCompany)
        {

            Logger.logMessage("                 ResetQBWindows " + " - Begin");

            List<Window> modalWin = null;
            int iteration = 0;
            bool menuEnabled = false;

            try
            {
                do
                {
                    try
                    {
                        Logger.logMessage("---------------Try-Catch Block------------------------"); 
                        Actions.SelectMenu(qbApp, qbWin, "Window", "Close All"); 
                    }
                    catch (Exception) { }

                    do
                    {
                        //Alert window handler
                        if (Actions.CheckDesktopWindowExists("Alert"))
                            Actions.CheckForAlertAndClose("Alert");

                        //Crash handler
                        if (Actions.CheckDesktopWindowExists("QuickBooks - Unrecoverable Error"))
                        {
                            Actions.QBCrashHandler();
                            break;
                        }

                        if (iteration <= 10)
                        {
                            iteration = iteration + 1;
                            modalWin = qbWin.ModalWindows();

                            foreach (Window item in modalWin)
                            {
                                //Alert window handler
                                if (Actions.CheckDesktopWindowExists("Alert"))
                                    Actions.CheckForAlertAndClose("Alert");

                                //Crash handler
                                if (Actions.CheckDesktopWindowExists("QuickBooks - Unrecoverable Error"))
                                {
                                    Actions.QBCrashHandler();
                                    break;
                                }

                                try
                                {
                                    Logger.logMessage("---------------Try-Catch Block------------------------");
                                    if (Actions.CheckMenuEnabled(qbApp, qbWin, "File"))
                                    {
                                        menuEnabled = true;
                                        break;
                                    }
                                }
                                catch (Exception)
                                { }

                                 //Enter memorize report window handler
                                if (item.Name.Contains("Memorize Report"))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "No");
                                        Thread.Sleep(int.Parse(Execution_Speed));
                                        break;
                                    }
                                    catch { }
                                }

                                //Handle Save commented report popup
                                if (item.Name.Contains("Save Your Commented Report?"))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "No");
                                        Thread.Sleep(int.Parse(Execution_Speed));
                                        break;
                                    }
                                    catch { }
                                }
 
                                //Register QB window handler
                                if (item.Name.Contains("Register QuickBooks"))
                                {
                                    try
                                    {

                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "Remind Me Later");
                                        //Actions.WaitForAnyChildWindow(qbWin, item.Name, int.Parse(Sync_Timeout));
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                        break;
                                    }
                                    catch { }
                                }

                                //Admin permission needed window handler
                                if (item.Name.Contains("Administrator Permissions Needed"))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "Continue");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                        break;
                                    }
                                    catch { }
                                }

                                //No company window handler
                                if (item.Name.Contains("No") && openFileOnNoCompany.Equals(true))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        QuickBooks.OpenOrUpgradeCompanyFile(PathBuilder.GetPath("DefaultCompanyFile.qbw"), qbApp, qbWin, false, false);
                                    }
                                    catch { }
                                }

                                //Update quickbooks window handler
                                if (item.Name.Contains("Update QuickBooks"))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "Close");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                        break;
                                    }
                                    catch { }
                                }

                                //Payroll update window handler
                                if (item.Name.Equals("Payroll Update"))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "Cancel");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                        break;
                                    }
                                    catch { }

                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "OK");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                        break;
                                    }
                                    catch { }

                                }

                                //Intuit payroll services window hadler
                                if (item.Name.Contains("Intuit Payroll Services"))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "OK");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                        break;
                                    }
                                    catch { }
                                }

                                //Employer services window handler
                                if (item.Name.Contains("Employer Services"))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "Cancel");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                        break;
                                    }
                                    catch { }
                                }

                                //Insights works on the accrual basis window handler
                                if (item.Name.Equals("Insights Works on Accrual Basis Only"))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "OK");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                        break;
                                    }
                                    catch { }
                                }

                                //Insights works on the accrual basis window handler
                                if (item.Name.Contains("Insights"))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "OK");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                        break;
                                    }
                                    catch { }
                                }

                                //Enter memorized transactions window handler
                                if (item.Name.Contains("Enter Memorized Transactions"))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(Actions.GetChildWindow(qbWin, "Enter Memorized Transactions"), "Enter All Later");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                    }
                                    catch { }

                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(Actions.GetChildWindow(qbWin, "Enter Memorized Transactions"), "OK");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                        break;
                                    }
                                    catch { }
                                }

                                //Recording transaction window handler
                                if (item.Name.Contains("Recording Transaction"))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "No");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                        break;
                                    }
                                    catch { }
                                }


                                //Login window handler
                                if (item.Name.Equals("QuickBooks Login"))
                                {
                                    Actions.SetFocusOnWindow(item);
                                    Actions.SendBCKSPACEToWindow(item);
                                    Actions.SetTextByAutomationID(item, "15922", UserName);
                                    Actions.SendTABToWindow(item);
                                    Actions.SendKeysToWindow(item, Password);
                                    Actions.ClickElementByAutomationID(item, "51");
                                    Actions.WaitForAnyChildWindow(qbWin, "QuickBooks Login", int.Parse(Sync_Timeout));
                                    Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                    break;
                                }

                                //Error window handler
                                if (item.Name.Contains("Error"))
                                {
                                    Actions.ClickElementByName(item, "Don't Send");
                                    Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                    break;
                                }


                                //QB Setup window handler
                                if (item.Name.Contains("Setup"))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "Close");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                    }
                                    catch (Exception)
                                    { }

                                    try
                                    {

                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "Yes");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                        break;
                                    }
                                    catch (Exception)
                                    { }

                                }

                                //Warning window handler
                                if (item.Name.Contains("Warning"))
                                {
                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "OK");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                    }
                                    catch (Exception)
                                    { }

                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "Cancel");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                        break;
                                    }
                                    catch (Exception)
                                    { }

                                }

                                else
                                {
                                    item.Focus();

                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(Actions.GetChildWindow(qbWin, "Recording Transaction"), "No");
                                    }
                                    catch { }

                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "Close");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                    }
                                    catch { }

                                    try
                                    {

                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "No");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                    }
                                    catch { }


                                    try
                                    {
                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        item.Close();
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                    }
                                    catch { }

                                    try
                                    {

                                        Logger.logMessage("---------------Try-Catch Block------------------------");
                                        Actions.ClickElementByName(item, "OK");
                                        Thread.Sleep(int.Parse(ResetWindow_Timeout));
                                    }
                                    catch { }
                                }
                            }
                            Thread.Sleep(int.Parse(Execution_Speed));
                        }
                        else
                        {
                            break;
                        }
                    }
                    while (modalWin.Count != 0 && menuEnabled.Equals(false));
                    Thread.Sleep(int.Parse(Execution_Speed));
                }
                while (!Actions.CheckMenuEnabled(qbApp, qbWin, "File"));

                Logger.logMessage("                 ResetQBWindows " + " - End");
                Logger.logMessage("------------------------------------------------------------------------------");
            }

            catch (Exception e)
            {
                Logger.logMessage("ResetQBWindows " + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************

        public static void CreateCompany(TestStack.White.Application qbApp, Window qbWindow, string businessName, string industry)
        {
            try
            {
                Actions.SelectMenu(qbApp, qbWindow, "File", "New Company...");
                Actions.WaitForChildWindow(qbWindow, "QuickBooks Setup", int.Parse(Sync_Timeout));

                Window QBSetupWindow = Actions.GetChildWindow(qbWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.QBSetup_Window_Name);
                Thread.Sleep(int.Parse(Execution_Speed));

                Actions.ClickElementByAutomationID(QBSetupWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.ExpressStart_Button_AutoID);
                Thread.Sleep(int.Parse(Execution_Speed));

                if (Actions.CheckElementExistsByAutomationID(QBSetupWindow, "txt_LoginEmail"))
                {
                    if (!Actions.CheckWindowExists(qbWindow, "Encountered a Problem"))
                    {
                        Actions.SetTextByAutomationID(QBSetupWindow, "txt_LoginEmail", businessName + "@hotmail.com");
                        Actions.ClickElementByAutomationID(QBSetupWindow, "btn_Next");
                        Actions.SetTextByAutomationID(QBSetupWindow, "pwd_NewPwd", "Intuit01");
                        Actions.SetTextByAutomationID(QBSetupWindow, "pwd_NewConfirm", "Intuit01");
                        Actions.SetTextByAutomationID(QBSetupWindow, "txt_FirstName", businessName);
                        Actions.SetTextByAutomationID(QBSetupWindow, "txt_LastName", "Test");
                        var uiaWindow = Actions.UIA_GetAppWindow(qbWindow.Name);
                        Actions.SetFocusOnElementByAutomationID(QBSetupWindow, "btn_Continue");
                        Actions.SendTABToWindow(QBSetupWindow);
                        Actions.SendTABToWindow(QBSetupWindow);
                        Actions.SendTABToWindow(QBSetupWindow);
                        Actions.SendTABToWindow(QBSetupWindow);
                        Actions.SendTABToWindow(QBSetupWindow);
                        Actions.SendTABToWindow(QBSetupWindow);
                        Actions.SendENTERoWindow(QBSetupWindow);
                    }

                }
                try
                {
                    Logger.logMessage("---------------Try-Catch Block------------------------");
                    Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Encountered a Problem"), "Skip");
                }
                catch (Exception) { }

                Actions.SetTextByAutomationID(QBSetupWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.BusinessName_TxtField_AutoID, businessName);
                Actions.SetTextByAutomationID(QBSetupWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.IndustryList_TxtField_AutoID, "Information");
                Thread.Sleep(int.Parse(Execution_Speed));
                Actions.SelectListBoxItemByText(QBSetupWindow, "lstBox_Industry", "Information Technology");
                Actions.SelectComboBoxItemByText(QBSetupWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.TaxStructure_CmbBox_AutoID, "Corporation");
                Actions.SetTextByAutomationID(QBSetupWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.TaxID_TxtField_AutoID, "123-45-6789");
                Actions.SelectComboBoxItemByText(QBSetupWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.HaveEmployees_CmbBox_AutoID, "No");
                Actions.ClickElementByAutomationID(QBSetupWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.Continue_Button_AutoID);
                Thread.Sleep(int.Parse(Execution_Speed));
                Actions.SelectComboBoxItemByText(QBSetupWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.StateName_CmbBox_AutoID, "DE");
                Actions.SetTextByAutomationID(QBSetupWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.ZipCode_TxtField_AutoID, "DE123");
                Actions.SetTextByAutomationID(QBSetupWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.Phone_TxtField_AutoID, "6104567890");
                Actions.ClickElementByAutomationID(QBSetupWindow, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.CreateCompany_Button_AutoID);
                Actions.WaitForAnyChildWindow(qbWindow, QBSetupWindow.Name, int.Parse(Sync_Timeout));
            }
            catch (Exception e)
            {
                String sMessage = e.Message;
                LastException.SetLastError(sMessage);
                throw new Exception(sMessage);
            }
        }

        //**************************************************************************************************************************************************************


    }

}
