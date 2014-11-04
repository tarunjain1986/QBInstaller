using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries.Utils;
using System.Windows.Automation;
using FrameworkLibraries.AppLibs.Payments;
using TestStack.White.UIItems;

namespace WhiteTraining
{
    [TestClass]
    public class UnitTest2
    {
        [TestMethod]
        public void TestMethod1()
        {
            Property conf = Property.GetPropertyInstance();
            string exe = conf.get("QBExePath");
            Logger log = new Logger("Training");


            ////Get the app and window handle - Method 1 (Static way of getting the handles) - When QB or the app under test is Open
            //var qbApp = QuickBooks.GetApp("QuickBooks");
            //var qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");
            ////------------------------------------------------------------------------------------

            ////Dynamic methods
            //var quickbooksApp = QuickBooks.Initialize(exe);
            ////var quickbooksWindow = QuickBooks.PrepareBaseState(quickbooksApp);

            //var payWin = Actions.GetChildWindow(qbWindow, "Create Invoices - Accounts Receivable");
            //Actions.SendALT_KeyToWindow(payWin, "a");


            var fiddlerApp = Actions.LaunchApp("Fiddler", "Fiddler");
            var fiddlerWindow = Actions.GetAppWindow(fiddlerApp, "Fiddler");
            Actions.CloseAllChildWindows(fiddlerWindow);
            ListViewCells x = Payments.GetFiddlerStackTrace(fiddlerWindow);
            fiddlerApp.Close();




        }
    }
}
