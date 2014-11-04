using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;

namespace WhiteTraining
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            Property conf = Property.GetPropertyInstance();
            string exe = conf.get("QBExePath");
            Logger log = new Logger("Training");

            var qbApp = QuickBooks.GetApp("QuickBooks");
            var qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");

            Actions.WaitForChildWindow(qbWindow, "Profit  Loss by Job", 9999999);

            var repWindow = Actions.GetChildWindow(qbWindow, "Profit  Loss by Job");

            try { Actions.ClickElementByName(repWindow, "Comment on Report"); }
            catch { }


            Actions.WaitForChildWindow(qbWindow, "Comment on Report: Profit  Loss by Job", 9999999);

            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Comment on Report: Profit  Loss by Job"), "Save");

            try { Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Save Your Commented Report"), "OK"); }
            catch { }

            Actions.WaitForChildWindow(qbWindow, "Saved Successfully", 9999999);

            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Saved Successfully"), "OK");

            bool s = Actions.CheckElementExistsByName(Actions.GetChildWindow(qbWindow, "Comment on Report: Profit  Loss by Job"), "helloworld..!!");

            




        }
    }
}
