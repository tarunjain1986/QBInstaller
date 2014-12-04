using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries.Utils;
using TestStack.White;
using TestStack.White.UIItems.WindowItems;

namespace Installer_Test.Tests
{
    [TestClass]
    public class UnitTest2
    {
        [TestMethod]
        public void TestMethod1()
        {
            Logger log = new Logger("Test");


            Window qb_install1 = Actions.GetDesktopWindow("QuickBooks Installation");
            TestStack.White.UIItems.Panel Pane2 = Actions.GetPaneByName(qb_install1, "Intuit QuickBooks Installer");

            try { Actions.ClickButtonInsidePanelByName(qb_install1, Pane2, "Finish"); }
            catch { }

            try { Actions.WaitForWindow("QuickBooks Installation", 50000); }
            catch { }

            Window qb_install2 = Actions.GetDesktopWindow("QuickBooks Installation");
            TestStack.White.UIItems.Panel Pane3 = Actions.GetPaneByName(qb_install2, "Intuit QuickBooks Installer");

            var win = Actions.GetDesktopWindow("Intuit QuickBooks Installer");

            Actions.ClickElementByName(win, "Open QuickBooks");

            Actions.WaitForWindow("Select QuickBooks Industry-Specific Edition", 50000);

            var win2 = Actions.GetDesktopWindow("Select QuickBooks Industry-Specific Edition");

            try { Actions.ClickButtonInsidePanelByName(qb_install2, Pane3, "Open QuickBooks"); }
            catch { }

            try { Actions.WaitForWindow("Select QuickBooks Industry-Specific Edition", 50000); }
            catch { }


        }
    }
}
