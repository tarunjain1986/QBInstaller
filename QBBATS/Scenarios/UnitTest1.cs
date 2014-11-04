using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.Utils;
using FrameworkLibraries.ActionLibs.WhiteAPI;

namespace QBBATS.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            Logger log = new Logger("Training");
            Random rand = new Random();

            var qbApp = QuickBooks.GetApp("QuickBooks");
            var qbWindow = QuickBooks.GetAppWindow(qbApp, "QuickBooks");

            Actions.SelectMenu(qbApp, qbWindow, "File", "New Company...");

            Actions.WaitForChildWindow(qbWindow, "QuickBooks Setup", 999999);

            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "QuickBooks Setup"), "Advanced Setup");

            Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "2201", "WhiteTest_"+rand.Next(123, 456));
            Actions.SendTABToWindow(Actions.GetChildWindow(qbWindow, "EasyStep Interview"));
            Actions.SendTABToWindow(Actions.GetChildWindow(qbWindow, "EasyStep Interview"));
            Actions.SendTABToWindow(Actions.GetChildWindow(qbWindow, "EasyStep Interview"));
            Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "2203", "WhiteTest Address");
            Actions.SendTABToWindow(Actions.GetChildWindow(qbWindow, "EasyStep Interview"));
            Actions.SendTABToWindow(Actions.GetChildWindow(qbWindow, "EasyStep Interview"));
            Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "2205", "Delaware");
            Actions.SendTABToWindow(Actions.GetChildWindow(qbWindow, "EasyStep Interview"));
            Actions.SendTABToWindow(Actions.GetChildWindow(qbWindow, "EasyStep Interview"));
            Actions.SendTABToWindow(Actions.GetChildWindow(qbWindow, "EasyStep Interview"));
            Actions.SendTABToWindow(Actions.GetChildWindow(qbWindow, "EasyStep Interview"));
            Actions.SetTextByAutomationID(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "2209", "6104567890");
            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "Next >");
            Actions.WaitForElementEnabled(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "Next >", 99999);
            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "Next >");
            Actions.ClickElementByAutomationID(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "2318");
            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "Next >");
            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "Next >");
            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "Next >");
            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "Next >");
            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "Filename for New Company"), "Save");
            try { Actions.WaitForElementEnabled(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "Next >", 99999); }
            catch { }             
            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "Leave...");
            Actions.ClickElementByName(Actions.GetChildWindow(qbWindow, "EasyStep Interview"), "OK");

        }
    }
}
