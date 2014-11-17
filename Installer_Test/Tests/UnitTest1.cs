using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FrameworkLibraries.ActionLibs.WhiteAPI;

namespace Installer_Test.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            FrameworkLibraries.Utils.Logger log = new FrameworkLibraries.Utils.Logger("Enabled test");

            var app = Actions.GetApp("QuickBooks", "QBW32");

            var win = Actions.GetAppWindow(app, "QuickBooks");

            var child = Actions.GetChildWindow(win, "Select QuickBooks Industry-Specific Edition");

            var e = Actions.CheckElementIsEnabled(child, "Enterprise Solutions ");


        }
    }
}
