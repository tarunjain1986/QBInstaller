using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.White.UIItems.WindowItems;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries.Utils;

namespace Installer_Test.Tests
{
    [TestClass]
    public class UnitTest3
    {
        [TestMethod]
        public void TestMethod1()
        {

            Logger log = new Logger("Test"); 
            Window win1 = Actions.GetDesktopWindow("Product Configuration");
            Logger.logMessage(win1.ToString());
          // Window desk1 =  Actions.GetDesktopWindow("QuickBooks Product Configuration");
            Window win2 = Actions.GetChildWindow(win1, "QuickBooks Product Configuration");
            
            Logger.logMessage(win2.ToString());
            //Thread.Sleep(1000);
            
           // Actions.WaitForChildWindow(win1, "QuickBooks Product Configuration", 60000);
            
           // Actions.WaitForElementEnabled(win2, "No", 30000);
            Actions.ClickElementByName(win2, "No");
        }
    }
}
