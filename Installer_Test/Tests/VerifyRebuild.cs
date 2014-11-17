using System;
using FrameworkLibraries.Utils;
using System.Windows.Automation;
using System.Windows.Forms;
using FrameworkLibraries.ActionLibs;
using TestStack.White.UIItems.WindowItems;
using System.Threading;
using TestStack.White.UIItems.Finders;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries;
using System.Collections.Generic;
using TestStack.White.UIItems;
using Xunit;
using TestStack.BDDfy;
using FrameworkLibraries.AppLibs.QBDT;
using System.IO;
using System.Reflection;

namespace VerifyRebuild
{
    
    public class VerifyRebuild
    {
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public Random rand = new Random();
        public string testName = "VerifyRebuild";

        [Given(StepTitle = "Given - QuickBooks App and Window instances are available")]
        public void Method()
        {
            string exe = conf.get("QBExePath");
            var qbApp = QuickBooks.Initialize(exe);
            var qbWindow = QuickBooks.PrepareBaseState(qbApp);


            Actions.SelectMenu(qbApp, qbWindow, "File", "Utilities", "Verify Data");

            if (qbWindow.Title.Equals("QuicKbooks Informartion"))
            {
                Actions.ClickButtonByName(Actions.GetChildWindow(qbWindow, "QuicKbooks Informartion"), "OK");

            }
            else
            {
                Actions.ClickButtonByName(Actions.GetChildWindow(qbWindow, "QuicKbooks Informartion"), "OK");
            }

            Actions.SelectMenu(qbApp, qbWindow, "File", "Utilities", "Rebuild Data");

            if (qbWindow.Title.Equals("QuicKbooks Informartion"))
            {
                Actions.ClickButtonByAutomationID(Actions.GetChildWindow(qbWindow, "Save backup Copy"), "1");

            }

            else
            {
                Actions.ClickButtonByName(Actions.GetChildWindow(qbWindow, "QuicKbooks Informartion"), "OK");
            }


        }

        [Fact]
        public void verifyBuild()
        {
            this.BDDfy();
        }
        
        
    }
}