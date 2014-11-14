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

namespace AutoPatch
{
    
    public class AutoPatch
    {
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public Random rand = new Random();
        public string testName = "AutoPatch";

        [Given(StepTitle = "Given - QuickBooks App and Window instances are available")]
        public void Method()
        {
            string exe = conf.get("QBExePath");
            var qbApp = QuickBooks.Initialize(exe);
            var qbWindow = QuickBooks.PrepareBaseState(qbApp);

            Actions.SelectMenu(qbApp, qbWindow, "Help", "Update QuickBooks...");
           


        }
    }
}
