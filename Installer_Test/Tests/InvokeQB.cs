using System;
using System.Windows.Forms;
using Xunit;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Installer_Test.Lib;
using System.Linq;


using Microsoft.Win32;



using Excel = Microsoft.Office.Interop.Excel;

namespace Installer_Test.Tests
{

    public class InvokeQB
    {

        public string testName = "InvokeQB";
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        Dictionary<String, String> dic;


        [Given(StepTitle = "Given - QuickBooks App and Window instances are available")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
            string readpath = "C:\\Temp\\Parameters.xlsx";
            List<string> listHeader = new List<string>();
            List<string> ListValue = new List<string>();
            dic = new Dictionary<string, string>();
            File_Functions.ReadExcelSheet(readpath, "InvokeQB", 1, ref listHeader);
            File_Functions.ReadExcelSheet(readpath, "InvokeQB", 2, ref ListValue);
            dic = listHeader.Zip(ListValue, (k, v) => new { k, v })
                 .ToDictionary(x => x.k, x => x.v);
           
        }
        [Then(StepTitle = "Then - InvokeQB")]
        public void Invoke_QB()
        {
           Install_Functions.InvokeQB(dic);
           
                

        }
        [Fact]
        public void Run_InvokeQB()
        {
            this.BDDfy();
        }
    }
}
