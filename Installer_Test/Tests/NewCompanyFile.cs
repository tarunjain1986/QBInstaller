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

   public class NewCompanyFile
    {

        public string testName = "NewCompanyFile";
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public static string exe = conf.get("QBExePath");
        public string bizName,industryList,industryType,businessType,address1,address2,state,country,zip,phone,city;
        Dictionary<String, String> keyvaluepairdic;

        [Given(StepTitle = "Given - QuickBooks App and Window instances are available")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
            QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
            string readpath = "C:\\Temp\\Parameters.xlsx";
            List<string> listHeader = new List<string>();
            List<string> ListValue = new List<string>();
            keyvaluepairdic = new Dictionary<string, string>();
            File_Functions.ReadExcelSheet(readpath, "CompanyFile", 1, ref listHeader);
            File_Functions.ReadExcelSheet(readpath, "CompanyFile", 3, ref ListValue);
            keyvaluepairdic = listHeader.Zip(ListValue, (k, v) => new { k, v })
                 .ToDictionary(x => x.k, x => x.v);
   
        }

        [Then(StepTitle = "Then - Create Company File")]
        public void CreateCompanyFile()
        {

           Install_Functions.CreateCompanyFile(keyvaluepairdic);

        }
        [Fact]
        public void Run_NewCompanyFile()
        {
            this.BDDfy();
        }
    }
}
