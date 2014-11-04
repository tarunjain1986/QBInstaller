using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.ActionLibs;
using FrameworkLibraries.AppLibs.QBDT;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;

using Xunit;

using Installer_Test;


namespace Installer_Test.Tests
{
   
    public class Installer
    {
       // public TestStack.White.Application qbApp = null;
       // public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        public static string testName = "Install";
        public string targetPath, installPath, fileName, wkflow, customOpt, License_No, Product_No, UserID, Passwd, firstName, lastName;
        string [] LicenseNo, ProductNo;


        [Given(StepTitle = @"The parameters for installation are available at C:\Installation\Parameters.txt")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
      
            //////////////////////////////////////////////////////////////////////////
            // Following code is for reading from text file
            //////////////////////////////////////////////////////////////////////////
            //string readpath = @"C:\Temp\Parameters.txt";
            //File.WriteAllLines(readpath, File.ReadAllLines(readpath).Where(l => !string.IsNullOrWhiteSpace(l))); // Remove white space from the file

            //string[] lines = File.ReadAllLines(readpath);
            //var dic = lines.Select(line => line.Split('=')).ToDictionary(keyValue => keyValue[0], bits => bits[1]);

            //targetPath = dic["Target Path"];
            //wkflow = dic["Workflow"];
            //customOpt = dic["Installation Type"];
            //License_No = dic["License No"];
            //Product_No = dic["Product No"];
            //UserID = dic["UserID"];
            //Passwd = dic["Password"];
            //firstName = dic["First Name"];
            //lastName = dic["Last Name"];

            //////////////////////////////////////////////////////////////////////////////////////////////
            // The following code is for reading from an excel file
            //////////////////////////////////////////////////////////////////////////////////////////////

            string readpath = "C:\\Temp\\Parameters.xlsx"; // "C:\\Installation\\Sample.txt";

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(readpath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Path");
            Excel.Range xlRng = (Excel.Range)xlWorkSheet.get_Range("B2:B4", Type.Missing);

            Dictionary<string, string> dic = new Dictionary<string, string>();

            foreach (Excel.Range cell in xlRng)
            {

                string cellIndex = cell.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                string cellValue = Convert.ToString(cell.Value2);
                dic.Add(cellIndex, cellValue);

            }
            
            targetPath = dic["B2"];
            installPath = dic["B9"];
            //wkflow = dic["Workflow"];
            //customOpt = dic["Installation Type"];
            //License_No = dic["License No"];
            //Product_No = dic["Product No"];
            //UserID = dic["UserID"];
            //Passwd = dic["Password"];
            //firstName = dic["First Name"];
            //lastName = dic["Last Name"];


            var regex = new Regex(@".{4}");
            string temp = regex.Replace(License_No, "$&" + "\n");
            LicenseNo = temp.Split('\n');

            regex = new Regex(@".{3}");
            temp = regex.Replace(Product_No, "$&" + "\n");
            ProductNo = temp.Split('\n');
            
        }

        [Then(StepTitle = "Then - Invoke QuickBooks installer")]
        public void InvokeQB()
        {
           OSOperations.InvokeInstaller(targetPath, "setup.exe");
        }


        [AndThen(StepTitle = "Then - Install QuickBooks")]
        public void RunInstallQB()
        {
            Install_Functions.Install_QB(targetPath, wkflow, customOpt, LicenseNo, ProductNo, UserID, Passwd, firstName, lastName, installPath);
        
        }

       [Fact]
        public void RunQBInstallTest()
        {
            this.BDDfy();
        }
    }
}
