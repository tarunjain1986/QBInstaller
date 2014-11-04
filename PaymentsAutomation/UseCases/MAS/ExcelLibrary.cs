using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems;
using Xunit;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using PaymentsAutomation.Stories.MAS;
using Microsoft.Office.Interop.Excel;

namespace PaymentsAutomation.UseCases.MAS
{
    //[TestClass]
    public class ExcelLibrary
    {
        public static string CCNumber, testType;
        public static Property conf = Property.GetPropertyInstance();
        public static string Sync_Timeout = conf.get("SyncTimeOut");
        public static string readpath = conf.get("DataPath");
        public static string Test_Environment = conf.get("TestEnvironment");
        public static string QB_Build = conf.get("QBBuild");
        public static string QB_ReleaseNumber = conf.get("QBReleaseNumber");
        public static string QB_BuildType = conf.get("QBBuildType");
        public static string QBLocalPath, QBVersionName;
        public string line, workFlow, QBVersion, Build, SKU, CustomOpt;
        public static string sourcePath, targetPath, fileName, sourceFile, destFile, Row, Column;
        public static string CCExpDate, CCExpYear, CCZipCode, CCPaymentAmount, CCCustName, CCType;
        public static object[,] valueArray;
        public static int lastUsedRow, lastUsedCloumn;
        public static Dictionary<string, string> dic = null;

        //string Product_No, UserID, Passwd, firstName, lastName;
        [Fact]
        public static void setUpEnvironment()
        {
            if (QB_Build.Equals("MangoPro"))
            {
                QBLocalPath = "QuickBooks 2015";
                QBVersionName = "Mango";
            }
            else if ((QB_Build.Equals("MangoEnt")))
            {
                QBLocalPath = "QuickBooks Enterprise Solutions 15.0";
                QBVersionName = "Mango";
            }
            if (File.Exists("C:\\Program Files (x86)\\Intuit\\"+QBLocalPath+"\\Data\\SBSSettings.dat"))
            {
                Console.WriteLine("The file exists.");
                File.Move("C:\\Program Files (x86)\\Intuit\\" + QBLocalPath + "\\Data\\SBSSettings.dat", "C:\\Program Files (x86)\\Intuit\\" + QBLocalPath + "\\Data\\Old"+DateTime.Now.Ticks+".dat");
                string src = "\\\\banfsalab02\\Users\\Nitish\\dat files\\" + QBVersionName + QB_ReleaseNumber + "\\" + QB_BuildType + "\\" + Test_Environment + "\\SBSSettings.dat";
                string dest = @"C:\Program Files (x86)\Intuit\QuickBooks Enterprise Solutions 15.0\Data\SBSSettings.dat";
                File.Copy(@"\\banfsalab02\Users\Nitish\dat files\MangoR3\release\PTC\SBSSettings.dat", "C:\\" + "Program" + " " + "Files" + " " + "(x86)" + "\\Intuit\\" + "QuickBooks" + " " + "Enterprise" + " " + "Solutions" + " " + "15.0" + "\\Data\\SBSSettings.dat");
                //File.Copy();
                Thread.Sleep(1000);
            }

        }
        
        public static void readExcel()
        {

            // Reading Entire Excel
            Logger log = new Logger("ExcelRead.txt");
            /**************************************The following code is for reading from an excel file******************************************/
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(readpath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("MAS");
            Excel.Range xlRange = xlWorkSheet.UsedRange;
            lastUsedRow = xlWorkSheet.UsedRange.Rows.Count;
            lastUsedCloumn = xlWorkSheet.UsedRange.Columns.Count;
            valueArray = (object[,])xlRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
            //return valueArray;  
        }
        
    }
}
            

