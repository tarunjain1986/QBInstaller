using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Collections.Generic;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using TestStack.White.UIItems.WindowItems;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;

using Xunit;
using Xunit.Extensions;

using BATS.DATA;


namespace BATS.Tests
{
    public class OpenQbBackupFile
    {
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public string exe = conf.get("QBExePath");
        public static string companyFilePath = null;
        public static string companyFileName = null;
        public Random rand = new Random();
        public string testName = "OpenQbBackupFile";
        public static string TestDataSourceDirectory = conf.get("TestDataSourceDirectory");
        public static string TestDataLocalDirectory = conf.get("TestDataLocalDirectory");


        [Given(StepTitle = "All the test company files are successfully copied from the source location")]
        public void CopyFiles()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
            FileOperations.DeleteCompanyFileInDirectory(TestDataLocalDirectory, companyFileName);
            FileOperations.CopyCompanyFileToDirectory(TestDataSourceDirectory, TestDataLocalDirectory, companyFileName);
        }

        [AndGiven(StepTitle = "Given - QuickBooks App and Window instances are available")]
        public void Setup()
        {
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
            QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
        }


        [Then(StepTitle = "Then - A QB backup company file should be opened or upgraded successfully")]
        public void OpenBackupCompanyFile()
        {
            QuickBooks.OpenOrUpgradeCompanyFile(companyFilePath, qbApp, qbWindow, true, false);
            var expectedTitleOfNewCompany = FrameworkLibraries.Utils.StringFunctions.RemoveExtentionFromFileName(companyFileName);
            Actions.XunitAssertContains(expectedTitleOfNewCompany.ToUpper(), qbWindow.Title.ToUpper());
        }

        [AndThen(StepTitle = "AndThen - Perform tear down activities to ensure that there are no on-screen exceptions")]
        public void TearDown()
        {
            QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
        }

        [Theory]
        [Category("P4")]
        [PropertyData("TestData", PropertyType = typeof(OpenBackupFileTestDataSource))]
        public void RunOpenBackupCompanyFileTest(string fileName)
        {
            companyFileName = fileName;
            companyFilePath = TestDataLocalDirectory + fileName;
            companyFilePath = companyFilePath.Replace("\\\\", "\\");
            this.BDDfy();
        }

    }
}
