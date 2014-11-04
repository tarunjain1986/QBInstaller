using System;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.Utils;
using FrameworkLibraries.ActionLibs;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White;
using System.Threading;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries.EntityFramework;
using Xunit;
using TestStack.BDDfy;
using QBBATS.Data;
using System.Windows.Forms;
using System.IO;


namespace BATS.Tests
{
    
    public class CreateInvoice
    {
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static Property conf = Property.GetPropertyInstance();
        public string exe = conf.get("QBExePath");
        public Random rand = new Random();
        public int invoiceNumber, poNumber;
        public string testName = "CreateInvoice";
        public static string TestDataSourceDirectory = conf.get("TestDataSourceDirectory");
        public static string TestDataLocalDirectory = conf.get("TestDataLocalDirectory");
        public static string DefaultCompanyFile = conf.get("DefaultCompanyFile");
        public static string DefaultCompanyFilePath = DefaultCompanyFile;

        [Given(StepTitle = "Given - QuickBooks App and Window instances are available")]
        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
            qbApp = FrameworkLibraries.AppLibs.QBDT.QuickBooks.Initialize(exe);
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
            QuickBooks.ResetQBWindows(qbApp, qbWindow, true);
            invoiceNumber = rand.Next(12345, 99999);
            poNumber = rand.Next(12345, 99999);
        }

        [When(StepTitle = "When - A company file is opened or upgraded successfully for creating a transaction")]
        public void OpenCompanyFile()
        {

            if (!qbWindow.Title.Contains("Falcon"))
            {
                QuickBooks.OpenOrUpgradeCompanyFile(PathBuilder.GetPath("DefaultCompanyFile.qbw"), qbApp, qbWindow, false, false);
            }
        }

        [Then(StepTitle = "Then - An Invoice should be created successfully")]
        public void CreateInvoiceTest()
        {
            var customer = Invoice.Default.Customer_Job;
            var clss = Invoice.Default.Class;
            var account = Invoice.Default.Account;
            var template = Invoice.Default.Template;
            var rep = Invoice.Default.REP;
            var fob = Invoice.Default.FOB;
            var via = Invoice.Default.VIA;
            var item = Invoice.Default.Item;
            var quantity = Invoice.Default.Quantity;
            var itemDescription = "QuickBooks BATS";

            FrameworkLibraries.AppLibs.QBDT.QuickBooks.CreateInvoice(qbApp, qbWindow, customer, clss, account, template, invoiceNumber,
                poNumber, rep, via, fob, quantity, item, itemDescription, false);
        }

        [AndThen(StepTitle = "AndThen - Perform tear down activities to ensure that there are no on-screen exceptions")]
        public void TearDown()
        {
            QuickBooks.ResetQBWindows(qbApp, qbWindow, false);
        }

        [Fact]
        [Category("P2")]
        public void RunCreateInovoiceTest()
        {
            this.BDDfy();
        }
    }
}
