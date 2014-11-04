using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries.AppLibs.QBDT;
using System.Threading;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.Payments;
using TestStack.BDDfy;

namespace PaymentsAutomation.UseCases.MAS
{
    public class Refund
    {
        Logger log = new Logger("Refund.txt");
        public TestStack.White.Application qbApp = null;
        public TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static String startupPath = System.IO.Path.GetFullPath("..\\..\\..\\");
        public static Property conf = Property.GetPropertyInstance();
        public string exe = conf.get("QBExePath");
        public string qbLoginUserName = conf.get("QBLoginUserName");
        public string qbLoginPassword = conf.get("QBLoginPassword");
        public static string Execution_Speed = conf.get("ExecutionSpeed");
        public static string Sync_Timeout = conf.get("SyncTimeOut");
        public Random rand = new Random();
        public string testName = "ProcessRefund";
        public string moduleName = "PaymentAutomation";
        public string exception = "Null";
        public string category = "Null";

        [Given(StepTitle = "QuickBooks Desktop should be up and running")]
        public void Setup()
        {
            exe = conf.get("QBExePath");
            qbApp = QuickBooks.Initialize(exe);
            qbWindow = QuickBooks.PrepareBaseState(qbApp);
        }
        [When(StepTitle = "Refund for a transaction should be processed successfully")]
        public void ProcessRefund()
        {
            Payments.RefundPayment(qbApp, qbWindow,"item1", "10", "4111111111111111", "09", "2025", "Cust1", "", "", "12345");
        }
    }
}
