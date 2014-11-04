using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestStack.White.UIItems.WindowItems;

namespace PaymentsAutomation.UseCases.MAS
{
    public class Constants
    {
        public static Logger log;
        public static TestStack.White.Application qbApp = null;
        public static TestStack.White.UIItems.WindowItems.Window qbWindow = null;
        public static String startupPath = System.IO.Path.GetFullPath("..\\..\\..\\");
        public static Property conf = Property.GetPropertyInstance();
        public static string exe = conf.get("QBExePath");
        public string qbLoginUserName = conf.get("QBLoginUserName");
        public string qbLoginPassword = conf.get("QBLoginPassword");
        public static string Execution_Speed = conf.get("ExecutionSpeed");
        public static string Sync_Timeout = conf.get("SyncTimeOut");
        public Random rand = new Random();
        public string testName = "ReceivePayments";
        public string moduleName = "PaymentAutomation";
        public string exception = "Null";
        public string category = "Null";

        public static void Initialize(string logger)
        {
            log = new Logger(logger);
            exe = conf.get("QBExePath");
            qbApp = QuickBooks.Initialize(exe);
            qbWindow = QuickBooks.PrepareBaseState(qbApp);
        }
    }
}
