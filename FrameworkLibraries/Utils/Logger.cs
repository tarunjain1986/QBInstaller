using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameworkLibraries.Utils
{
    public class Logger
    {
        public static string logFile = null;
        public static string logFilePath = null;

        public Logger(string logFileName)
        {
            Utils.BDDfyReporterConfig.ConfigReport();
            logFile = logFileName;
            var strBuilder = new StringBuilder();
            var logDir = Property.GetPropertyInstance().get("LogDirectory");

            if (!Directory.Exists(logDir))
                Directory.CreateDirectory(logDir);

            logFilePath = strBuilder.Append(logDir).ToString();
            logFilePath = strBuilder.Append(logFile).ToString();

            if (File.Exists(logFilePath))
                File.Delete(logFilePath);
        }

        public static void logMessage(string msg)
        {
            using (StreamWriter w = File.AppendText(logFilePath))
            {
                w.WriteLine(msg);
                w.Close();
            }
        }
    }
}
