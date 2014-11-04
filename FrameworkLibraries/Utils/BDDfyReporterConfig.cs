using MayaConnected;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestStack.BDDfy.Configuration;
using TestStack.BDDfy.Reporters.Html;

namespace FrameworkLibraries.Utils
{
    public class BDDfyReporterConfig
    {
        public static void ConfigReport()
        {
            Configurator.BatchProcessors.Add(new HtmlReporter(new BDDfyCustomReporter()));
            Configurator.Processors.ConsoleReport.Enable();
            Configurator.BatchProcessors.DiagnosticsReport.Enable();
            Configurator.BatchProcessors.MarkDownReport.Enable();
            Configurator.BatchProcessors.HtmlReport.Enable();
        }
    }
}
