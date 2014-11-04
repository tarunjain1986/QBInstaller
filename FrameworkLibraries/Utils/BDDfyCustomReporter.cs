using FrameworkLibraries.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestStack.BDDfy.Reporters.Html;

namespace MayaConnected
{
    public class BDDfyCustomReporter : DefaultHtmlReportConfiguration
    {
        public static Property conf = Property.GetPropertyInstance();

        public override string OutputFileName
        {
            get
            {
                return conf.get("ReportFileName")+"_"+DateTimeOperations.GetTimeStamp(DateTime.Now);
            }
        }

        public override string ReportDescription
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Run On : " + DateTime.Now + " || ");
                sb.AppendLine("Environment : " + conf.get("Environment"));
                return sb.ToString();
                //return sb.ToString().Replace(Environment.NewLine, "<br />");
            }
        }

        public override string ReportHeader
        {
            get
            {
                return conf.get("ReportHeader");
            }
        }

    }
}
