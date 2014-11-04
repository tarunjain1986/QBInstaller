using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameworkLibraries.Utils
{
    public class DateTimeOperations
    {
        public static string GetTimeStamp(DateTime value)
        {
            return value.ToString("yyyyMMddHHmm");
        }
    }
}
