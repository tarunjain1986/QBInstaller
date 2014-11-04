using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameworkLibraries.Utils
{
    class LastException
    {
        private static String sExceptionMessage = null;

        public static void SetLastError(String sError)
        {
            sExceptionMessage = sError;
        }

        public static String GetLastError()
        {
            return sExceptionMessage;
        }
    }
}
