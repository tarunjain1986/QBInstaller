using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FrameworkLibraries.Utils
{
    public class PathBuilder
    {
        public static string GetPath(string fileName)
        {
            var pathBuilder = new StringBuilder();
            UriBuilder uri = new UriBuilder(Assembly.GetExecutingAssembly().CodeBase);
            string location = Path.GetDirectoryName(Uri.UnescapeDataString(uri.Path));
            pathBuilder.Append(location);
            pathBuilder.Append("\\"+fileName);
            return pathBuilder.ToString();
        }
    }
}
