using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameworkLibraries.Utils
{
    public class StringFunctions
    {
        public static string RemoveExtentionFromFileName(string fileName)
        {
            var splitFileName = fileName.Split('.');
            List<string> listFileNames = new List<string>(splitFileName);

            foreach (string item in listFileNames)
            {
                if (item.Equals("qbw") || item.Equals("QBW"))
                {
                    listFileNames.Remove(item);
                    break;
                }
            }
            return listFileNames[0];
        }

        public static string RandomString(int size)
        {
            StringBuilder builder = new StringBuilder();
            Random random = new Random();
            char ch;
            for (int i = 0; i < size; i++)
            {
                ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));
                builder.Append(ch);
            }

            return builder.ToString();
        }
 
        public static List<string> SplitString(string licenseNumber)
        {
            var splitString = licenseNumber.Split('-');
            return splitString.ToList();
        }
    }
}
