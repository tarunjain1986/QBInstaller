using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BATS.DATA
{
    public static class OpenPortableFileTestDataSource
    {
        private static List<object[]> _data = new List<object[]>
        {
            new object[]{"Matrix"}

        };

        public static IEnumerable<object[]> TestData
        {
            get { return _data; }
        }
    }
}
