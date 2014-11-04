using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XunitDemo
{
    public static class PropertyTestDataSource
    {
        private static List<object[]> _data = new List<object[]>
        {
            new object[]{9},
            new object[]{10},
            new object[]{102}
        };

        public static IEnumerable<object[]> TestData
        {
            get { return _data; }
        }
    }
}
