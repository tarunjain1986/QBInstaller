using FrameworkLibraries.ActionLibs.WhiteAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using TestStack.White.UIItems.WindowItems;
using Xunit;
using Xunit.Extensions;

namespace XunitDemo
{
    public class CrashTest : IDisposable
    {

        public CrashTest()
        {
            //Setup method...
        }

        public void Dispose()
        {
            //teardown method..
        }

        [Fact]
        public void CrashTestMethod()
        {
            Window test = null;
            Actions.SelectComboBoxItemByText(test, FrameworkLibraries.ObjMaps.QBDT.WhiteAPI.Common.Objects.StateName_CmbBox_AutoID, "DE");
        }



    }
}
