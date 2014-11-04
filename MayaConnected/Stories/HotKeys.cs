using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using Xunit;

namespace MayaConnected.Stories
{
        [Story(AsA = "As a Customer",
       IWant = "I want to hot keys for most of the common actions to work in Maya Client",
       SoThat = "So that accountants can use key short cuts effectively",
       Title = "Hot keys implementation in Maya Client")]

        public class HotKeys
        {
            [Fact]
            [Category("Maya")]
            [Category("Maya - High")]
            public void RunMayaLoginTest()
            {
                throw new NotImplementedException("Not implemented yet..");
            }
        }
}
