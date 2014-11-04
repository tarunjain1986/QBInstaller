using System;
using Xunit;

namespace XunitDemo
{
    public class XunitGroupTests
    {
        [Fact]
        [Trait("Category", "P1")]
        public void TestMethod1()
        {
            Assert.True(true);
        }

        [Fact]
        [Trait("Category", "P2")]
        public void TestMethod2()
        {
            Assert.True(false);
        }

        [Fact]
        [Trait("Category", "P1")]
        public void TestMethod3()
        {
            Assert.NotNull(true);
        }
    }
}
