using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using Xunit.Extensions;

namespace XunitDemo
{
    public class XunitTest : IDisposable
    {

        public XunitTest()
        {
            //Setup method...
        }

        public void Dispose()
        {
            //teardown method..
        }

        [Fact]
        public void ValueTyeEqual()
        {
            IEnumerable<int> numbers1 = Enumerable.Range(1, 10);
            IEnumerable<int> numbers2 = Enumerable.Range(1, 10);
            Assert.Equal(numbers1, numbers2);
        }

        [Fact]
        public void ValueTyeNotEqual()
        {
            IEnumerable<int> numbers1 = Enumerable.Range(1, 10);
            IEnumerable<int> numbers2 = Enumerable.Range(1, 11);
            Assert.Equal(numbers1, numbers2);
        }

        [Fact]
        public void DecimalWithPrecision()
        {
            var d1 = new Decimal(24.2111);
            var d2 = new Decimal(24.2112);

            //Assert.Equal(d1, d2);

            Assert.Equal(d1, d2, 2);
        }

        [Fact]
        public void AssertThrowsTest()
        {
            Assert.Throws<DivideByZeroException>(() =>
                {
                    var a = 10;
                    var b = 0;
                    var c = (a / b);
                });
        }

        [Fact(Skip="No Run..")]
        public void SkipTest()
        {
            Assert.Throws<DivideByZeroException>(() =>
            {
                var a = 10;
                var b = 0;
                var c = (a / b);
            });
        }

        [Theory]
        [InlineData(9)]
        [InlineData(10)]
        public void DataDrivenTest(int num)
        {
            var a = num;
        }

        [Theory]
        [PropertyData("TestData", PropertyType = typeof(PropertyTestDataSource))]
        public void PropertyDrivenTest(int num)
        {
            var a = num;
        }

        [Fact]
        [Category("Custom Category - P1")]
        public void CustomCategoryTest()
        {
            Assert.Equal(true, false);
        }

        [Fact]
        [SmokeTest]
        public void CustomCategoryDefinitionTest()
        {
            Assert.Equal(false, false);
        }


    }
}
