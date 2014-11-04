using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using Xunit;
using MayaConnected.Scenarios;

namespace MayaConnected.Stories
{
    [Story(AsA = "As a Customer",
        IWant = "I want to log-in to Maya Client",
        SoThat = "So that i can access company file in QBO",
        Title = "User story-QBWG-23548 - Log-in to Maya Client")]
    
    public class MayaLoginStory
    {
        [Fact]
        [Category("Maya")]
        [Category("Maya - High")]
        public void RunMayaLoginTest()
        {
            new MayaLoginTest().BDDfy<MayaLoginStory>();
        }

        [Fact]
        [Category("Maya")]
        [Category("Maya - High")]
        public void RunMayaInvalidPasswordTest()
        {
            new MayaInvalidPasswordTest().BDDfy<MayaLoginStory>();
        }

        [Fact]
        [Category("Maya")]
        [Category("Maya - Medium")]
        public void RunMayaInvalidUserNameTest()
        {
            new MayaInvalidUserNameTest().BDDfy<MayaLoginStory>();
        }

        [Fact]
        [Category("Maya")]
        [Category("Maya - Medium")]
        public void RunMayaStaySignedInTest()
        {
            new MayaStaySignedInTest().BDDfy<MayaLoginStory>();
        }

        [Fact]
        [Category("Maya")]
        [Category("Maya - Low")]
        public void RunMayaSignUpTest()
        {
            new MayaInvalidUserNameTest().BDDfy<MayaLoginStory>();
        }

    }
}
