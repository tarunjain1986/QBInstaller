using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using Xunit;
using MayaConnected.Scenarios;

namespace MayaConnected.Stories
{
    [Story(AsA = "As a Customer",
        IWant = "I want to see a CEF browser inside the Client Window",
        SoThat = "So that it can render QBO pages ",
        Title = "User story QBWG-23544- The QBO web pages should be rendered inside CEF Browser ")]

    public class CEFBrowserStory
    {
        [Fact]
        [Category("Maya")]
        [Category("Maya - High")]
        public void RunMayaCEFBrowserMainWindowTest()
        {
            new CEFBrowserMainWindowTest().BDDfy<CEFBrowserStory>();
        }

        [Fact]
        [Category("Maya")]
        [Category("Maya - High")]
        public void RunMayaCEFBrowserChildWindowTest()
        {
            new CEFBrowserChildWindowTest().BDDfy<CEFBrowserStory>();
        }
    }
}

