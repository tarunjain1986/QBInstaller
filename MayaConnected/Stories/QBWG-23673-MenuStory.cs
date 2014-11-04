using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using Xunit;
using MayaConnected.Scenarios;

namespace MayaConnected.Stories
{
    [Story(AsA = "As a Customer",
        IWant = "I want to see Menu options on the native client",
        SoThat = "So that i have native functionalities for Money In and Money Out Flow",
        Title = "User story-QBWG-23673 Menu on Maya Client")]

    public class MenuStory
    {
        [Fact]
        [Category("Maya")]
        [Category("Maya - High")]
        public void RunMayaFileMenuTest()
        {
            new FileMenuTest().BDDfy<MenuStory>();
        }

        [Fact]
        [Category("Maya")]
        [Category("Maya - Medium")]
        public void RunMayaEditMenuTest()
        {
            new EditMenuTest().BDDfy<MenuStory>();
        }

        [Fact]
        [Category("Maya")]
        [Category("Maya - Low")]
        public void RunMayaReportMenuTest()
        {
            new ReportMenuTest().BDDfy<MenuStory>();
        }

        
    }
}
