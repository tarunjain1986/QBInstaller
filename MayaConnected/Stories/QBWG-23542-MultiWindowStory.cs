using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using Xunit;
using MayaConnected.Scenarios;

namespace MayaConnected.Stories
{
    [Story(AsA = "As a Customer",
        IWant = "I want to see multiple windows on the native client",
        SoThat = "So that these support the native windows functionality",
        Title = "QBWG-23542 - MVVM on Maya Client")]

    public class MultiWindowStory
    {
        [Fact]
        [Category("Maya")]
        [Category("Maya - High")]
        public void RunMayaResizeMainWindowTest()
        {
            new ResizeMainWindowTest().BDDfy<MultiWindowStory>();
        }

        [Fact]
        [Category("Maya")]
        [Category("Maya - Medium")]
        public void RunMayaToggleWindowTest()
        {
            new ToggleWindowTest().BDDfy<MultiWindowStory>();
        }

        [Fact]
        [Category("Maya")]
        [Category("Maya - Low")]
        public void RunMayaCloseMultipleWindowsTest()
        {
            new CloseMultipleWindowsTest().BDDfy<MultiWindowStory>();
        }
    }
}
