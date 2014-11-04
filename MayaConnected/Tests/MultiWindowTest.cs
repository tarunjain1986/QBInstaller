using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries.Utils;
using System.Windows.Automation;
using FrameworkLibraries.AppLibs.MayaConnected;
using TestStack.BDDfy;
using TestStack.White.UIItems.WindowItems;
using Xunit;
using TestStack.BDDfy.Configuration;

namespace MayaConnected.Tests
{
    public class ResizeMainWindowTest
    {

        [Given(StepTitle = "Maya Client should be up and user should be logged in")]
        public void Setup()
        {

        }

        [When(StepTitle = "The Main window is resized using Mouse/Keyboard shortcuts")]
        public void Resize()
        {

        }


        [Then(StepTitle = "Then the Main Window should minimise/maximise accordingly")]
        public void Test()
        {
            throw new NotImplementedException("");

        }

    }

    public class ToggleWindowTest
    {

        [Given(StepTitle = "Maya Client should be up and user should be logged in")]
        public void Setup1()
        {

        }

        [AndGiven(StepTitle = "Multiple Windows i.e both parent and child windows should be open")]
        public void Setup2()
        {

        }

        [When(StepTitle = "The user toggles between different windows using mouse/ keyboard shortcut ")]
        public void Toggle()
        {

        }


        [Then(StepTitle = "Then the toggled to window should be in focus")]
        public void Test()
        {
            throw new NotImplementedException("");

        }

    }

    public class CloseMultipleWindowsTest
    {

        [Given(StepTitle = "Maya Client should be up and user should be logged in")]
        public void Setup1()
        {

        }

        [AndGiven(StepTitle = "Multiple Windows i.e both parent and child windows should be open")]
        public void Setup2()
        {

        }

        [When(StepTitle = "The Main Window is closed by the user using mouse/Menu options ")]
        public void Close()
        {

        }


        [Then(StepTitle = "All the open windows should close simultaneosuly ")]
        public void Test()
        {
            throw new NotImplementedException("");

        }

    }


}
