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

namespace MayaConnected.Scenarios
{
    public class ResizeMainWindowTest
    {

        [Given(StepTitle = "Maya Client should be up and user should be logged in")]
        public void Setup()
        {
            throw new NotImplementedException("");
        }

        [When(StepTitle = "The Main window is resized using Mouse/Keyboard shortcuts")]
        public void Resize()
        {
            throw new NotImplementedException("");
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
            throw new NotImplementedException("");
        }

        [AndGiven(StepTitle = "Multiple Windows i.e both parent and child windows should be open")]
        public void Setup2()
        {
            throw new NotImplementedException("");
        }

        [When(StepTitle = "The user toggles between different windows using mouse/ keyboard shortcut ")]
        public void Toggle()
        {
            throw new NotImplementedException("");
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
            throw new NotImplementedException("");
        }

        [AndGiven(StepTitle = "Multiple Windows i.e both parent and child windows should be open")]
        public void Setup2()
        {
            throw new NotImplementedException("");
        }

        [When(StepTitle = "The Main Window is closed by the user using mouse/Menu options ")]
        public void Close()
        {
            throw new NotImplementedException("");
        }


        [Then(StepTitle = "All the open windows should close simultaneosuly ")]
        public void Test()
        {
            throw new NotImplementedException("");

        }

    }


}
