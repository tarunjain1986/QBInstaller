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
    public class CEFBrowserMainWindowTest
    {
        
        [Given(StepTitle = "Maya Client should be up and running")]
        public void Setup()
        {
            throw new NotImplementedException("");
        }

        [When(StepTitle = "Login is successful ")]
        public void Login()
        {
            throw new NotImplementedException("");
        }
  

        [Then(StepTitle = "The QBO Homepage should load within a CEF Browser")]
        public void Test()
        {
             throw new NotImplementedException("");
  
        }

  }

    public class CEFBrowserChildWindowTest
    {

        [Given(StepTitle = "Maya Client should be up and running and Login is successful")]
        public void Setup()
        {
            throw new NotImplementedException("");
        }

        [When(StepTitle = " On pressing a shortcut key for a new  child window")]
        public void Login()
        {
            throw new NotImplementedException("");
        }


        [Then(StepTitle = "The new child window should load within a CEF Browser")]
        public void Test()
        {
            throw new NotImplementedException("");

        }

    }



    
}
