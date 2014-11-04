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
    public class MayaLoginTest
    {
        Logger log = new Logger("Maya Test");
        Window mayaMainWindow = null;
        Window mayaWindow_2 = null;
        static Property conf = Property.GetPropertyInstance();
        static int Execution_Speed = int.Parse(conf.get("ExecutionSpeed"));
        static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        string User_Name = null;
        string Password = null;


        [Given(StepTitle = "Maya Client should be up and running")]
        public void Setup()
        {
            mayaMainWindow = Actions.GetDesktopWindow("Maya Client");
            Actions.WaitForWindow("Maya Client", Sync_Timeout);
        }

        [When(StepTitle = "Login is successful using a valid user name and password and able to select a company file")]
        public void Login()
        {
            Actions.SelectMenu(null, mayaMainWindow, "Favorites", "Customers");
            Actions.WaitForWindow("Maya Window - 2", Sync_Timeout);
            mayaWindow_2 = Actions.GetDesktopWindow("Maya Window - 2");
            Actions.WaitForElementEnabled(mayaWindow_2, "Sign In", Sync_Timeout);
            Maya.LoginIntoMaya(mayaWindow_2, false);
        }

        [AndWhen(StepTitle = "There should be no open dialog elements and Maya pages are accessible")]
        public void BaseState()
        {
            Maya.PrepareMayaBaseState(mayaWindow_2);
        }

        [Then(StepTitle = "Home link on the QBO home page is enabled and selection should be successful")]
        public void Test()
        {
            var panels = Actions.GetAllPanesInWindow(mayaWindow_2);
            Actions.ClickTextInsidePanel(mayaWindow_2, panels[2], "Home");
        }

        [AndThen(StepTitle = "Logging out of Maya should be successful")]
        public void Teardown()
        {
            Maya.SingOutMaya(mayaWindow_2, "Manish Sinha");
        }

    }

    

    public class MayaInvalidPasswordTest
    {


        [Given(StepTitle = "Maya Client should be up and running")]
        public void Setup()
        {
            throw new NotImplementedException("");
        }

        [When(StepTitle = "Login is unsuccessful using a valid user name and invalid password ")]
        public void Login()
        {
            throw new NotImplementedException("");
        }

        [AndWhen(StepTitle = "There should be no open dialog elements and Maya pages are accessible")]
        public void BaseState()
        {
            throw new NotImplementedException("");
        }

        [Then(StepTitle = "The dialog box should throw an error that password is invalid")]
        public void Test()
        {
            throw new NotImplementedException("");
        }

        public void Teardown()
        {
            throw new NotImplementedException("");
            
        }

    }
    public class MayaInvalidUserNameTest
    {


        [Given(StepTitle = "Maya Client should be up and running")]
        public void Setup()
        {
            throw new NotImplementedException("");
        }

        [When(StepTitle = "Login is unsuccessful using a invalid user name ")]
        public void Login()
        {
            throw new NotImplementedException("");
        }

        [AndWhen(StepTitle = "There should be no open dialog elements and Maya pages are accessible")]
        public void BaseState()
        {
            throw new NotImplementedException("");
        }

        [Then(StepTitle = "The dialog box should throw an error that password is invalid")]
        public void Test()
        {
            throw new NotImplementedException("");
        }

        public void Teardown()
        {
            throw new NotImplementedException("");
            
        }

    }
    public class MayaStaySignedInTest
    {


        [Given(StepTitle = "Maya Client should be up and running")]
        public void Setup()
        {
            throw new NotImplementedException("");
        }

        [When(StepTitle = "User has selected an option to stay signed in  ")]
        public void Login()
        {
            throw new NotImplementedException("");
        }

        [AndWhen(StepTitle = "User enters correct username and password and logins into Maya")]
        public void BaseState1()
        {
            throw new NotImplementedException("");
        }

        [AndWhen(StepTitle = "User closes the Maya Client and launches it again")]

        public void BaseState2()
        {
            throw new NotImplementedException("");
        }
        
        [Then(StepTitle = "The previously closed Company file should open again and no prompt for login")]
        public void Test1()
        {
            throw new NotImplementedException("");
        }

        [AndThen(StepTitle = "The previously open child windows should also open again")]
        public void Test2()
        {
            throw new NotImplementedException("");
        }

        public void Teardown()
        {
            throw new NotImplementedException("");
            
        }

    }

    public class MayaSignUpTest
    {


        [Given(StepTitle = "Maya Client should be up and running")]
        public void Setup()
        {
            throw new NotImplementedException("");
        }

        [When(StepTitle = "User has selected to create a new account with valid credentials  ")]
        public void Login()
        {
            throw new NotImplementedException("");
        }

        [AndWhen(StepTitle = "User enters a SKU option Simple/Essentials and signs up")]
        public void BaseState()
        {
            throw new NotImplementedException("");
        }

        

        [Then(StepTitle = "A new company creation window should come for the user")]
        public void Test1()
        {
            throw new NotImplementedException("");
        }

    

        public void Teardown()
        {
            throw new NotImplementedException("");

        }

    }
}
