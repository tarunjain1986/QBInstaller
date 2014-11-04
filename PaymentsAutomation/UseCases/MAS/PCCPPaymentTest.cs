using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using FrameworkLibraries.AppLibs.QBDT;
using System.Threading;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.Payments;
using TestStack.BDDfy;
using PaymentsAutomation.Stories.MAS;
using PaymentsAutomation.UseCases.MAS;

namespace PaymentsAutomation.UseCases.MAS
{
    public class PCCPPaymentReadynessTest
    {        
        [Given(StepTitle = "QuickBooks Desktop should be up and running")]
        public void Setup()
        {
            Constants.Initialize(("ReceivePaymentPCCP.txt"));
        }
        [When(StepTitle = "Receive Payment Transaction should be invoked")]
        public void ReceivePayments()
        {
            //Payments.ProcessCCPayment(qbApp, qbWindow, "10", "4111111111111111", "09", "2025", "Cust1", "", "", "12345");
            //throw new NotImplementedException("need to be implemented");
        }
        [Then(StepTitle = "Process transaction")]
        public void voidPayment()
        {
            //Payments.voidCCPayment(qbApp, qbWindow, "Yes");
            //throw new NotImplementedException("need to be implemented");
        }
    }

    public class PCCPPaymentTestVisa
    {
        [Given(StepTitle = "QuickBooks Desktop is running")]
        public void Setup()
        {
            Constants.Initialize(("ReceivePaymentPCCP.txt"));
        }
        [When(StepTitle = "Receive Payment Transaction is getting invoked and transaction processed for VISA CARD")]
        public void ReceivePayments()
        {
            Payments.ProcessCCPayment(Constants.qbApp, Constants.qbWindow, ExcelLibrary.CCPaymentAmount, ExcelLibrary.CCNumber, ExcelLibrary.CCExpDate, ExcelLibrary.CCExpYear, ExcelLibrary.CCCustName, "", "", ExcelLibrary.CCZipCode);
        }
        [Then(StepTitle = "Void the peocessed transaction")]
        public void voidPayment()
        {
            Payments.voidCCPayment(Constants.qbApp, Constants.qbWindow, "Yes");
        }
    }

    public class PCCPPaymentTestMaster
    {
        [Given(StepTitle = "QuickBooks Desktop is running")]
        public void Setup()
        {
            Constants.Initialize(("ReceivePaymentPCCP.txt"));
        }
        [When(StepTitle = "Receive Payment Transaction is getting invoked and transaction processed for MASTER CARD")]
        public void ReceivePayments()
        {
            throw new NotImplementedException("need to be implemented");
            //Payments.ProcessCCPayment(Constants.qbApp, Constants.qbWindow, "10", "4111111111111111", "09", "2025", "Cust1", "", "", "12345");
        }
        [Then(StepTitle = "Accept the transactions")]
        public void voidPayment()
        {
            //Payments.voidCCPayment(Constants.qbApp, Constants.qbWindow, "No");
        }
    }

    public class PCCPPaymentTestAmex
    {
        [Given(StepTitle = "QuickBooks Desktop is running")]
        public void Setup()
        {
            Constants.Initialize(("ReceivePaymentPCCP.txt"));
        }
        [When(StepTitle = "Receive Payment Transaction is getting invoked and transaction processed for AMEX CARD")]
        public void ReceivePayments()
        {
            throw new NotImplementedException("need to be implemented");
            //Payments.ProcessCCPayment(Constants.qbApp, Constants.qbWindow, "10", "4111111111111111", "09", "2025", "Cust1", "", "", "12345");
        }
        [Then(StepTitle = "Accept the transactions")]
        public void voidPayment()
        {
            throw new NotImplementedException("need to be implemented");
            //Payments.voidCCPayment(Constants.qbApp, Constants.qbWindow, "No");
        }
    }
}

