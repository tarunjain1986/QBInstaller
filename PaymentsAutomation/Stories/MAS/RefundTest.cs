using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using Xunit;
using PaymentsAutomation.UseCases.MAS;

namespace PaymentsAutomation.Stories.MAS
{
    
        [Story(AsA = "As a Merchant",
            IWant = "I want to process a refund",
            SoThat = "So that I can refund the amount for the transactions made",
            Title = "User story 1 - Process Refund")]

        public class RefundTest
        {
            [Fact]
            public static void RunRefundTest()
            {
                new Refund().BDDfy<RefundTest>();
            }
        }
        /*[Story(AsA = "As a Merchant",
            IWant = "I want to process a Credit Card - VISA transaction",
            SoThat = "So that I can receive payment from customer",
            Title = "User story 2 - Process credit card payment")]
        public class PCCP
        {
            [Fact]
            public void RunCreateSalesReceiptTest()
            {
                new PCCPPaymentTest().BDDfy<PCCP>();
            }
        }
        [Story(AsA = "As a Merchant",
            IWant = "I want to create a Sale Recipt",
            SoThat = "So that I can process the payment for the created sales receipt",
            Title = "User story 1 - Create Sales Receipt and process payment")]
        public class ProcessSalesReceipt
        {
            [Fact]
            public void RunCreateSalesReceiptTest()
            {
                new SalesScript().BDDfy<ProcessSalesReceipt>();
            }
        }
        [Story(AsA = "As a Merchant",
            IWant = "I want to process a refund",
            SoThat = "So that I can refund the amount for the transactions made",
            Title = "User story 1 - Process Refund")]
        public class AuthrizeAndCapture
        {
            [Fact]
            public void RunAuthandCaptureTest()
            {
                new AuthandCapture().BDDfy<AuthrizeAndCapture>();
            }
        }*/

    }
