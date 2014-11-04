using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using Xunit;
using PaymentsAutomation.UseCases.MAS;

namespace PaymentsAutomation.Stories.MAS
{
    [Story(AsA = "As a Merchant",
        IWant = "I want to process a Credit Card - VISA transaction",
        SoThat = "So that I can receive payment from customer",
        Title = "User story 1 - Process credit card payment")]
    public class ProcessCreditCardTransaction
    {
        [Fact]
        public static void RunPCCPPaymentReadynessTest()
        {

            new PCCPPaymentReadynessTest().BDDfy<ProcessCreditCardTransaction>();
        }


        [Fact]
        public static void RunProcessPaymentTest1()
        {
            new PCCPPaymentTestVisa().BDDfy<ProcessCreditCardTransaction>();
        }

        [Fact]
        public static void RunProcessPaymentTest2()
        {
            new PCCPPaymentTestMaster().BDDfy<ProcessCreditCardTransaction>();
        }

        [Fact]
        public static void RunProcessPaymentTest3()
        {
            new PCCPPaymentTestAmex().BDDfy<ProcessCreditCardTransaction>();
        }


    }
}
