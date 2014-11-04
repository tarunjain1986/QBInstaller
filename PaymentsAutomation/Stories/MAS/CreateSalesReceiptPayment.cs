using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using Xunit;
using PaymentsAutomation.UseCases.MAS;

namespace PaymentsAutomation.Stories.MAS
{
    [Story(AsA = "As a Merchant",
        IWant = "I want to create a Sale Recipt",
        SoThat = "So that I can process the payment for the created sales receipt",
        Title = "User story 1 - Create Sales Receipt and process payment")]
    public class CreateSalesReceiptPayment
    {
        [Fact]
        public static void RunCreateSalesReceiptTest()
        {
            new SalesScript().BDDfy<CreateSalesReceiptPayment>();
        }
    }
}