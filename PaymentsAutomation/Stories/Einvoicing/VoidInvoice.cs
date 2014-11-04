using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Xunit;
using PaymentsAutomation.UseCases.Einvoicing;
using TestStack.BDDfy;

namespace PaymentsAutomation.Stories.Einvoicing
{
    [Story(AsA = "As a Merchant",
         IWant = "I want to create and send an Invoice to customer via WebMail and then void the same invoice",
         SoThat = "So that I can ensure ICN services are getting called",
         Title = "User story 2 - Void Invoice and check if ICN calls are going")]
    public class VoidInvoice
    {
        [Fact]
        public void RunVoidInvoiceTest()
        {
            new VoidInvoiceTest().BDDfy<VoidInvoice>();
        }
    }
}
