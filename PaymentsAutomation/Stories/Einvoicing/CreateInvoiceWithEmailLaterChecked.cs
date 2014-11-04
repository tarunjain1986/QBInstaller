using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using Xunit;
using PaymentsAutomation.UseCases.Einvoicing;

namespace PaymentsAutomation.Stories.Einvoicing
{
    [Story(AsA = "As a Merchant",
        IWant = "I want to ensure that Email Later is checked when I Create an Invoice for Customer with their email id",
        SoThat = "So that I can access them later in Send Forms",
        Title = "User story 1 - Create Invoice With Email Later Checked and access them later in Send Forms")]
    public class CreateInvoiceWithEmailLaterChecked
    {
        [Fact]
        public void RunCreateInvoiceTest()
        {
            new CreateInvoiceTest().BDDfy<CreateInvoiceWithEmailLaterChecked>();
        }
    }
}
