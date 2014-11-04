using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using Xunit;
using PaymentsAutomation.UseCases.Einvoicing;

namespace PaymentsAutomation.Stories.Einvoicing
{
    [Story(AsA = "As a Merchant",
        IWant = "I want to create and send an Invoice to customer via WebMail",
        SoThat = "So that they can recieve the invoice copy in their mailbox",
        Title = "User story 2 - Send Invoice via WebMail to Customer")]
    public class SendInvoiceViaWebMail
    {
        [Fact]
        public void RunSendInvoiceTest()
        {
            new SendInvoiceTest().BDDfy<SendInvoiceViaWebMail>();
        }
    }
}
