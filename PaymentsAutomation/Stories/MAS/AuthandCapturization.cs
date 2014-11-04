using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using Xunit;
using PaymentsAutomation.UseCases.MAS;

namespace PaymentsAutomation.Stories.MAS
{
    [Story(AsA = "As a Merchant",
        IWant = "I want to process an auth",
        SoThat = "So that I can capture the authorized amount",
        Title = "User story 1 - Auth/Capture")]
    public class AuthandCapturization
    {
        [Fact]
        public static void RunAuthandCaptureTest()
        {
            new AuthandCapture().BDDfy<AuthandCapturization>();
        }
    }
    
}
