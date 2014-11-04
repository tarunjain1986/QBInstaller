using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PaymentsAutomation.UseCases.MAS;
using System.Collections.Generic;

namespace PaymentsAutomation.Stories.MAS
{
    [TestClass]
    public class UnitTest1
    {
        public static string CCNumber, testType;
        public int Row;
        public static Dictionary<string, string> dic = new Dictionary<string, string>();
        [TestMethod]
        public static void getData(Dictionary<string, string> dic, int Row)
        {
            CCNumber = dic["B" + Row];
            testType = dic["F" + Row];
            if (testType.Equals("CC"))
            {
                ProcessCreditCardTransaction.RunProcessPaymentTest1();
            }
            if (testType.Equals("anyType"))
            {
                //corresspondingFunction();
            }
        }
    }
}
