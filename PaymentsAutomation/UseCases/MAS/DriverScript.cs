using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.BDDfy;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems;
using Xunit;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using PaymentsAutomation.Stories.MAS;
using Microsoft.Office.Interop.Excel;


namespace PaymentsAutomation.UseCases.MAS
{
    //[TestClass]
    public class DriverScript
    {
        [Fact]
        public static void DriveTests()
        {
            //ExcelLibrary.setUpEnvironment();
            // Todo read module name ( MAS,PAAF,UA,RECON, Einvoicing )from config file
            ExcelLibrary.readExcel();

            //Todo read the userstory/tescase name from excel and excute the same method
            for (int m = 2; m < ExcelLibrary.valueArray.GetUpperBound(0); m++)
            {
                if (ExcelLibrary.valueArray[m, 10].Equals("EXECUTE"))
                {
                    for (int n = 1; n < ExcelLibrary.lastUsedCloumn + 1; n++)
                    {
                        ExcelLibrary.CCNumber = ExcelLibrary.valueArray[m, 3].ToString();
                        ExcelLibrary.CCExpDate = ExcelLibrary.valueArray[m, 4].ToString();
                        ExcelLibrary.CCExpYear = ExcelLibrary.valueArray[m, 5].ToString();
                        ExcelLibrary.CCZipCode = ExcelLibrary.valueArray[m, 6].ToString();
                        ExcelLibrary.CCPaymentAmount = ExcelLibrary.valueArray[m, 7].ToString();
                        ExcelLibrary.CCCustName = ExcelLibrary.valueArray[m, 8].ToString();
                        ExcelLibrary.CCType = ExcelLibrary.valueArray[m, 9].ToString();
                        if (ExcelLibrary.CCType.Equals("PCCP"))
                        {
                            ProcessCreditCardTransaction.RunPCCPPaymentReadynessTest();
                            ProcessCreditCardTransaction.RunProcessPaymentTest1();
                        }
                        else if ((ExcelLibrary.CCType.Equals("Refund")))
                        {
                            RefundTest.RunRefundTest();
                        }
                        else if ((ExcelLibrary.CCType.Equals("SalesReceipt")))
                        {
                            CreateSalesReceiptPayment.RunCreateSalesReceiptTest();
                        }
                        else if ((ExcelLibrary.CCType.Equals("AuthandCapture")))
                        {
                            AuthandCapturization.RunAuthandCaptureTest();
                        }
                    }
                }
            }
        }

    }
}
