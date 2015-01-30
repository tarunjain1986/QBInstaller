using System;
using TestStack.White.UIItems.WindowItems;
using FrameworkLibraries.ActionLibs.WhiteAPI;
using Microsoft.VisualStudio.TestTools.UnitTesting;




namespace Installer_Test.Properties.Lib
{
  
    public class SwitchEdition
    {
        string[,] arrEdition;
        string currEdition;

        public void SwitchEdition(TestStack.White.Application qbApp, string SKU)
        {
            TestStack.White.UIItems.WindowItems.Window qbWindow = null;
            qbWindow = FrameworkLibraries.AppLibs.QBDT.QuickBooks.PrepareBaseState(qbApp);
            Window editionWindow = Actions.GetChildWindow(qbWindow, "Select QuickBooks Industry-Specific Edition");

            if (SKU == "Enterprise")
            {
                arrEdition = new string [6,2]{ {"Enterprise Solutions Business", "false"} , {"Enterprise Solutions Contractor", "false"},
                {"Enterprise Solutions Manufacturing & Wholesale", "false"}, {"Enterprise Solutions Nonprofit","false"},
                {"Enterprise Solutions Professional Services", "false"}, {"Enterprise Solutions Retail","false"}};
            }

          
            for (int i = 0; i < arrEdition.Length; i++)
            {
                currEdition = arrEdition[i, 1] + " - Currently open  ";

                if (Actions.CheckElementIsEnabled(editionWindow, currEdition))
                {
                    Actions.ClickElementByName(editionWindow, currEdition);
                }

                else
                {
                    Actions.ClickElementByName(editionWindow, "Cancel");
                }

            }




        }
    }
}
