using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Reflection;
using System.Diagnostics;
using System.Collections.Generic;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;

using Xunit;

using Installer_Test;

namespace Installer_Test.Tests
{
   
    public class AntiVirus
    {
        //string AVPath = @"\\banfsalab02\Users\RajSunder\AntiVirus-Trial";
        public string AVName;
        public string testName = "Anti Virus Install";
        public string [] AntiVirusSW;
        public static Property conf = Property.GetPropertyInstance();
        public static string Sync_Timeout = conf.get("SyncTimeOut");

       
       [Given(StepTitle = @"The anti virus software(s) to be installed are mentioned in C:\Installation\Parameters.txt")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);
           
            string readpath = @"C:\Temp\Parameters.txt";
            File.WriteAllLines(readpath, File.ReadAllLines(readpath).Where(l => !string.IsNullOrWhiteSpace(l))); // Remove white space from the file

            string[] lines = File.ReadAllLines(readpath);
            var dic = lines.Select(line => line.Split('=')).ToDictionary(keyValue => keyValue[0], bits => bits[1]);

            AVName = dic["AntiVirusSW"];
        }
        

        [Then(StepTitle = @"Then - Copy the AntiVirus Software(s) to C:\Temp\AntiVirus\")]
        public void Copy_AntiVirus()
        {
             Install_Functions.Copy_AVSoftware(AVName); 
        }

        [AndThen(StepTitle = "And Then - Install the selected AntiVirus software.")]
        public void Install_AntiVirus()
        {
            Install_Functions.Install_AVSoftware(AVName);
        }

        [AndThen(StepTitle = "And Then - Scan the QuickBooks Installer with the installed antivirus software.")]
        public void Scan_AntiVirus()
        {
            Install_Functions.Scan_AVSoftware(AVName);
        }

        [Fact]
        public void Run_AntiVirusTest()
        {
          this.BDDfy();
        }
    }
}
