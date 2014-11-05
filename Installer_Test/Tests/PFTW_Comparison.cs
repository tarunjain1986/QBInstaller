using System;
using System.IO;
// using System.Linq;
// using System.Threading;
// using System.Reflection;
// using System.Diagnostics;
// using System.Windows.Forms;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.ActionLibs;
using FrameworkLibraries.AppLibs.QBDT;

using TestStack.BDDfy;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;

using Xunit;

using Installer_Test;


namespace Installer_Test.Tests
{

    public class PFTW_Comparison
    {

        public static Property conf = Property.GetPropertyInstance();
        public static int Sync_Timeout = int.Parse(conf.get("SyncTimeOut"));
        public static string testName = "PFTW_Comparison";
        string Build01_Path, Build02_Path, Windiff_Path, Local_B1Path, Local_B2Path, Local_Windiff;

        [Given(StepTitle = @"The Builds and windiff tool are copied on the local machine")]

        public void Setup()
        {
            var timeStamp = DateTimeOperations.GetTimeStamp(DateTime.Now);
            Logger log = new Logger(testName + "_" + timeStamp);

            //////////////////////////////////////////////////////////////////////////////////////////////
            // The following code is for reading from an excel file
            //////////////////////////////////////////////////////////////////////////////////////////////

            string readpath = "C:\\Temp\\Parameters.xlsx"; // "C:\\Installation\\Sample.txt";

            Dictionary<string, string> dic = new Dictionary<string, string>();
           
            dic = Installer_Test.Lib.File_Functions.ReadExcelValues(readpath, "PFTW", "B2:B4");

            Build01_Path = dic["B2"];
            Build02_Path = dic["B3"];
            Windiff_Path = dic["B4"];

            Local_B1Path = @"C:\Temp\PFTW\ReleaseCandidate\";
            Local_B2Path = @"C:\Temp\PFTW\Web\";
            Local_Windiff = @"C:\Temp\PFTW\";

           

            /////////////////////////////////////////////////////////////////////////////////////////////////////
            //if (!Directory.Exists(Local_B1Path))
            //{
            //    Actions.DirectoryCopy(Build01_Path, Local_B1Path, true);
            //}

            //if (!Directory.Exists(Local_B2Path))
            //{
            //    Actions.DirectoryCopy(Build02_Path, Local_B2Path, true);
            //}

            //if (!File.Exists (@"C:\Temp\PFTW\windiff.exe"))
            //{
            //    File.Copy (Windiff_Path + "windiff.exe", Local_Windiff+"windiff.exe" );
            //}
            //if (!File.Exists(@"C:\Temp\PFTW\gutils.dll"))
            //{
            //    File.Copy(Windiff_Path + "gutils.dll", Local_Windiff + "gutils.dll");
            //}
            /////////////////////////////////////////////////////////////////////////////////////////////////////
        }


        [Then(StepTitle = "Then - Run the diff")]
        public void Run_Windiff_Compare()
        {
            PFTW.Windiff_Compare(Local_Windiff, Local_B1Path, Local_B2Path);
        }

        [AndThen (StepTitle = "And Then - Compare the file size" )]
        public void Run_FileSize_Compare()
        {
            string[] B1_files = Directory.GetFiles(Local_B1Path, "*.exe");
            string[] B2_files = Directory.GetFiles(Local_B2Path, "*.exe");
            string file_name;

           foreach (string file in B1_files)
           {
               file_name = Path.GetFileName(file);
               PFTW.Compare_FileSize(Local_B1Path, Local_B2Path, file_name);
           }
        }

        //[AndThen(StepTitle = "And Then - Compare the Digital Signatures")]
        //public void Run_DigSign_Compare()
        //{
           
        //}

        [Fact]
        public void Run_PFTWComparison()
        {
            this.BDDfy();
        }
    }
}
