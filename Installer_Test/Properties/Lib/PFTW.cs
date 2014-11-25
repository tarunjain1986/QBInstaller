using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Drawing;
using System.IO.Packaging;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs;

//using Microsoft.VisualStudio.TestTools.UnitTesting;
using ScreenShotDemo;
using TestStack.BDDfy;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems;

namespace Installer_Test
{

    public class PFTW
    {
        public static void Windiff_Compare(string Local_Windiff, string Local_B1Path, string Local_B2Path)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("Windiff comparison started:" + Local_B1Path + Local_B2Path + " - Started..");

            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
            string cmdText = "/c cd " + Local_Windiff + " && windiff " + Local_B1Path + " " + Local_B2Path + " -Sslrdx " + Local_Windiff + "Diff.txt";
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = cmdText;
            startInfo.UseShellExecute = true;
            process.StartInfo = startInfo;
            try
            {
                process.Start();
                process.WaitForExit();
                Logger.logMessage("Windiff comparison: " + Local_B1Path + Local_B2Path + " - Successful");
                Logger.logMessage("------------------------------------------------------------------------------");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                Logger.logMessage("Windiff comparison: " + Local_B1Path + Local_B2Path + " - Failed");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }

        public static void Compare_FileSize(string filePath1, string filePath2, string file)
        {
            Logger.logMessage("Function call @ :" + DateTime.Now);
            Logger.logMessage("PFTW File Comparison " + filePath1 + file + " " + filePath2 + file + " - Started..");
            file = "Test.txt";

            FileInfo FileInfo_01 = new FileInfo(filePath1 + file);
            FileInfo FileInfo_02 = new FileInfo(filePath2 + file);
            try
            {

                DateTime DateTime1;
                //DateTime1 = System.IO.File.GetLastWriteTime(filePath2 + file);

                // Open the package.

                //FileStream fr = new FileStream(filePath2 + file, FileMode.Open, FileAccess.Read, FileShare.Read);
                
                /////////////////////////////////////////////////////////////
                byte[] buffer = File.ReadAllBytes(filePath2 + file);
                //string base64Encoded = Convert.ToBase64String(buffer);
                MemoryStream stream = new MemoryStream(buffer);

                //// TODO: do something with the bas64 encoded string
                //buffer = Convert.FromBase64String(base64Encoded);
                //File.WriteAllBytes(filePath2 + file, buffer);
                /////////////////////////////////////////////////////////////
                //Stream sr = File.OpenRead(filePath2 + file);

                // Stream fs1 = File.Open(filePath2 + file, FileMode.OpenOrCreate);
               // stream.Position = 0;
               //using (var package1 = Package.Open(stream,FileMode.Open));//,fs1.Read);// Package.Open(fs1,FileMode.Open);
               // //sr.Close();

                using (var package = Package.Open(stream))
                {
                    // do something with package
                }


                using (var stream1 = new FileStream(filePath2 + file, FileMode.Open, FileAccess.Read))
                {
                    using (var package = Package.Open(stream1))
                    {
                        // do something with package
                    }
                }
            }

            catch (Exception e)
            {
                Logger.logMessage(e.Message);
            }

            // Create the PackageDigitalSignatureManager
           // PackageDigitalSignatureManager dsm = new PackageDigitalSignatureManager(package1);
            //foreach (PackageDigitalSignature signature in dsm.Signatures)
            //{
            //    DateTime1 = GetDigTimeStamp(signature);
            //}


            try
            {
                if (FileInfo_01.Length == FileInfo_02.Length)
                {
                    Logger.logMessage("Files " + filePath1 + file + " and " + filePath2 + file + " are of same size: " + FileInfo_01.Length +" - Passed");
                    Logger.logMessage("------------------------------------------------------------------------------------------------------------------");
                }
                else
                {
                    Logger.logMessage("Files " + filePath1 + file + "(" + FileInfo_01.Length + ")" + " and " + filePath2 + file + "(" + FileInfo_02.Length + ")" + " are of not of same size - Failed");
                    Logger.logMessage("---------------------------------------------------------------------------------------------------------------------------------------------------------------");
                }
            }

            catch (Exception e)
            {
                Logger.logMessage("Error in file comparison!");
                Logger.logMessage(e.Message);
                Logger.logMessage("------------------------------------------------------------------------------");
            }
        }

        public static DateTime GetDigTimeStamp(PackageDigitalSignature p)
        {
            return p.SigningTime;
        }

    }
}
