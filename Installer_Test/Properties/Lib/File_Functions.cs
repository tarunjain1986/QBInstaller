using System;
using System.IO;
using System.Linq;
using System.Management;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using Microsoft.Win32;


using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;
using FrameworkLibraries.Utils;

namespace Installer_Test.Lib
{
   
    public class File_Functions
    {
        public string expected_ver, installed_product, installed_dataPath, installed_commonPath;
        public static string reg_ver;

        public static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            // If the destination directory doesn't exist, create it. 
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, false);
            }

            // If copying subdirectories, copy them and their contents to new location. 
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }

        public static Dictionary<string, string> ReadTextValues(string readpath)
        {
   
            File.WriteAllLines(readpath, File.ReadAllLines(readpath).Where(l => !string.IsNullOrWhiteSpace(l))); // Remove white space from the file

            string[] lines = File.ReadAllLines(readpath);
            var dic = lines.Select(line => line.Split('=')).ToDictionary(keyValue => keyValue[0], bits => bits[1]);

            //ver = dic["Version"];
            //reg_ver = dic["Registry Folder"];
      
            return dic;
        }

        public static Dictionary<string, string> ReadExcelValues (string readpath,string workSheet, string Range)
        {
           
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(readpath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(workSheet);
            Excel.Range xlRng = (Excel.Range)xlWorkSheet.get_Range(Range, Type.Missing);

            Dictionary<string, string> dic = new Dictionary<string, string>();

            foreach (Excel.Range cell in xlRng)
            {

                string cellIndex = cell.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                string cellValue = Convert.ToString(cell.Value2);
                dic.Add(cellIndex, cellValue);

            }

            // xlWorkBook.Close();
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook.Close(false, misValue, misValue);

            xlApp.Quit();
            return dic;



            ////if (xlRng != null) Marshal.FinalReleaseComObject(xlRng);
            ////if (xlWorkSheet != null) Marshal.FinalReleaseComObject(xlWorkSheet);
            ////if (xlWorkBooks != null) Marshal.FinalReleaseComObject(xlWorkBooks);
            ////if (xlWorkBook != null)
            ////{
            ////    xlWorkBook.Close(Type.Missing, Type.Missing, Type.Missing);
            ////    Marshal.FinalReleaseComObject(xlWorkBook);
            ////}
            ////if (xlApp != null)
            ////{
            ////    xlApp.Quit();
            ////    Marshal.FinalReleaseComObject(xlApp);
            ////}

            ////xlWorkBook.Close();
            ////xlApp.Quit();

            //   GC.Collect();
            //   GC.WaitForPendingFinalizers();
            ////   GC.Collect();


            //   Marshal.FinalReleaseComObject(xlRng);
            //   Marshal.FinalReleaseComObject(xlWorkSheet);
            //   Marshal.FinalReleaseComObject(xlWorkBook);
            //   Marshal.FinalReleaseComObject(xlWorkBooks);
            //   Marshal.FinalReleaseComObject(xlApp);

            //   xlApp = null;
            //   xlWorkBooks = null;
            //   xlWorkBook = null;
            //   xlWorkSheet = null;
            //   xlRng = null;

        }

        public static Dictionary<string, string> ReadExcelCellValues(string readpath, string workSheet)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(readpath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(workSheet);
            //Excel.Range xlRng = (Excel.Range)xlWorkSheet.get_Range(Range, Type.Missing);
            
            //Excel.Application xlApp;
            //Excel.Workbook xlWorkBook;
            //Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            

            // object misValue = System.Reflection.Missing.Value;

            //xlApp = new Excel.Application();
            //Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(readpath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);


            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(workSheet);

            range = xlWorkSheet.UsedRange;

            Dictionary<string, string> dic = new Dictionary<string, string>();

            string str1, str2;

            int rCnt = 0;

            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {

                str1 = Convert.ToString((range.Cells[rCnt, 1] as Excel.Range).Value2);
            
                if (str1 == null)
                {
                    //str1 = rCnt.ToString();
                    continue;
                }
                str2 = Convert.ToString((range.Cells[rCnt, 2] as Excel.Range).Value2);
                dic.Add(str1, str2);

            }

            // Cleanup
            //GC.Collect();
            //GC.WaitForPendingFinalizers();

            //Marshal.FinalReleaseComObject(range);
            //Marshal.FinalReleaseComObject(xlWorkSheet);

            object misValue = System.Reflection.Missing.Value;
            xlWorkBook.Close(false, misValue, misValue);

            // xlWorkBook.Close(0);
            Marshal.ReleaseComObject(xlWorkBook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);

            return dic;

        }

        public static string ReadExcelBizName(String readpath)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
           

            string Bizname;



            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(readpath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);


            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CompanyFile");

            range = xlWorkSheet.UsedRange;

            Bizname = (string)(xlWorkSheet.Cells[2,1] as Excel.Range).Value;
            
           // xlWorkBook.Close();
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook.Close(false, misValue, misValue);

            xlApp.Quit();
            return Bizname;

        }


        //-------------Sunder Raj Added----
        public static void  ReadExcelSheet(string readpath, string workSheet,int rowno,ref List<string> stList)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;
            int rCnt = 0;
            int cCnt = 0;
            // object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(readpath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            //// Get worksheet names
            //foreach (Excel.Worksheet sh in xlWorkBook.Worksheets)
            //    Debug.WriteLine(sh.Name);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(workSheet);
            range = xlWorkSheet.UsedRange;

            for (rCnt = 1; rCnt <= rowno; rCnt++)
            {
                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                {
                    if(rCnt == rowno)
                    {
                        str = Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                        stList.Add(str);
                    }
                    
                }

            }

            // xlWorkBook.Save(); 
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook.Close(false, misValue, misValue);
            // xlWorkBook.Close();
            xlApp.Quit();


        }
    //-----End of Sunder Raj Code
        
        public static string GetOS()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem");
            string OS_Name = string.Empty;
            foreach (ManagementObject os in searcher.Get())
            {
                OS_Name = os["Caption"].ToString();
                break;
            }
            return OS_Name;
        }

        public static string GetRegPath ()
        {
            string regPath;
            if (Environment.Is64BitOperatingSystem)
            {
                regPath = "Software\\Wow6432Node\\Intuit\\QuickBooks\\";
            }
            else
            {
                regPath = "Software\\Intuit\\QuickBooks\\";
            }

            return regPath;
        }

        public static string GetRegVer (string SKU)
        {

            switch (SKU)
            {
                case "Enterprise":
                    reg_ver = "bel";
                    break;
                case "Enterprise Accountant":
                    reg_ver = "belacct";
                    break;

                case "Premier":
                case "Premier Plus":
                    reg_ver = "superpro";
                    break;

                case "Premier Accountant":
                    reg_ver = "accountant";
                    break;

                case "Pro":
                case "Pro Plus":
                    reg_ver = "pro";
                    break;
            }
            return reg_ver;
        }

        public static string GetQBVersion(string ver, string reg_ver)
        {
            Object QBVer;
            string installed_version = string.Empty;

            string regPath = GetRegPath();

            RegistryKey key = Registry.LocalMachine.OpenSubKey(regPath + ver + "\\" + reg_ver);
                if (key != null)
                {
                    QBVer = key.GetValue("QBVersion");
                    if (QBVer != null)
                    {
                        // Version version = new Version(o as String);  //"as" because it's REG_SZ...otherwise ToString() might be safe(r)
                        installed_version = QBVer as string;

                    }
                }
                 if (key == null)
                {
                    // Install_QB?
                     
                }
           return installed_version; 
          }

        public static string GetProduct(string ver, string reg_ver)
        {
            Object product;
            string installed_product = string.Empty;

            string regPath = GetRegPath();

            RegistryKey key = Registry.LocalMachine.OpenSubKey(regPath + ver + "\\" + reg_ver);
            if (key != null)
            {
               product = key.GetValue("Product");
               if (product != null)
               {
                  installed_product = product as string;
               }
            }

               
            if (key == null)
            {
               // Install_QB?
             }
          
            return installed_product;
          }

        public static string GetDataPath(string ver, string reg_ver)
        {
            Object dataPath;
            string installed_dataPath = string.Empty;
            
            string regPath = GetRegPath();
            RegistryKey key = Registry.LocalMachine.OpenSubKey(regPath + ver + "\\" + reg_ver);

            if (key != null)
            {
                 dataPath = key.GetValue("DataPath");
                 if (dataPath != null)
                 {
                    installed_dataPath = dataPath as string;
                 }
            }
                    
            if (key == null)
            {
                 // Install_QB?
            }
           
             return installed_dataPath;
          }

        public static string GetPath(string ver, string reg_ver)
        {
            Object QBPath;
            string installed_QBPath = string.Empty;
            
            string regPath = GetRegPath();

            RegistryKey key = Registry.LocalMachine.OpenSubKey(regPath + ver + "\\" + reg_ver);

            if (key != null)
            {
              QBPath = key.GetValue("Path");
              if (QBPath != null)
              {
                 installed_QBPath = QBPath as string;
              }
            }


            if (key == null)
            {
               // Install_QB?
            }
            
            return installed_QBPath;
        }

        public static string GetDataPathKey(string ver, string reg_ver)
        {
            Object dataPath = new Object();
            string dataPath_key = string.Empty;

            string regPath = GetRegPath();

            RegistryKey key = Registry.LocalMachine.OpenSubKey(regPath + ver + "\\" + reg_ver);
            if (key != null)
            {
               dataPath = key.GetValue("DataPath");
               if (dataPath != null)
               {
                    dataPath_key = dataPath as string;
               }
            }

            if (key == null)
            {
                 // Install_QB?
            }
          
            return dataPath_key;
        }

        public static string GetCommonFilesPath(string ver, string reg_ver)
        {
            Object commonFilesPath;
            string installed_commonPath = string.Empty;

            string regPath = GetRegPath();

            RegistryKey key = Registry.LocalMachine.OpenSubKey(regPath + ver + "\\" + reg_ver);
            if (key != null)
            {
               commonFilesPath = key.GetValue("CommonFilesPath");
               if (commonFilesPath != null)
               {
                  installed_commonPath = commonFilesPath as string;
               }
             }

            if (key == null)
            {
               // Install_QB?
            }
             return installed_commonPath;
          }
 
        public static void Update_Automation_Properties ()
        {
            string readpath = "C:\\Temp\\Parameters.xlsm";
            Dictionary<string, string> dic_QBDetails = new Dictionary<string, string>();
            string SKU, ver, reg_ver, data_path, installed_product, installed_path, installed_dir, regPath;


            //dic_QBDetails = File_Functions.ReadExcelValues(readpath, "PostInstall", "B2:B4");
            //ver = dic_QBDetails["B2"];
            //reg_ver = dic_QBDetails["B3"];

            dic_QBDetails = File_Functions.ReadExcelValues(readpath, "Install", "B8:B12");
            SKU = dic_QBDetails["B12"];
            ver = dic_QBDetails ["B8"];
            reg_ver = GetRegVer(SKU);

            regPath = File_Functions.GetRegPath();
            installed_product = File_Functions.GetProduct(ver, reg_ver);
            installed_path = File_Functions.GetPath(ver, reg_ver);
            data_path = File_Functions.GetDataPath(ver, reg_ver);
            installed_dir = Path.GetDirectoryName(installed_path); // Get the path (without the exe name)


            string curr_dir, aut_file;
            curr_dir = Directory.GetCurrentDirectory();
            aut_file = curr_dir + @"\Automation.Properties";

            List<string> prop_value = new List<string>(File.ReadAllLines(aut_file));
            int lineIndex = prop_value.FindIndex(line => line.StartsWith("QBExePath="));
            if (lineIndex != -1)
            {
                prop_value[lineIndex] = "QBExePath=" + installed_path;
                File.WriteAllLines(aut_file, prop_value);
            }

            lineIndex = prop_value.FindIndex(line => line.StartsWith("QBW.ini="));
            if (lineIndex != -1)
            {
                prop_value[lineIndex] = "QBW.ini=" + data_path + "qbw.ini";
                File.WriteAllLines(aut_file, prop_value);
            }

            //lineIndex = prop_value.FindIndex(line => line.StartsWith("LogDirectory="));
            //if (lineIndex != -1)
            //{
            //    prop_value[lineIndex] = "LogDirectory=";
            //    File.WriteAllLines(aut_file, prop_value);
            //}
           
        }

       }
    }
