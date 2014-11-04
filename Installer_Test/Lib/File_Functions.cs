﻿using System;
using System.IO;
using System.Linq;
using System.Management;
using System.Collections.Generic;

using Microsoft.Win32;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;

namespace Installer_Test.Lib
{
   
    public class File_Functions
    {
        public string expected_ver, installed_product, installed_dataPath, installed_commonPath;

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
            //string readpath = "C:\\Temp\\Parameters.xlsx"; // "C:\\Installation\\Sample.txt";
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
            return dic;
        }
        
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

        public static string GetQBVersion (string OS_Name, string ver, string reg_ver)
        {
            Object QBVer;
            string installed_version = string.Empty;
            RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\Intuit\\QuickBooks\\" + ver + "\\" + reg_ver);
            if (OS_Name.Contains("Windows 7"))
            {
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
                   
            }
            return installed_version; 
          }
  
        public static string GetProduct (string OS_Name, string ver, string reg_ver)
        {
            Object product;
            string installed_product = string.Empty;
            RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\Intuit\\QuickBooks\\" + ver + "\\" + reg_ver);
            if (OS_Name.Contains("Windows 7"))
            {
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

            }
                return installed_product;
          }

        public static string GetDataPath (string OS_Name, string ver, string reg_ver)
        {
            Object dataPath;
            string installed_dataPath = string.Empty;
            RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\Intuit\\QuickBooks\\" + ver + "\\" + reg_ver);
            if (OS_Name.Contains("Windows 7"))
            {
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

            }
                return installed_dataPath;
          }

        public static string GetDataPathKey(string OS_Name, string ver, string reg_ver)
        {
            Object dataPath = new Object();
            string dataPath_key = string.Empty;
            RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\Intuit\\QuickBooks\\" + ver + "\\" + reg_ver);
            if (OS_Name.Contains("Windows 7"))
            {
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

            }
            return dataPath_key;
        }

        public static string GetCommonFilesPath (string OS_Name, string ver, string reg_ver)
        {
            Object commonFilesPath;
            string installed_commonPath = string.Empty;
            RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\Intuit\\QuickBooks\\" + ver + "\\" + reg_ver);
            if (OS_Name.Contains("Windows 7"))
            {
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

            }
                return installed_commonPath;
          }

       }
    }
