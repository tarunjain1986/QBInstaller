using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using FrameworkLibraries;
using FrameworkLibraries.Utils;
using FrameworkLibraries.AppLibs.QBDT;
using FrameworkLibraries.ActionLibs.WhiteAPI;

using Installer_Test;
using Installer_Test.Lib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestStack.BDDfy;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems;

namespace Installer_Forms
{
   
        public partial class Form1 : Form
    {
        public static Property conf = Property.GetPropertyInstance();
        public static string Sync_Timeout = conf.get("SyncTimeOut");

        public string line, sel_rel, Build, skuFolder, ver, reg_ver;
        public string [] workFlow = { "Skip", "Signup", "Lookup" };
        public string [] installType = {"Local","Shared","Server"};
        public string [] SKU = { "Enterprise 30 User", "Enterprise 5 User MLI Active", "Enterprise Accountant 5 User", "Premier 3 User", "Premier Accountant", "Premier Accountant 3 User", "Premier Plus 3 User", "Pro", "Pro 3 User", "Pro Plus", "Pro Plus 3 User" };
        public string [] release, LicenseNo, QBVersion, buildNo, ProductNo;
        public string sourcePath, targetPath, fileName, sourceFile, destFile, customOpt, AVName;
        public string initialPath = @"\\mtvfsqbscm\Release\";
        public string License_No = "" , Product_No = "" , wkflow = "", UserID = "", Passwd = "", firstName = "", lastName = ""; 
              
     //   [Given]
        public Form1()
        {
          InitializeComponent();
          Invoke_Installer(); 
              
        }

    
        public void Invoke_Installer ()
        {
            string readpath = @"C:\Temp\Input_QB_Version.txt"; 

            File.WriteAllLines(readpath, File.ReadAllLines(readpath).Where(l => !string.IsNullOrWhiteSpace(l))); // Remove white space from the file

            string[] lines = File.ReadAllLines(readpath);
            var dic = lines.Select(line => line.Split('=')).ToDictionary(keyValue => keyValue[0], bits => bits[1]);

            sel_rel = dic["[Release]"];
                 
            release = sel_rel.Split(',');
                 
            releaseCB.Items.AddRange(release);
            workflowCB.Items.AddRange(workFlow);
            installationCB.Items.AddRange(installType);
            skuCB.Items.AddRange(SKU);
          
        }
                   
         /// <summary>
         /// Click on Start Installation
         /// </summary>
         /// <param name="sender"></param>
         /// <param name="e"></param>
            private void startPB_Click(object sender, EventArgs e)
            {
                
                License_No = licensenoTB.Text;
                Product_No = productnoTB.Text;
                Build = buildCB.Text;
                UserID = userIDTB.Text;
                Passwd = passwordTB.Text;
                firstName = firstnameTB.Text;
                lastName = lastnameTB.Text;
                wkflow = workflowCB.Text;
                customOpt = installationCB.Text;

                switch (skuCB.Text)
                {
                    case "Enterprise 30 User":
                    case "Enterprise 5 User MLI Active":
                    case "Enterprise Accountant 5 User":
                    skuFolder = "CD_BEL";
                    reg_ver = "bel";
                    break;

                    case "Premier 3 User":
                    case "Premier Accountant":
                    case "Premier Accountant 3 User":
                    skuFolder = "CD_SPRO";
                    reg_ver = "pre";
                    break;

                    case "Premier Plus 3 User":
                    skuFolder = "CD_SPROPLUS";
                    reg_ver = "superpro";
                    break;

                    case "Pro":
                    case "Pro 3 User":
                    skuFolder = "CD_PRO";
                    reg_ver = "pro";
                    break;

                    case "Pro Plus":
                    case "Pro Plus 3 User":
                    skuFolder = "CD_PROPLUS";
                    reg_ver = "superpro";
                    break;

                }
                   
              
                var regex = new Regex(@".{4}");
                string temp = regex.Replace(License_No, "$&" + "\n");
                LicenseNo = temp.Split('\n');

                regex = new Regex(@".{3}");
                temp = regex.Replace(Product_No, "$&" + "\n");
                ProductNo = temp.Split('\n');

                sourcePath = initialPath + releaseCB.Text + @"\" + versionCB.Text + @"\" + buildCB.Text + @"\" + skuFolder + @"\QBooks\";
                targetPath = @"C:\Installer_Build\" + releaseCB.Text + @"\" + versionCB.Text + @"\" + buildCB.Text + @"\" + skuFolder + @"\";


                if (!Directory.Exists(targetPath))
                {
                    File_Functions.DirectoryCopy(sourcePath, targetPath, true);
                }
                targetPath = targetPath + @"QBooks\";

                TextWriter tw = new StreamWriter(@"C:\Temp\Parameters.txt");

                switch (releaseCB.Text)
                {
                    case "Mango":
                    ver = "25.0";
                    break;

                    case "Ruby":
                    ver = "24.0";
                    break;

                    case "Nirvana":
                    ver = "23.0";
                    break;
                }

                // write lines of text to the file
                tw.WriteLine("Target Path=" + targetPath);
                tw.WriteLine("Version=" + ver);
                tw.WriteLine("Registry Folder=" + reg_ver);
                tw.WriteLine("Workflow=" + wkflow);
                tw.WriteLine("Installation Type=" + customOpt);
                tw.WriteLine("License No=" + License_No);
                tw.WriteLine("Product No=" + Product_No);
                tw.WriteLine("UserID=" + UserID);
                tw.WriteLine("Password=" + Passwd);
                tw.WriteLine("First Name=" + firstName);
                tw.WriteLine("Last Name=" + lastName);

                if (AntiVirusChkB.Checked == true) 
                {
                    tw.WriteLine("AntivirusTest=true");
                    AVName = "";
                    if (MSEInstallRB.Checked == true)
                    {
                        AVName = AVName + "MSEInstall.exe";
                    }

                    if (AvastRB.Checked == true)
                    {
                        AVName = AVName + "avast_internet_security_setup.exe";
                    }
                    if (AviraRB.Checked == true)
                    {
                        AVName = AVName + "avira_en_av___ws2.exe";
                    }
                    if (BitdefenderRB.Checked == true)
                    {
                        AVName = AVName + "bitdefender_antivirus.exe";
                    }
                    if (NodRB.Checked == true)
                    {
                        AVName = AVName + "eset_nod32_antivirus_live_installer_.exe";
                    }
                    tw.WriteLine("AntiVirusSW=" + AVName);
                }
                else
                {
                    tw.WriteLine("AntivirusTest=false");
                }
               
                // close the stream     
                tw.Close();

                this.Close();

            }

            private void releaseCB_SelectedIndexChanged(object sender, EventArgs e)
            {
                versionCB.Text = "";
                buildCB.Text = "";
                versionCB.Items.Clear();
                buildCB.Items.Clear();
                sourcePath = initialPath + releaseCB.Text + @"\";
                QBVersion = System.IO.Directory.GetDirectories(sourcePath);
              
                for (int i = 0; i < QBVersion.Length;i++ )
                {
                   QBVersion[i] = QBVersion[i].Replace(sourcePath, "");
                }
                        
                versionCB.Items.AddRange(QBVersion);

            }
                        
            private void versionCB_SelectedIndexChanged(object sender, EventArgs e)
            {
                buildCB.Text = "";
                buildCB.Items.Clear();
                sourcePath = initialPath + releaseCB.Text + @"\" + versionCB.Text + @"\";
                buildNo = System.IO.Directory.GetDirectories(sourcePath);

                for (int i = 0; i < buildNo.Length; i++)
                {
                    buildNo[i] = buildNo[i].Replace(sourcePath, "");
                }
                
                buildCB.Items.AddRange(buildNo);
            }

            private void workflowCB_SelectedIndexChanged(object sender, EventArgs e)
            {
                switch (workflowCB.Text)
                {
                    case "Skip":
                        userIDLbl.Visible = false;
                        userIDTB.Visible = false;
                        passwordLbl.Visible = false;
                        passwordTB.Visible = false;

                        firstLbl.Visible = false;
                        firstnameTB.Visible = false;
                        lastLbl.Visible = false;
                        lastnameTB.Visible = false;
                        break;

                    case "Signup":
                        userIDLbl.Visible = true;
                        userIDTB.Visible = true;
                        passwordLbl.Visible = true;
                        passwordTB.Visible = true;

                        firstLbl.Visible = true;
                        firstnameTB.Visible = true;
                        lastLbl.Visible = true;
                        lastnameTB.Visible = true;
                        break;

                    case "Lookup":
                        userIDLbl.Visible = true;
                        userIDTB.Visible = true;
                        passwordLbl.Visible = true;
                        passwordTB.Visible = true;

                        firstLbl.Visible = false;
                        firstnameTB.Visible = false;
                        lastLbl.Visible = false;
                        lastnameTB.Visible = false;
                        break;
                }
            }

            private void AntiVirusChkB_CheckedChanged(object sender, EventArgs e)
            {
                if (AntiVirusChkB.Checked == true)
                {
                    MSEInstallRB.Visible = true;
                    AvastRB.Visible = true;
                    AviraRB.Visible = true;
                    BitdefenderRB.Visible = true;
                    NodRB.Visible = true;
                }
                if (AntiVirusChkB.Checked == false)
                {
                    MSEInstallRB.Visible = false;
                    AvastRB.Visible = false;
                    AviraRB.Visible = false;
                    BitdefenderRB.Visible = false;
                    NodRB.Visible = false;
                    MSEInstallRB.Checked = false;
                    AvastRB.Checked = false;
                    AviraRB.Checked = false;
                    BitdefenderRB.Checked = false;
                    NodRB.Checked = false;

                }
            }

            private void installationCB_SelectedIndexChanged(object sender, EventArgs e)
            {
               switch (installationCB.Text)
               {
                case "Local":
                case "Shared":
                  licensenoTB.Visible = true;
                  productnoTB.Visible = true;
                  Licenselbl.Visible = true;
                  Productlbl.Visible = true;
                  Workflowlbl.Visible = true;
                  workflowCB.Visible = true;
                break;

                case "Server":
                  licensenoTB.Visible = false;
                  productnoTB.Visible = false;
                  Licenselbl.Visible = false;
                  Productlbl.Visible = false;
                  Workflowlbl.Visible = false;
                  workflowCB.Visible = false;
                break;
               }
            }
    }
}
