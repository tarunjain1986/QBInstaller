using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FrameworkLibraries.Utils
{
    public class FileOperations
    {
        public static void DeleteAllFilesInDirectory(string dir)
        {
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            string[] filePaths = Directory.GetFiles(dir);
            foreach (string filePath in filePaths)
            {
                try
                {
                    File.GetAccessControl(filePath);
                    File.Delete(filePath);
                }
                catch(Exception)
                {
                }
            }

        }

        public static void DeleteCompanyFileInDirectory(string dir, string fileName)
        {
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            string[] filePaths = Directory.GetFiles(dir);
            foreach (string filePath in filePaths)
            {
                if (filePath.Contains(fileName))
                {
                    File.GetAccessControl(filePath);
                    File.Delete(filePath);
                }
            }

        }

        public static void CopyCompanyFilesToDirectory(string source, string destination)
        {
            string destinationFile = null;

            string[] filePaths = Directory.GetFiles(source);
            foreach (string filePath in filePaths)
            {
                string[] split = filePath.Split('\\');
                foreach(string s in split)
                {
                    if (s.Contains(".qbw") || s.Contains(".QBW") || s.Contains(".QBB") || s.Contains(".qbb") || s.Contains(".QBM") || s.Contains(".qbm"))
                    {
                        destinationFile = s;
                        break;
                    }
                }
                File.Copy(filePath, destination+destinationFile, true);
            }
        }

        public static void CopyCompanyFileToDirectory(string sourceDir, string destinationDir, string fileName)
        {
            if (!Directory.Exists(destinationDir))
                Directory.CreateDirectory(destinationDir);

            string destinationFile = null;

            string[] filePaths = Directory.GetFiles(sourceDir);
            foreach (string filePath in filePaths)
            {
                string[] split = filePath.Split('\\');
                foreach (string s in split)
                {
                    if (s.Contains(fileName))
                    {
                        destinationFile = s;
                        File.Copy(filePath, destinationDir + destinationFile, true);
                        break;
                    }
                }
            }
        }

        public static bool CheckForStringInFile(string filePath, string searchString)
        {
            if (File.ReadLines(filePath).Any(line => line.Contains(searchString)))
                return true;
            else
                return false;
        }

        public static void AppendStringToFile(string filePath, string appendString)
        {
            using (StreamWriter w = File.AppendText(filePath))
            {
                w.WriteLine(appendString);
                w.Close();
            }
        }

        public static bool FileEquals(string path1, string path2)
        {
            byte[] file1 = File.ReadAllBytes(path1);
            byte[] file2 = File.ReadAllBytes(path2);
            if (file1.Length == file2.Length)
            {
                for (int i = 0; i < file1.Length; i++)
                {
                    if (file1[i] != file2[i])
                    {
                        return false;
                    }
                }
                return true;
            }
            return false;
        }

    }
}
