using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Brown_Prepress_Automation.Properties;
using System.Windows.Forms;
using System.Drawing;
using ExcelLibrary.SpreadSheet;
using System.Diagnostics;
using Ghostscript.NET.Processor;
using Ghostscript.NET;
using System.Net;
using PdfToImage;
using Microsoft.Win32;


namespace Brown_Prepress_Automation
{
    class MethodsCommon
    {
        public bool networkCheck()
        {
            bool status = true;
            List<string> paths = new List<string>();            
            paths.Add(Settings.Default.hotFolder);
            paths.Add(Settings.Default.errorFolder);
            paths.Add(Settings.Default.archiveFolder);
            

            foreach (var path in paths)
            {
                if (!Directory.Exists(path))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Please connect to " + path);
                    Console.ResetColor();
                    Console.WriteLine("-------------------------------------------------------------");
                    status = false;
                }
            }
            return status;
        }

        public int countValidCells(string filename, int startInt, int worksheet, int checkedColumn)
        {
            Workbook book = Workbook.Load(filename);
            Worksheet sheet = book.Worksheets[worksheet];
            while (!sheet.Cells[startInt, checkedColumn].IsEmpty)
            {
                startInt++;
            }
            return startInt;
        }

        public void SendToPrinter(string printFile, bool ticket, bool tabloid)
        {
            //GhostscriptVersionInfo gvi = new GhostscriptVersionInfo(new Version(0, 0, 0), System.IO.Path.Combine(Brown_Prepress_Automation.Properties.Resources.gsdll32.ToString()), string.Empty, GhostscriptLicense.GPL);
            GhostscriptVersionInfo gvi = new GhostscriptVersionInfo(new Version(0, 0, 0), System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "gsdll32.dll"), string.Empty, GhostscriptLicense.GPL);
            using (GhostscriptProcessor processor = new GhostscriptProcessor(gvi))
            {
                List<string> switches = new List<string>();
                switches.Add("-dPrinted");
                switches.Add("-dBATCH");
                switches.Add("-dNOPAUSE");
                switches.Add("-dNOSAFER");
                if (tabloid)
                {
                    switches.Add("-g792x1224");
                }
                if (ticket)
                {
                    switches.Add("-dPDFFitPage");

                }
                switches.Add("-dNumCopies=1");
                switches.Add("-sDEVICE=mswinpr2");
                switches.Add(Convert.ToString("-sOutputFile=%printer%") + Settings.Default.printer);
                switches.Add("-f");
                switches.Add(printFile);
                processor.StartProcessing(switches.ToArray(), null);
            }
        }

        public void SendToPrinter11x17(string printFile, bool ticket, bool tabloid)
        {
            //GhostscriptVersionInfo gvi = new GhostscriptVersionInfo(new Version(0, 0, 0), System.IO.Path.Combine(Brown_Prepress_Automation.Properties.Resources.gsdll32.ToString()), string.Empty, GhostscriptLicense.GPL);
            GhostscriptVersionInfo gvi = new GhostscriptVersionInfo(new Version(0, 0, 0), System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "gsdll32.dll"), string.Empty, GhostscriptLicense.GPL);
            using (GhostscriptProcessor processor = new GhostscriptProcessor(gvi))
            {
                List<string> switches = new List<string>();
                switches.Add("-dPrinted");
                switches.Add("-dBATCH");
                switches.Add("-dNOPAUSE");
                switches.Add("-dNOSAFER");
                if (tabloid)
                {
                    switches.Add("-g792x1224");
                }
                if (ticket)
                {
                    switches.Add("-dPDFFitPage");
                    
                }
                switches.Add("-dNumCopies=1");
                switches.Add("-sDEVICE=mswinpr2");
                switches.Add(Convert.ToString("-sOutputFile=%printer%") + Settings.Default.printer);
                switches.Add("-f");
                switches.Add(printFile);
                processor.StartProcessing(switches.ToArray(), null);
            }
        }

        public void updateCheck()
        {
            FileInfo setupFile = new FileInfo("\\\\192.168.19.5\\Programs\\Brown Prepress Automation.msi");
            FileInfo appFile = new FileInfo("\\Temp\\Brown Prepress Automation.msi");
            FileInfo updateFile = new FileInfo(Brown_Prepress_Automation.FormMain.Globals.appDir + "\\Brown Prepress Automation Updater.exe");
            updateFile.CopyTo("\\Temp\\Brown Prepress Automation Updater.exe", true);
            
            string app = Application.ProductVersion.ToString();
            if (appFile.Exists)
            {
                if (setupFile.LastWriteTime > appFile.LastWriteTime)
                {
                    Process i = new Process();
                    i.StartInfo.FileName = "\\Temp\\Brown Prepress Automation Updater.exe";
                    i.Start();
                    Environment.Exit(0);
                }
            }
            else
            {
                setupFile.CopyTo("\\Temp\\Brown Prepress Automation.msi", true);
            }
        }

        public void xmlDownloadTuftex(string wo)
        {
            WebClient Client = new WebClient();
            string file = "http://salsaprd.shawinc.com/SALSAWeb/ServiceRequest?service=getSpec&id=" + wo + "&nostatusupdate=true";

            Client.DownloadFile(file, Settings.Default.hotFolder + "\\" + wo + " tuftex.xml");

            Console.WriteLine(file);
        }

        public void xmlDownloadShaw(string wo, string location)
        {
            WebClient Client = new WebClient();
            string file = "http://salsaprd.shawinc.com/SALSAWeb/ServiceRequest?service=getSpec&id=" + wo + "&nostatusupdate=true";

            Client.DownloadFile(file, location);

            Console.WriteLine(file);
        }

        public void jpgCreate(string input, string output, int quality, int xRes, int yRes, int firstPage, int lastPage)
        {            
            PdfToImage.PDFConvert pp = new PDFConvert();
            pp.OutputFormat = "jpegcmyk"; //format
            pp.JPEGQuality = quality; //100% quality
            pp.ResolutionX = xRes; //dpi
            pp.ResolutionY = yRes;
            pp.FirstPageToConvert = firstPage; //pages you want
            pp.LastPageToConvert = lastPage;
            pp.Convert(input, output);
        }

        public void tifCreate(string input, string output, int quality, int xRes, int yRes, int firstPage, int lastPage)
        {
            PdfToImage.PDFConvert pp = new PDFConvert();
            pp.OutputFormat = "tiff32nc"; //format
            pp.JPEGQuality = quality; //100% quality
            pp.ResolutionX = xRes; //dpi
            pp.ResolutionY = yRes;
            pp.FirstPageToConvert = firstPage; //pages you want
            pp.LastPageToConvert = lastPage;
            pp.Convert(input, output);
        }

        public bool IsFileLocked(FileInfo file)
        {
            try
            {
                using (FileStream stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }

            //file is not locked
            return false;
        }  

        public void SetString(string Key, string Value)
        {
            RegistryKey myHive = Registry.CurrentUser.CreateSubKey("Software\\Matt\\Brown Prepress Automation");
            myHive.SetValue(Key, Value, RegistryValueKind.String);
            myHive.Close();
        }

        public void SetStringCustomer(string Key, string Value, string customer)
        {
            RegistryKey myHive = Registry.CurrentUser.CreateSubKey("Software\\Matt\\Brown Prepress Automation\\" + customer);
            myHive.SetValue(Key, Value, RegistryValueKind.String);
            myHive.Close();
        }

        public void SetBool(string Key, bool Value)
        {
            RegistryKey myHive = Registry.CurrentUser.CreateSubKey("Software\\Matt\\Brown Prepress Automation");
            myHive.SetValue(Key, Convert.ToString(Value), RegistryValueKind.String);
            myHive.Close();
        }

        public void SetBoolCustomer(string Key, bool Value, string customer)
        {
            RegistryKey myHive = Registry.CurrentUser.CreateSubKey("Software\\Matt\\Brown Prepress Automation\\" + customer);
            myHive.SetValue(Key, Convert.ToString(Value), RegistryValueKind.String);
            myHive.Close();
        }

        public object GetValue(string Key)
        {
            using (RegistryKey myHive = Registry.CurrentUser.OpenSubKey("Software\\Matt\\Brown Prepress Automation", true))
            {
                if (myHive != null)
                {
                    return myHive.GetValue(Key);
                }
                else
                {
                    return null;
                }
            }
        }

        public object GetValueCustomer(string Key, string customer)
        {
            using (RegistryKey myHive = Registry.CurrentUser.OpenSubKey("Software\\Matt\\Brown Prepress Automation\\" + customer, true))
            {
                if (myHive != null)
                {
                    return myHive.GetValue(Key);
                }
                else
                {
                    return null;
                }
            }
        }

        public void SettingsDelete()
        {
            using (RegistryKey myHive = Registry.CurrentUser.OpenSubKey("Software\\Matt\\", true))
            {
                myHive.DeleteSubKeyTree("Brown Prepress Automation");
            }
        }

        public long DirectorySize(DirectoryInfo dInfo, bool includeSubDir)
        {
            // Enumerate all the files
            long totalSize = dInfo.EnumerateFiles()
                         .Sum(file => file.Length);

            // If Subdirectories are to be included
            if (includeSubDir)
            {
                // Enumerate all sub-directories
                totalSize += dInfo.EnumerateDirectories()
                         .Sum(dir => DirectorySize(dir, true));
            }
            return totalSize;
        }

        public void XlsWrite(int row, string name, string page, string fileName)
        {
            Workbook book = Workbook.Load(fileName);
            Worksheet sheet = book.Worksheets[0];
            name = name.Replace(" - ", " [ ");
            name = name.Replace("-", " ");
            name = name.Replace(" [ ", " - ");
            if (page.Trim() == "")
            {
                sheet.Cells[row, 0] = new Cell(name);                
            }
            else
            {
                sheet.Cells[row, 0] = new Cell(name + " - " + page);
            }
            for (int i = 0; i < 100; i++)
            {
                sheet.Cells[i, 20] = new Cell("");
            }
            book.Save(fileName);
        }        
    }
}
