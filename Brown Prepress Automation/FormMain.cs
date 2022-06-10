using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using Brown_Prepress_Automation.Properties;
using System.Net;
using System.Diagnostics;
using System.Threading;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using ExcelLibrary.SpreadSheet;

namespace Brown_Prepress_Automation
{
    public partial class FormMain : Form
    {
        //External Classes
        MethodsCommon methods = new MethodsCommon();
        DownloadShaw downloadShaw = new DownloadShaw();
        PdfProcessing pdfProcessing = new PdfProcessing();
        ShawParse shawParse = new ShawParse();

        //Lists
        List<string> hotfolderFiles = new List<string>();
        List<string> hotfolderFilesDownload = new List<string>();
        List<string> shawHotfolderFiles = new List<string>();
        List<string> nourisonHotfolderFiles = new List<string>();
        List<string> armstrongHotfolderFiles = new List<string>();
        List<string> prepressHotfolderFiles = new List<string>();
        string prepressHotfolder = Settings.Default.prepressHotfolder;
        DirectoryInfo dInfoSize = new DirectoryInfo(Settings.Default.tempDir);
        public class Globals
        {
            //Publc Global Variables
            public static string appDir = Directory.GetCurrentDirectory() + "\\";
            public static List<string> shawCheckList = new List<string>();
        }

        public FormMain()
        {
            InitializeComponent();

            DateTime today = DateTime.Today;
            this.Text += " Last Reset (" + today.ToString("d") + ")";

            if (Settings.Default.debugOn)
            {
                button1.Visible = true;
            }
            else
            {
                button1.Visible = false;
            }

            //Create Temp Dir if it does not Exist
            if (!Directory.Exists(Settings.Default.tempDir))
            {
                try
                {
                    Directory.CreateDirectory(Settings.Default.tempDir);
                }

                catch (IOException ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            //Check for Update on startup to program if update is checked and debug is off
            if (Settings.Default.updateCheck == true && Settings.Default.debugOn == false)
            {
                methods.updateCheck();
            }

            //Check Network connections before running the hotfolders
            if (methods.networkCheck() == true)
            {
                //Run Hotfolder Check
                tMain.Tick += new EventHandler(hotFolderParse);                
            }
            else
            {
                //Server Connection Error
                rtMain.AppendText(DateTime.Now + " | Connect to the listed servers and restart the program.\r\n", Color.Red, FontStyle.Bold);
            }
        }

        private void hotFolderParse(Object source, EventArgs e)
        {
            tMain.Stop();
            hotfolderFiles.Clear();
            hotfolderFilesDownload.Clear();
            shawHotfolderFiles.Clear();
            nourisonHotfolderFiles.Clear();
            armstrongHotfolderFiles.Clear();
            prepressHotfolderFiles.Clear();
            Preflight preflight = new Preflight();
            try
            {
                diskSize();

                //Check files in the hotfolder directory and add supported file types to the to do list
                DirectoryInfo dinfo = new DirectoryInfo(Settings.Default.hotFolder);
                FileInfo[] Files = dinfo.GetFiles("*.xls").Union(dinfo.GetFiles("*.xml")).Union(dinfo.GetFiles("*.pdf")).ToArray();

                foreach (FileInfo file in Files)
                {
                    if (file.FullName.ToLower().Contains("pdf"))
                    {
                        hotfolderFiles.Add(file.Name);
                    }
                    /*
                    else if (file.FullName.ToLower().Contains("xls"))
                    {
                        Settings.Default.workerWait = true;
                        hotfolderFiles.Add(file.Name);
                        preflight.Download(this, file.Name);
                        tMain.Stop();
                    }*/
                }

                DirectoryInfo dinfoDownload = new DirectoryInfo(Settings.Default.shawHotfolder + "download//");
                FileInfo[] FilesDownload = dinfoDownload.GetFiles("*.xls").Union(dinfoDownload.GetFiles("*.xml")).Union(dinfoDownload.GetFiles("*.pdf")).ToArray();
                foreach (FileInfo file in FilesDownload)
                {
                    if (file.FullName.ToLower().Contains("pdf"))
                    {
                        hotfolderFilesDownload.Add(file.Name);
                    }
                    else if (file.FullName.ToLower().Contains("xls"))
                    {
                        while (!Settings.Default.workerWait)
                        {
                            if (File.Exists(file.FullName))
                            {
                                Workbook book = Workbook.Load(Settings.Default.shawHotfolder + "download\\\\" + file.Name);
                                Worksheet sheet = book.Worksheets[0];
                                hotfolderFilesDownload.Add(file.Name);
                                preflight.Download(this, file.Name);
                                if ((!file.Name.ToLower().Contains("board")) && (sheet.Cells[1, 0].StringValue.Trim().Contains("http")))
                                {
                                    Settings.Default.workerWait = true;
                                }
                                Thread.Sleep(1000);                                
                            }
                            else
                            {
                                break;
                            }
                        }
                    }                    
                }

                if (hotfolderFiles.Count == 0 && hotfolderFilesDownload.Count == 0)
                {
                    DirectoryInfo dinfo2 = new DirectoryInfo(Settings.Default.shawHotfolder);
                    FileInfo[] shawFiles = dinfo2.GetFiles("*.xls").Union(dinfo2.GetFiles("*.xml")).Union(dinfo2.GetFiles("*.pdf")).ToArray();

                    foreach (FileInfo file in shawFiles)
                    {
                        if (file.FullName.ToLower().Contains("xlsx"))
                        {
                            using (StreamWriter errorFile = new StreamWriter(Settings.Default.shawErrorFolder + "\\" + file.Name + ".txt", true))
                            {
                                rtMain.AppendText(DateTime.Now + " | Please convert file " + file.Name + " to an xls.\r\n", Color.Red, FontStyle.Regular);
                                errorFile.WriteLine(DateTime.Now + " | Please convert file " + file.Name + " to an xls.\r\n");
                            }
                            if (File.Exists(Settings.Default.shawErrorFolder + "\\" + file.Name))
                            {
                                File.Delete(Settings.Default.shawErrorFolder + "\\" + file.Name);
                            }
                            if (File.Exists(Settings.Default.shawHotfolder + "\\" + file.Name))
                            {
                                File.Move(Settings.Default.shawHotfolder + "\\" + file.Name, Settings.Default.shawErrorFolder + "\\" + file.Name);
                            }
                        }
                        else if (file.FullName.ToLower().Contains("xls"))
                        {
                            shawHotfolderFiles.Add(file.Name);
                        }
                        else if (file.FullName.ToLower().Contains("xml"))
                        {
                            shawHotfolderFiles.Add(file.Name);
                        }
                    }

                    DirectoryInfo dinfo4 = new DirectoryInfo(Settings.Default.nourisonHotfolder);
                    FileInfo[] nourisonFiles = dinfo4.GetFiles("*.xls").Union(dinfo4.GetFiles("*.xml")).Union(dinfo4.GetFiles("*.pdf")).ToArray();

                    foreach (FileInfo file in nourisonFiles)
                    {
                        if (file.Extension.ToLower().Contains("xlsx"))
                        {
                            using (StreamWriter errorFile = new StreamWriter(Settings.Default.nourisonErrorFolder + "\\" + file.Name + ".txt", true))
                            {
                                rtMain.AppendText(DateTime.Now + " | Please convert file " + file.Name + " to an xls.\r\n", Color.Red, FontStyle.Regular);
                                errorFile.WriteLine(DateTime.Now + " | Please convert file " + file.Name + " to an xls.\r\n");
                            }
                            if (File.Exists(Settings.Default.nourisonErrorFolder + "\\" + file.Name))
                            {
                                File.Delete(Settings.Default.nourisonErrorFolder + "\\" + file.Name);
                            }
                            if (File.Exists(Settings.Default.nourisonHotfolder + "\\" + file.Name))
                            {
                                File.Move(Settings.Default.nourisonHotfolder + "\\" + file.Name, Settings.Default.nourisonErrorFolder + "\\" + file.Name);
                            }
                        }
                        else if (file.FullName.ToLower().Contains("xls"))
                        {
                            nourisonHotfolderFiles.Add(file.Name);
                        }
                        else if (file.FullName.ToLower().Contains("xml"))
                        {
                            nourisonHotfolderFiles.Add(file.Name);
                        }
                    }

                    DirectoryInfo dinfo5 = new DirectoryInfo(Settings.Default.armstrongHotfolder);
                    FileInfo[] armstrongFiles = dinfo5.GetFiles("*.xls").Union(dinfo5.GetFiles("*.xml")).Union(dinfo5.GetFiles("*.pdf")).ToArray();

                    foreach (FileInfo file in armstrongFiles)
                    {
                        if (file.Extension.ToLower().Contains("xlsx"))
                        {
                            using (StreamWriter errorFile = new StreamWriter(Settings.Default.armstrongErrorFolder + "\\" + file.Name + ".txt", true))
                            {
                                rtMain.AppendText(DateTime.Now + " | Please convert file " + file.Name + " to an xls.\r\n", Color.Red, FontStyle.Regular);
                                errorFile.WriteLine(DateTime.Now + " | Please convert file " + file.Name + " to an xls.\r\n");
                            }
                            if (File.Exists(Settings.Default.armstrongErrorFolder + "\\" + file.Name))
                            {
                                File.Delete(Settings.Default.armstrongErrorFolder + "\\" + file.Name);
                            }
                            if (File.Exists(Settings.Default.armstrongHotfolder + "\\" + file.Name))
                            {
                                File.Move(Settings.Default.armstrongHotfolder + "\\" + file.Name, Settings.Default.armstrongErrorFolder + "\\" + file.Name);
                            }
                        }
                        else if (file.Extension.ToLower().Contains("xls"))
                        {
                            armstrongHotfolderFiles.Add(file.Name);
                        }
                        else if (file.Extension.ToLower().Contains("xml"))
                        {
                            armstrongHotfolderFiles.Add(file.Name);
                        }
                    }

                    DirectoryInfo dinfo6 = new DirectoryInfo(prepressHotfolder);
                    FileInfo[] prepressFiles = dinfo6.GetFiles("*.xls").ToArray();

                    foreach (FileInfo file in prepressFiles)
                    {
                        if (file.Extension.ToLower().Contains("xlsx"))
                        {
                            if (File.Exists(prepressHotfolder + "\\" + file.Name))
                            {
                                File.Delete(prepressHotfolder + "\\" + file.Name);
                            }
                        }
                        else if (file.Extension.ToLower().Contains("xls"))
                        {
                            prepressHotfolderFiles.Add(file.Name);
                        }
                    }
                }

                //List all files added to the todo list.
                for (int i = 0; i < hotfolderFiles.Count; i++)
                {
                    rtMain.AppendText(DateTime.Now + " | Added file: " + hotfolderFiles[i] + "\r\n", Color.Black, FontStyle.Regular);
                }
                for (int i = 0; i < hotfolderFilesDownload.Count; i++)
                {
                    rtMain.AppendText(DateTime.Now + " | Added Download file: " + hotfolderFilesDownload[i] + "\r\n", Color.Black, FontStyle.Regular);
                }
                for (int i = 0; i < shawHotfolderFiles.Count; i++)
                {
                    rtMain.AppendText(DateTime.Now + " | Added Shaw file: " + shawHotfolderFiles[i] + "\r\n", Color.Black, FontStyle.Regular);
                }
                for (int i = 0; i < nourisonHotfolderFiles.Count; i++)
                {
                    rtMain.AppendText(DateTime.Now + " | Added Nourison file: " + nourisonHotfolderFiles[i] + "\r\n", Color.Black, FontStyle.Regular);
                }
                for (int i = 0; i < armstrongHotfolderFiles.Count; i++)
                {
                    rtMain.AppendText(DateTime.Now + " | Added Armstrong file: " + armstrongHotfolderFiles[i] + "\r\n", Color.Black, FontStyle.Regular);
                }
                for (int i = 0; i < prepressHotfolderFiles.Count; i++)
                {
                    rtMain.AppendText(DateTime.Now + " | Added Prepress Log file: " + prepressHotfolderFiles[i] + "\r\n", Color.Black, FontStyle.Regular);
                }

                tMain.Stop();

                ////////////////////////////
                //Do Work///////////////////
                ////////////////////////////
                //Start separate thread for todo list
                if (hotfolderFiles.Count != 0 && !bgwMain.IsBusy)
                {
                    object[] hotfolderArgs = { hotfolderFiles.ToArray(), hotfolderFiles.Count(), "none" };
                    bgwMain.RunWorkerAsync(hotfolderArgs);
                    hotfolderFiles.Clear();
                }
                if (hotfolderFilesDownload.Count != 0 && !bgwMain.IsBusy)
                {
                    object[] hotfolderArgs = { hotfolderFilesDownload.ToArray(), hotfolderFilesDownload.Count(), "download" };
                    bgwMain.RunWorkerAsync(hotfolderArgs);
                    hotfolderFilesDownload.Clear();
                }
                if (shawHotfolderFiles.Count != 0 && !bgwMain.IsBusy)
                {
                    object[] hotfolderArgs = { shawHotfolderFiles.ToArray(), shawHotfolderFiles.Count(), "shaw" };
                    bgwMain.RunWorkerAsync(hotfolderArgs);
                    shawHotfolderFiles.Clear();
                }
                if (nourisonHotfolderFiles.Count != 0 && !bgwMain.IsBusy)
                {
                    object[] hotfolderArgs = { nourisonHotfolderFiles.ToArray(), nourisonHotfolderFiles.Count(), "nourison" };
                    bgwMain.RunWorkerAsync(hotfolderArgs);
                    nourisonHotfolderFiles.Clear();
                }
                if (armstrongHotfolderFiles.Count != 0 && !bgwMain.IsBusy)
                {
                    object[] hotfolderArgs = { armstrongHotfolderFiles.ToArray(), armstrongHotfolderFiles.Count(), "armstrong" };
                    bgwMain.RunWorkerAsync(hotfolderArgs);
                    armstrongHotfolderFiles.Clear();
                }
                if (prepressHotfolderFiles.Count != 0 && !bgwMain.IsBusy)
                {
                    object[] hotfolderArgs = { prepressHotfolderFiles.ToArray(), prepressHotfolderFiles.Count(), "prepress" };
                    bgwMain.RunWorkerAsync(hotfolderArgs);
                    prepressHotfolderFiles.Clear();
                }
                if (!bgwDownload.IsBusy)
                {
                    tMain.Start();
                }

                //////////////////////////////////////////////////

                //Check for update while running
                if (Settings.Default.updateCheck && !Settings.Default.debugOn)
                {
                    methods.updateCheck();
                }
            }
            catch (Exception ex)
            {
                rtMain.AppendText(DateTime.Now + " | " + ex.Message + "\r\n", Color.Red, FontStyle.Regular);
                hotfolderFiles.Clear();
                tMain.Start();
            }
        }

        private void diskSize()
        {
            double sizeOfDir = (double)methods.DirectorySize(dInfoSize, true) / (1024 * 1024);
            lSize.Text = sizeOfDir.ToString("0.00") + " MB";
            double sizeOfDirGb = (double)methods.DirectorySize(dInfoSize, true) / (1024 * 1024 * 1024);
            int pValue = (int)(((sizeOfDirGb) / 25) * 100);
            if (pValue > 100)
            {
                pValue = 100;
            }
            pbSize.Value = pValue;

            if (pbSize.Value > 90)
            {
                pbSize.SetState(2);
            }
            else if (pbSize.Value > 70)
            {
                pbSize.SetState(3);
            }
            else
            {
                pbSize.SetState(1);
            }
        }

        //Start Parsing Button
        private void bStart_Click(object sender, EventArgs e)
        {
            bClearTemp.Enabled = false;
            bClearTemp.Visible = false;
            lSize.Visible = true;
            bStart.Visible = false;
            bStop.Visible = true;
            bSettings.Enabled = false;
            pbSize.Visible = true;
            /*if (Settings.Default.debugOn)
            {
                lSpeed.Visible = true;
                lPerc.Visible = true;
                lDownloaded.Visible = true;
            }
            else
            {
                lSpeed.Visible = false;
                lPerc.Visible = false;
                lDownloaded.Visible = false;
            }*/
            rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
            rtMain.AppendText(DateTime.Now + " | HotFolder Parsing Started...\r\n", Color.Black, FontStyle.Regular);
            rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
            tMain.Start();
            diskSize();
        }

        //Stop Parsing Button
        private void bStop_Click(object sender, EventArgs e)
        {
            pbSize.Visible = false;
            bClearTemp.Enabled = true;
            bClearTemp.Visible = true;
            lSize.Visible = false;
            bStop.Visible = false;
            bStart.Visible = true;
            bSettings.Enabled = true;
            try
            {
                //webClient.CancelAsync();
            }
            catch (Exception webe)
            {
                Console.WriteLine(webe.Message);
            }
            rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
            rtMain.AppendText(DateTime.Now + " | HotFolder Parsing Stopped...\r\n", Color.Black, FontStyle.Regular);
            rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
            tMain.Stop();
            diskSize();
        }

        //Open Setting Button
        private void bSettings_Click(object sender, EventArgs e)
        {
            rtMain.AppendText(DateTime.Now + " | Settings Opened\r\n", Color.Black, FontStyle.Regular);
            FormSettings settingForm = new FormSettings(this);
            settingForm.Show();
        }

        //Background Worker for todo list
        public void bgwMain_DoWork(object sender, DoWorkEventArgs e)
        {
            Invoke(new Action(() => { pbMain.Value = 0; }));
            Invoke(new Action(() => { pbIndividual.Value = 0; }));

            //Stop parsing while current todo list is running
            tMain.Stop();
            //diskSize();
            object[] arg = e.Argument as object[];
            string[] passedArray = (string[])arg[0];
            string hotfolderCustomer = (string)arg[2];
            List<string> passedList = passedArray.ToList();
            int fileProgressStep = (int)Math.Ceiling(((double)100) / (int)arg[1]);
            Preflight preflight = new Preflight();
            PreflightTuftexXml preflightTuftex = new PreflightTuftexXml();
            PreflightTuftexMiscOld preflightTuftexMiscOld = new PreflightTuftexMiscOld();
            PreflightArmstrong armstrong = new PreflightArmstrong();
            PreflightNourison nourison = new PreflightNourison();
            PreflightShaw shaw = new PreflightShaw();
            PrepressLog prepressLog = new PrepressLog();
            DownloadShaw downloadShaw = new DownloadShaw();
            ModelArmstrong.ArmstrongDB armstrongDB = new ModelArmstrong.ArmstrongDB();
            armstrongDB = GetArmstrongDB();

            foreach (string runfile in passedList)
            {
                try
                {
                    if (hotfolderCustomer == "none")
                    {
                        if (runfile.ToLower().Contains("pdf"))
                        {
                            preflight.PreflightPdf(this, runfile);
                        }
                    }
                    else if (hotfolderCustomer == "download")
                    {
                        if (runfile.ToLower().Contains("pdf"))
                        {
                            preflight.PreflightPdf(this, runfile);
                        }
                        else
                        {
                            while (Settings.Default.workerWait)
                            {
                            }
                        }
                        //preflight.Download(this, runfile);
                    }
                    else if (hotfolderCustomer == "shaw")
                    {
                        shaw.Preflight(this, runfile);
                    }
                    else if (hotfolderCustomer == "armstrong")
                    {                        
                        armstrong.Preflight(this, runfile, armstrongDB);
                    }
                    else if (hotfolderCustomer == "nourison")
                    {
                        nourison.PreflightNourisonPop(this, runfile);
                    }
                    else if (hotfolderCustomer == "prepress")
                    {
                        prepressLog.Log(this, runfile);
                    }
                    else if (runfile.ToLower().Contains("pdf"))
                    {
                        preflight.PreflightPdf(this, runfile);
	                }
                    /* LEGACY
                    else if (runfile.ToLower().Contains("tuftex"))
                    {
                        preflight.PreflightTuftex(this, runfile);
                    }
                    else if (runfile.ToLower().Contains("xml"))
                    {
                        preflight.PreflightTuftex(this, runfile);
                    }*/
                    else
                    {
                        throw new Exception("Unsupported Customer.  Did you forget to add the customer name to the file name? Example \"Shaw\"");
                    }

                    //Archive everything that is not a pdf file type since they are only temporary
                    if (!runfile.ToLower().Contains("pdf") && hotfolderCustomer != "shaw" && hotfolderCustomer != "armstrong" && hotfolderCustomer != "nourison" && hotfolderCustomer != "download")
                    {
                        Directory.CreateDirectory(Settings.Default.archiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\");
                        File.Copy(Settings.Default.hotFolder + runfile, Settings.Default.archiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + runfile, true);
                    }
                    if (!runfile.ToLower().Contains("pdf") && hotfolderCustomer == "download" && hotfolderCustomer != "armstrong" && hotfolderCustomer != "nourison")
                    {
                        Directory.CreateDirectory(Settings.Default.shawArchiveFolder + "download\\");
                        Directory.CreateDirectory(Settings.Default.shawArchiveFolder + "download\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\");
                        if (File.Exists(Settings.Default.shawHotfolder + "download\\" + runfile))
                        {
                            File.Copy(Settings.Default.shawHotfolder + "download\\" + runfile, Settings.Default.shawArchiveFolder + "download\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + runfile, true);
                            File.Delete(Settings.Default.shawHotfolder + "download\\" + runfile);
                        }
                        //passedList.Clear();
                    }
                    if (File.Exists(Settings.Default.hotFolder + "\\" + runfile))
                    {
                        File.Delete(Settings.Default.hotFolder + "\\" + runfile);
                    }
                    
                    /*if (File.Exists(Settings.Default.shawHotfolder + "\\" + runfile))
                    {
                        System.IO.File.Delete(Settings.Default.shawHotfolder + "\\" + runfile);
                    }*/
                    bgwMain.ReportProgress(fileProgressStep);
                }
                catch (Exception workerError)
                {
                    Invoke(new Action(() => { rtMain.AppendText(DateTime.Now + " | " + workerError.Message + ". \r\n", Color.Red, FontStyle.Regular); }));
                    bgwMain.ReportProgress(fileProgressStep);
                }
            }
            if (passedList.Count > 0)
            {
                e.Result = "Done";                
            }
        }

        private void bgwMain_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbMain.Step = e.ProgressPercentage;
            pbMain.PerformStep();
        }

        private void bgwMain_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                Invoke(new Action(() => { rtMain.AppendText(DateTime.Now + " | " + (string)e.Error.Message + "Error. \r\n\r\n", Color.Red, FontStyle.Regular); }));
                tMain.Start();
            }
            else
            {
                if ((string)e.Result == "Done")
                {
                    Invoke(new Action(() => { rtMain.AppendText(DateTime.Now + " | " + "Files Processed. \r\n\r\n", Color.Black, FontStyle.Regular); }));
                }
                tMain.Start();
            }
        }

        private void rtMain_TextChanged(object sender, EventArgs e)
        {
            rtMain.SelectionStart = rtMain.Text.Length;
            rtMain.ScrollToCaret();
        }

        private void bClearTemp_Click(object sender, EventArgs e)
        {
            //Try to Delete Temp Dir
            if (Directory.Exists(Settings.Default.tempDir))
            {
                try
                {
                    Directory.Delete(Settings.Default.tempDir, true);
                }

                catch (IOException ex)
                {
                    rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
                    rtMain.AppendText(DateTime.Now + " | " + ex.Message, Color.Red, FontStyle.Regular);
                    rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
                }
            }

            //Try to Re-Create Temp Dir with necessary contents
            try
            {
                Directory.CreateDirectory(Settings.Default.tempDir);
                Directory.CreateDirectory(Settings.Default.tempDir + "jpgs\\");
                Directory.CreateDirectory(Settings.Default.tempDir + "Shaw\\");
                File.Copy(Globals.appDir + "Images\\Blank.jpg", "\\Temp\\jpgs\\Blank.jpg", true);
                File.Copy(Globals.appDir + "Images\\Blank.pdf", "\\Temp\\Blank.pdf", true);
            }
            catch (IOException ex)
            {
                rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
                rtMain.AppendText(DateTime.Now + " | " + ex.Message, Color.Red, FontStyle.Regular);
                rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
            }

            //Run Update Check
            if (Settings.Default.updateCheck == true && Settings.Default.debugOn == false)
            {
                methods.updateCheck();
            }
            rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
            rtMain.AppendText(DateTime.Now + " | Temp Folder Cleared...\r\n", Color.Black, FontStyle.Regular);
            rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);

            diskSize();
        }

        WebClient webClient;               // WebClient that will be doing the downloading
        Stopwatch sw = new Stopwatch();    // Stopwatch to calculate the download speed
        string httpType = "https://";
        string downloadName = "";
        string tempPath = Settings.Default.tempDir + "Shaw";
        List<string> urlAddressList = new List<string>();
        bool isSSGlobal = false;
        public void DownloadFile(List<string> urlList, string extension, bool isSS)
        {
            isSSGlobal = isSS;
            urlAddressList = urlList;
            using (webClient = new WebClient())
            {
                if (urlAddressList.Count > 0)
                {
                    string urlAddress = urlAddressList[0];
                    string fileName = urlAddress.Split('/').Last();
                    
                    downloadName = fileName;
                    webClient.DownloadFileCompleted += new AsyncCompletedEventHandler(Completed);
                    webClient.DownloadProgressChanged += new DownloadProgressChangedEventHandler(ProgressChanged);

                    // The variable that will be holding the url address (making sure it starts with http://)
                    Uri URL = urlAddress.StartsWith(httpType, StringComparison.OrdinalIgnoreCase) ? new Uri(urlAddress) : new Uri(httpType + urlAddress);

                    // Start the stopwatch to calculate the download speed
                    sw.Start();

                    try
                    {

                        // Start downloading the file
                        Directory.CreateDirectory(tempPath + "\\" + fileName + "\\");
                        rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
                        rtMain.AppendText(DateTime.Now + " | " + "Downloading " + fileName + "...\r\n", Color.Black, FontStyle.Regular);
                        webClient.DownloadFileAsync(URL, tempPath + "\\" + fileName + "." + extension);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    Settings.Default.workerWait = false;
                }
            }

            //return pdfName;
        }

        // The event that will fire whenever the progress of the WebClient is changed
        public void ProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            // Update the progressbar percentage only when the value is not the same.
            pbDownload.Value = e.ProgressPercentage;

            
            /*if (Settings.Default.debugOn)
            {
                // Calculate download speed and output it to labelSpeed.
                lSpeed.Text = string.Format("{0} kb/s", (e.BytesReceived / 1024d / sw.Elapsed.TotalSeconds).ToString("0.00"));

                // Show the percentage on our label.
                lPerc.Text = e.ProgressPercentage.ToString() + "%";

                // Update the label with how much data have been downloaded so far and the total size of the file we are currently downloading
                lDownloaded.Text = string.Format("{0} MB's / {1} MB's",
                    (e.BytesReceived / 1024d / 1024d).ToString("0.00"),
                    (e.TotalBytesToReceive / 1024d / 1024d).ToString("0.00"));
            }*/




        }

        // The event that will trigger when the WebClient is completed
        public void Completed(object sender, AsyncCompletedEventArgs e)
        {
            // Reset the stopwatch.
            sw.Reset();
            if (e.Cancelled == true)
            {
                rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
                rtMain.AppendText(DateTime.Now + " | " + "Cancelling download...\r\n", Color.Black, FontStyle.Regular);
                rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
            }
            else
            {          
                downloadShaw.Process(downloadName, this, isSSGlobal);
                downloadShaw.CleanUp(downloadName);
                urlAddressList.RemoveAt(0);
                if (urlAddressList.Count > 0)
                {
                    string urlAddress = urlAddressList[0];
                    string fileName = urlAddress.Split('/').Last();
                    downloadName = fileName;
                    Uri URL = urlAddress.StartsWith(httpType, StringComparison.OrdinalIgnoreCase) ? new Uri(urlAddress) : new Uri(httpType + urlAddress);
                    string extension = "zip";
                    Directory.CreateDirectory(tempPath + "\\" + fileName + "\\");
                    rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
                    rtMain.AppendText(DateTime.Now + " | " + "Downloading " + fileName + "...\r\n", Color.Black, FontStyle.Regular);
                    webClient.DownloadFileAsync(URL, tempPath + "\\" + fileName + "." + extension);
                    //while (webClient.IsBusy) { };
                }
                else
                {
                    tMain.Start();
                    Settings.Default.workerWait = false;
                }                
            }

            
            
        }

        public ModelArmstrong.ArmstrongDB GetArmstrongDB()
        {
            ModelArmstrong.ArmstrongDB armstrongDB = new ModelArmstrong.ArmstrongDB();
            MethodsMySQL methodsMySQL = new MethodsMySQL();
            List<string>[] ahfDB = methodsMySQL.AhfDB();
            armstrongDB.FileName = ahfDB[0];
            armstrongDB.FileNameAlt = ahfDB[1];
            armstrongDB.PartNumber = ahfDB[2];
            armstrongDB.Stock = ahfDB[3];
            
            return armstrongDB;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] files = Directory.GetFiles("\\\\192.168.1.45\\Output1\\Shaw XML\\", "*.pdf");
            foreach (string f in files)
            {
                shawParse.buildShawSpreadSheet(f);
            }
            MessageBox.Show("Done");           
        }
    }

    public static class RichTextBoxExtensions
    {
        public static void AppendText(this RichTextBox box, string text, Color color, FontStyle style)
        {
            box.SelectionStart = box.TextLength;
            box.SelectionLength = 0;
            box.SelectionFont = new Font(box.Font, style);
            box.SelectionColor = color;
            box.AppendText(text);
            box.SelectionColor = box.ForeColor;
        }
    }
    public static class ModifyProgressBarColor
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr w, IntPtr l);
        public static void SetState(this ProgressBar pBar, int state)
        {
            SendMessage(pBar.Handle, 1040, (IntPtr)state, IntPtr.Zero);
        }
    }
}
