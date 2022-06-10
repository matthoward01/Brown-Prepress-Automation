using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading;
using ExcelLibrary.SpreadSheet;
using Brown_Prepress_Automation.Properties;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Net;

namespace Brown_Prepress_Automation
{
    class Preflight
    {
        MethodsCommon commonMethods = new MethodsCommon();
        PreflightShaw preflightShaw = new PreflightShaw();
        PreflightTuftexXml preflightTuftexXml = new PreflightTuftexXml();
        PreflightTuftexMiscOld preflightTuftexMiscOld = new PreflightTuftexMiscOld();
        PreflightNourison preflightNourison = new PreflightNourison();
        PreflightPdf preflightPdf = new PreflightPdf();
        PdfProcessing pdfProcessing = new PdfProcessing();
        DownloadShaw downloadShaw = new DownloadShaw();

        public void PreflightPdf(FormMain mainForm, string passedFile)
        {
            preflightPdf.PreflightPdfPrint(mainForm, Settings.Default.hotFolder + "\\" + passedFile);
        }

        public void Download(FormMain mainForm, string passedFile)
        {
            List<string> downloadList = new List<string>();
            string prevUrl = "";
            string fileName = "";
            bool isSS = false;
            File.Copy(Settings.Default.shawHotfolder + "download\\" + passedFile, Settings.Default.shawHotfolder + "\\" + passedFile, true);
            string workingFile = Settings.Default.shawHotfolder + "download\\" + passedFile;
            Workbook book = Workbook.Load(workingFile);
            Worksheet sheet = book.Worksheets[0];
            int validCellsCheck = commonMethods.countValidCells(workingFile, 1, 0, 0);
            for (int i = 1; i < validCellsCheck; i++)
            {
                if (sheet.Cells[i, 0].StringValue.Contains("http"))
                {
                    if (prevUrl != sheet.Cells[i, 0].StringValue)
                    {
                        string httpType = "https";
                        string urlAddress = sheet.Cells[i, 0].StringValue.ToString();
                        Uri URL = urlAddress.StartsWith(httpType, StringComparison.OrdinalIgnoreCase) ? new Uri(urlAddress) : new Uri(httpType + urlAddress);
                        try
                        {
                            HttpWebRequest request = (HttpWebRequest)System.Net.WebRequest.Create(URL);
                            request.Method = "HEAD";
                            request.KeepAlive = false;
                            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                            {
                                //HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                                string disposition = response.Headers["Content-Disposition"];
                                fileName = disposition.Substring(disposition.IndexOf("filename=") + 10).Replace("\"", "");
                                Brown_Prepress_Automation.FormMain.Globals.shawCheckList.Add(Path.GetFileName(fileName));
                                //response.Close();
                            }
                        }
                        catch (WebException e)
                        {
                            throw e;
                        }
                        //response.Dispose();
                        //fileName = sheet.Cells[i, 0].StringValue.ToString().Split('/').Last();
                        fileName = Path.GetFileNameWithoutExtension(fileName);
                        downloadList.Add(sheet.Cells[i, 0].StringValue);
                        prevUrl = sheet.Cells[i, 0].StringValue;
                    }
                    
                    if(sheet.Cells[i, 8].StringValue.ToLower().Contains("book"))
                    {
                        commonMethods.XlsWrite(i, fileName.Trim() + "LookBook", sheet.Cells[i, 1].StringValue, Settings.Default.shawHotfolder + "\\" + passedFile);
                    }
                    else if (sheet.Cells[i, 8].StringValue.ToLower().Contains("photopack") || sheet.Cells[i, 8].StringValue.ToLower().Contains("photo"))
                    {
                        commonMethods.XlsWrite(i, fileName.Trim() + "PhotoPack", sheet.Cells[i, 1].StringValue, Settings.Default.shawHotfolder + "\\" + passedFile);
                    }
                    else if ((sheet.Cells[i, 8].StringValue.ToLower().Contains("s/s")))
                    {
                        commonMethods.XlsWrite(i, fileName.Trim() + "-SilkScreen", sheet.Cells[i, 1].StringValue, Settings.Default.shawHotfolder + "\\" + passedFile);
                        isSS = true;
                    }
                    else
                    {
                        commonMethods.XlsWrite(i, fileName.Trim(), sheet.Cells[i, 1].StringValue, Settings.Default.shawHotfolder + "\\" + passedFile);
                    }
                }
                else
                {
                    Settings.Default.workerWait = false;
                    if (File.Exists(Settings.Default.shawHotfolder + "download\\" + passedFile))
                    {
                        File.Delete(Settings.Default.shawHotfolder + "download\\" + passedFile);
                    }
                }
                //Thread.Sleep(2000);
            }            
            mainForm.DownloadFile(downloadList, "zip", isSS);
        }

        public void PreflightXmlDownload(FormMain mainForm, string passedFile)
        {
            string workingFile = Settings.Default.hotFolder + "\\" + passedFile;
            Workbook book = Workbook.Load(workingFile);
            Worksheet sheet = book.Worksheets[0];
            int validCellsCheck = commonMethods.countValidCells(workingFile, 1, 0, 0);
            for (int i = 1; i < validCellsCheck; i++)
            {
                if (passedFile.ToLower().Contains("tuftex"))
                {
                    commonMethods.xmlDownloadTuftex(sheet.Cells[i, 0].StringValue.Trim());
                }
            }            
        }        

        /* LEGACY
        public void PreflightShaw(FormMain mainForm, string passedFile)
        {
            if (passedFile.ToLower().Contains("board"))
            {
                preflightShaw.PreflightShawBoardXLS(mainForm, passedFile);
            }
            else if (passedFile.ToLower().Contains("xml"))
            {
                preflightTuftexXml.TuftexRun(mainForm, passedFile);
            }
            else if (passedFile.ToLower().Contains("misc") || passedFile.ToLower().Contains("mill") || passedFile.ToLower().Contains("ddp"))
            {
                preflightTuftexMiscOld.TuftexMisOldRun(mainForm, passedFile);
            }
            else if (passedFile.ToLower().Contains("get"))
            {
                PreflightXmlDownload(mainForm, passedFile);
            }
            else
            {
                preflightShaw.PreflightShawLabelsXLS(mainForm, passedFile);
            }
        }

        public void PreflightTuftex(FormMain mainForm, string passedFile)
        {
            if (passedFile.ToLower().Contains("xml"))
            {
                preflightTuftexXml.TuftexRun(mainForm, passedFile);
            }
            else if (passedFile.ToLower().Contains("misc") || passedFile.ToLower().Contains("mill") || passedFile.ToLower().Contains("ddp"))
            {
                preflightTuftexMiscOld.TuftexMisOldRun(mainForm, passedFile);
            }
            else if (passedFile.ToLower().Contains("get"))
            {
                PreflightXmlDownload(mainForm, Settings.Default.hotFolder + "\\" + passedFile);
            }
            else                    
            {
                throw new Exception("Unsupported Label Type");
            }                      
        }

        public void PreflightNourison(FormMain mainForm, string passedFile)
        {
            preflightNourison.PreflightNourisonPop(mainForm, passedFile);
        }*/
    }
}
