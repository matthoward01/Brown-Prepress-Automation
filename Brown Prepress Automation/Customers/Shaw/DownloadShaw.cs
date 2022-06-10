using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Brown_Prepress_Automation.Properties;
using System.Drawing;
using System.IO.Compression;

namespace Brown_Prepress_Automation
{
    class DownloadShaw
    {
        PdfProcessing pdfProcessing = new PdfProcessing();
        MethodsCommon methodsCommon = new MethodsCommon();
        Stopwatch sw = new Stopwatch();    // Stopwatch to calculate the download speed
        string tempPath = Settings.Default.tempDir + "Shaw";    

        private static void Unzip(string sourceFile, string destination, FormMain mainForm)
        {
            Shell32.Shell sc = new Shell32.Shell();
            Shell32.Folder SrcFlder = sc.NameSpace(sourceFile);
            Shell32.Folder DestFlder = sc.NameSpace(destination);
            Shell32.FolderItems items = SrcFlder.Items();
            DestFlder.CopyHere(items, 20);
        }

        private static void UnzipNew(string sourceFile, string destination, FormMain mainForm)
        {
            List<int> linkRemoveList = new List<int>();
            mainForm.pbIndividual.Value = 0;
            using (ZipArchive archive = ZipFile.OpenRead(sourceFile))
            {
                int zipCount = 0;
                for (int i = 0; i < archive.Entries.Count; i++)
                {
                    if (archive.Entries[i].FullName.ToLower().Contains("links"))
                    {
                        linkRemoveList.Add(i);
                    }
                }
                for (int i = 0; i < archive.Entries.Count; i++)
                //foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if ((archive.Entries[i].FullName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) && (!linkRemoveList.Contains(i)))
                    {
                        zipCount++;
                    }
                }
                for (int i = 0; i < archive.Entries.Count; i++)
                //foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    int fileProgressStep = (int)Math.Ceiling(((double)100) / zipCount);
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                    if ((archive.Entries[i].FullName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) && (!linkRemoveList.Contains(i)))
                    {
                        string destinationPath = "";
                        string filename = archive.Entries[i].FullName.Substring(0, archive.Entries[i].FullName.IndexOf("/"));
                        filename = filename.Replace("-", " ");
                        if (archive.Entries[i].FullName.Contains("book"))
                        {
                            destinationPath = Path.GetFullPath(Path.Combine(destination, filename + "LookBook.pdf"));
                        }
                        else if (archive.Entries[i].FullName.Contains("pack"))
                        {
                            destinationPath = Path.GetFullPath(Path.Combine(destination, filename + "PhotoPack.pdf"));
                        }
                        else
                        {
                            destinationPath = Path.GetFullPath(Path.Combine(destination, Path.GetFileName(archive.Entries[i].FullName)));
                        }  
                        if (destinationPath.StartsWith(destination, StringComparison.Ordinal))
                        {
                            archive.Entries[i].ExtractToFile(destinationPath, true);
                        }
                        mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                    }
                }
            }
        }

        public void Process(string fileName, FormMain mainForm, bool isSS)
        {
            try
            {
                mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + " | Extracting " + fileName + "...\r\n", Color.Black, FontStyle.Regular); }));
                UnzipNew(tempPath + "\\" + fileName + ".zip", tempPath + "\\" + fileName + "\\", mainForm);
                //Unzip(tempPath + "\\" + fileName + ".zip", tempPath + "\\" + fileName + "\\", mainForm);
                mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular); }));
                List<string> files = new List<string>(Directory.GetFiles(tempPath + "\\" + fileName + "\\", "*.pdf", SearchOption.AllDirectories));
                Directory.CreateDirectory(tempPath + "\\" + fileName + "\\pdfs\\");
                foreach (string f in files)
                {
                    if ((!f.Contains("._")) && (!Path.GetFileName(f).StartsWith("_")))
                    {
                        File.Copy(f, tempPath + "\\" + fileName + "\\pdfs\\" + Path.GetFileName(f), true);
                        File.Decrypt(tempPath + "\\" + fileName + "\\pdfs\\" + Path.GetFileName(f));
                        CopyPdf(tempPath + "\\" + fileName + "\\pdfs\\" + Path.GetFileName(f), isSS);                        
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void CopyPdf(string fileName, bool isSS)
        {
            string formatSize = pdfProcessing.FormatGetSize(fileName, "trim", 1);
            if (!Directory.Exists(tempPath + "\\" + formatSize + "\\"))
            {
                Directory.CreateDirectory(tempPath + "\\" + formatSize + "\\");                
            }
            if (isSS)
            {
                File.Copy(fileName, tempPath + "\\" + formatSize + "\\" + Path.GetFileNameWithoutExtension(fileName) + " SilkScreen" + Path.GetExtension(fileName), true);
            }
            else
            {
                File.Copy(fileName, tempPath + "\\" + formatSize + "\\" + Path.GetFileName(fileName), true);
            }
            var sizes = formatSize.Split('x');
            double area = double.Parse(sizes[0].Trim()) * double.Parse(sizes[1].Trim());
            //if ((area <= 216) && (!fileName.ToLower().Contains("lookbook")) && (!fileName.ToLower().Contains("photopack")))
            if ((area <= 216) && (!fileName.ToLower().Contains("lookbook")) && (!fileName.ToLower().Contains("photopack")) && (formatSize != "7.00 x 12.00") && (formatSize != "6.00 x 12.00") && (formatSize != "6.00 x 11.00"))
            {
                pdfProcessing.SplitPdf(tempPath + "\\" + formatSize + "\\" + Path.GetFileName(fileName), tempPath + "\\" + formatSize + "\\");
                pdfProcessing.SplitPdfCleanup(tempPath + "\\" + formatSize + "\\" + Path.GetFileName(fileName));
            }
            else if ((area <= 216) && (!fileName.ToLower().Contains("lookbook")) && (!fileName.ToLower().Contains("photopack")) && (formatSize == "7.00 x 12.00") && (formatSize == "6.00 x 12.00") && (formatSize == "6.00 x 11.00"))
            {
                pdfProcessing.SplitPdf2Set(tempPath + "\\" + formatSize + "\\" + Path.GetFileName(fileName), tempPath + "\\" + formatSize + "\\");
                pdfProcessing.SplitPdfCleanup(tempPath + "\\" + formatSize + "\\" + Path.GetFileName(fileName));
            }
            /*else if ((area > 216) && (pdfProcessing.GetPdfTotalPages(fileName) > 1) && (!isSS) && (formatSize == pdfProcessing.FormatGetSize(fileName, "trim", 2)))
            {
                pdfProcessing.SplitPdf(tempPath + "\\" + formatSize + "\\" + Path.GetFileName(fileName), tempPath + "\\" + formatSize + "\\");
                pdfProcessing.SplitPdfCleanup(tempPath + "\\" + formatSize + "\\" + Path.GetFileName(fileName));
            }*/
            else if ((area > 216) && (pdfProcessing.GetPdfTotalPages(fileName) > 2) && (!isSS) && (formatSize != pdfProcessing.FormatGetSize(fileName, "trim", 2))) 
            {
                pdfProcessing.SplitPdf2Set(tempPath + "\\" + formatSize + "\\" + Path.GetFileName(fileName), tempPath + "\\" + formatSize + "\\");
                pdfProcessing.SplitPdfCleanup(tempPath + "\\" + formatSize + "\\" + Path.GetFileName(fileName));
            }
            Settings.Default.downloadPdfName = fileName;
        }
        
        public void CleanUp(string fileName)
        {
            try
            {
                bool go = false;
                while (!go)
                {
                    FileInfo dn = new FileInfo(tempPath + "\\" + fileName + ".zip");
                    if (!methodsCommon.IsFileLocked(dn))
                    {
                        Directory.Delete(tempPath + "\\" + fileName, true);
                        File.Delete(tempPath + "\\" + fileName + ".zip");
                        go = true;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }    
    }
}
