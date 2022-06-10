using System;
using System.Collections.Generic;
using System.Linq;
using ExcelLibrary.SpreadSheet;
using Brown_Prepress_Automation.Properties;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Drawing;

namespace Brown_Prepress_Automation
{
    class PreflightArmstrong
    {
        PdfProcessing pdfProcessing = new PdfProcessing();
        MethodsCommon methods = new MethodsCommon();
        MethodsTicket methodTicket = new MethodsTicket();
        MethodsMail mail = new MethodsMail();
        MethodsMySQL methodsMySQL = new MethodsMySQL();
        PreflightPdf preflightPdf = new PreflightPdf();
        string path = Settings.Default.tempDir;
        List<string> errorEmailList = new List<string>();        

        public void Preflight(FormMain mainForm, string passedFile, ModelArmstrong.ArmstrongDB armstrongDB)
        {
            ModelArmstrong.ArmstrongSheet armstrongSheet = new ModelArmstrong.ArmstrongSheet();
            ModelArmstrong.ArmstrongParse armstrongParse = new ModelArmstrong.ArmstrongParse();
            //List<string> flList = new List<string>();
            //List<string> blList = new List<string>();
            //List<int> labelQty = new List<int>();
            //List<int> flLines = new List<int>();
            //List<int> blLines = new List<int>();
            string customer = "armstrong";
            string workingFile = passedFile;
            Workbook book = Workbook.Load(Settings.Default.armstrongHotfolder + "\\" + workingFile);
            Worksheet sheet = book.Worksheets[0];
            string pdfs = Settings.Default.armstrongPdfs;
            if (workingFile.Contains("AHF Products"))
            {
                pdfs = Settings.Default.armstrongAHFpdfs;
                customer = "ahf products";
            }
            int validCellsCheck = methods.countValidCells(Settings.Default.armstrongHotfolder + "\\" + workingFile, 1, 0, 1);
            for (int i = 0; i < validCellsCheck + 1; i++)
            {
                armstrongSheet.PartNumber.Add(sheet.Cells[i, 0].StringValue.Trim());
                armstrongSheet.FileName.Add(sheet.Cells[i, 1].StringValue.Trim());
                armstrongSheet.Size.Add(sheet.Cells[i, 2].StringValue.Trim());
                armstrongSheet.Quantity.Add(sheet.Cells[i, 3].StringValue.Trim());
                armstrongSheet.Stock.Add(sheet.Cells[i, 4].StringValue.Trim());
                armstrongSheet.SalesOrder.Add(sheet.Cells[i, 5].StringValue.Trim());
            }
            if (customer == "ahf products")
            {
                for (int i = 1; i < validCellsCheck; i++)
                {
                    if (armstrongDB.FileName.Contains(armstrongSheet.FileName[i]))
                    {
                        if (armstrongSheet.PartNumber[i].Trim() == "")
                        {
                            armstrongSheet.PartNumber[i] = armstrongDB.PartNumber[armstrongDB.FileName.IndexOf(armstrongSheet.FileName[i])];
                        }
                        if (armstrongSheet.Size[i].Trim() == "")
                        {
                            armstrongSheet.Size[i] = "5 x 1.75";
                        }
                        if (armstrongSheet.Stock[i].Trim() == "")
                        {
                            armstrongSheet.Stock[i] = "PRINTS 4/C DIGITAL - INDIGO 6800 ON SG GUM PERMANENT + GLOSS FILM + DIE CUT WITH OMEGA DIE #284 ) REMOVE MATRIX + CARTON PACK BY VERSION";
                        }
                    }
                    else if (armstrongDB.FileNameAlt.Contains(armstrongSheet.FileName[i]))
                    {
                        if (armstrongSheet.PartNumber[i].Trim() == "")
                        {
                            armstrongSheet.PartNumber[i] = armstrongDB.PartNumber[armstrongDB.FileNameAlt.IndexOf(armstrongSheet.FileName[i])];
                        }
                        armstrongSheet.FileName[i] = armstrongDB.FileNameAlt[armstrongDB.FileNameAlt.IndexOf(armstrongSheet.FileName[i])];
                        if (armstrongSheet.Size[i].Trim() == "")
                        {
                            armstrongSheet.Size[i] = "5 x 1.75";
                        }
                        if (armstrongSheet.Stock[i].Trim() == "")
                        {
                            armstrongSheet.Stock[i] = "PRINTS 4/C DIGITAL - INDIGO 6800 ON SG GUM PERMANENT + GLOSS FILM + DIE CUT WITH OMEGA DIE #284 ) REMOVE MATRIX + CARTON PACK BY VERSION";
                        }
                    }                    
                    /*for (int z = 0; z < armstrongDB.FileName.Count(); z++)
                    {
                        if (armstrongSheet.FileName[i] == armstrongDB.FileName[z])
                        {
                            if (armstrongSheet.PartNumber[i].Trim() == "")
                            {
                                armstrongSheet.PartNumber[i] = armstrongDB.PartNumber[z];
                            }
                            if (armstrongSheet.Size[i].Trim() == "")
                            {
                                armstrongSheet.Size[i] = "5 x 1.75";
                            }
                            if (armstrongSheet.Stock[i].Trim() == "")
                            {
                                armstrongSheet.Stock[i] = "PRINTS 4/C DIGITAL - INDIGO 6800 ON SG GUM PERMANENT + GLOSS FILM + DIE CUT WITH OMEGA DIE #284 ) REMOVE MATRIX + CARTON PACK BY VERSION";
                            }
                        }
                        else if (armstrongSheet.FileName[i] == armstrongDB.FileNameAlt[z])
                        {
                            if (armstrongSheet.PartNumber[i].Trim() == "")
                            {
                                armstrongSheet.PartNumber[i] = armstrongDB.PartNumber[z];
                            }
                            armstrongSheet.FileName[i] = armstrongDB.FileName[z];
                            if (armstrongSheet.Size[i].Trim() == "")
                            {
                                armstrongSheet.Size[i] = "5 x 1.75";
                            }
                            if (armstrongSheet.Stock[i].Trim() == "")
                            {
                                armstrongSheet.Stock[i] = "PRINTS 4/C DIGITAL - INDIGO 6800 ON SG GUM PERMANENT + GLOSS FILM + DIE CUT WITH OMEGA DIE #284 ) REMOVE MATRIX + CARTON PACK BY VERSION";
                            }
                        }
                    }*/
                }
            }
            bool error = false;
            for (int i = 1; i < validCellsCheck; i++)
            {
                string armstrongPart = armstrongSheet.FileName[i];
                //string armstrongPart = sheet.Cells[i, 1].StringValue.Trim();
                string[] allFiles = Directory.GetFiles(pdfs, armstrongPart + ".pdf", SearchOption.AllDirectories);
                if (armstrongSheet.PartNumber[i].Trim() == "")
                {
                    error = true;
                    using (StreamWriter errorFile = new StreamWriter(Settings.Default.armstrongErrorFolder + "\\" + workingFile + ".txt", true))
                    {
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + " | " + armstrongPart + " is missing the Brown Part Number.", Color.Red, FontStyle.Regular); }));
                        errorFile.WriteLine(DateTime.Now + " | " + armstrongPart + " is missing the Brown Part Number.\r\n");
                        errorEmailList.Add(DateTime.Now + " | " + armstrongPart + "  is missing the Brown Part Number.");
                    }
                }
                if (allFiles.Count() > 1)
                {
                    error = true;
                    using (StreamWriter errorFile = new StreamWriter(Settings.Default.armstrongErrorFolder + "\\" + workingFile + ".txt", true))
                    {
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + " | " + armstrongPart + " is finding more than 1 entry.", Color.Red, FontStyle.Regular); }));
                        errorFile.WriteLine(DateTime.Now + " | " + armstrongPart + " is finding more than 1 entry. Get Matt to check.\r\n");
                        errorEmailList.Add(DateTime.Now + " | " + armstrongPart + " is finding more than 1 entry. Get Matt to check.");
                    }
                }
                else if (allFiles.Count() < 1)
                {
                    error = true;
                    using (StreamWriter errorFile = new StreamWriter(Settings.Default.armstrongErrorFolder + "\\" + workingFile + ".txt", true))
                    {
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + " | " + armstrongPart + " is missing.\r\n", Color.Red, FontStyle.Regular); }));
                        errorFile.WriteLine(DateTime.Now + " | " + armstrongPart + " is missing. Should you remove the FP from the beginning of the armstrong part name?\r\n");
                        errorEmailList.Add(DateTime.Now + " | " + armstrongPart + " is missing.");
                    }
                }
                else
                {
                    string partNumber = armstrongSheet.PartNumber[i];
                    int qty = int.Parse(armstrongSheet.Quantity[i]);
                    //string partNumber = sheet.Cells[i, 0].StringValue.Trim();
                    //int qty = int.Parse(sheet.Cells[i, 3].StringValue.Trim());

                    if (pdfProcessing.FormatGetSize(allFiles[0], "trim", 1) == "5.00 x 1.75")
                    {
                        armstrongParse.FlList.Add(allFiles[0]);
                        armstrongParse.LabelQty.Add(qty);
                        armstrongParse.FlLines.Add(i);
                    }
                    else if (pdfProcessing.FormatGetSize(allFiles[0], "trim", 1) == "16.00 x 16.00")
                    {
                        armstrongParse.BlList.Add(allFiles[0]);
                        armstrongParse.BlLines.Add(i);
                    }
                    else
                    {
                        throw new Exception("Size is not supported.");
                    }
                }
                /*
                if (!File.Exists(pdfs + "\\" + armstrongPart + ".pdf"))
                {
                   error = true;
                   using (System.IO.StreamWriter errorFile = new System.IO.StreamWriter(Settings.Default.armstrongErrorFolder + "\\" + workingFile + ".txt", true))
                   {
                       mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + armstrongPart + " is missing.\r\n", Color.Red, FontStyle.Regular); }));
                       errorFile.WriteLine(DateTime.Now + "| " + armstrongPart + " is missing. Should you remove the FP from the beginning of the armstrong part name?\r\n");
                       errorEmailList.Add(DateTime.Now + "| " + armstrongPart + " is missing.");
                   }
                }
                */
            }
            if (error)
            {
                if (Settings.Default.sendEmails == true)
                {
                    mail.SendMailArmstrongTeam(workingFile, errorEmailList, true);
                    errorEmailList.Clear();
                }
                if (File.Exists(Settings.Default.armstrongErrorFolder + "\\" + workingFile))
                {
                    File.Delete(Settings.Default.armstrongErrorFolder + "\\" + workingFile);
                }
                if (File.Exists(Settings.Default.armstrongHotfolder + "\\" + workingFile))
                {
                    File.Move(Settings.Default.armstrongHotfolder + "\\" + workingFile, Settings.Default.armstrongErrorFolder + "\\" + workingFile);
                }
            }
            if (!error)
            {                
                int fileProgressStep = (int)Math.Ceiling(((double)100) / (armstrongParse.BlList.Count + armstrongParse.FlList.Count));
                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                if (armstrongParse.BlList.Count != 0)
                {
                    string output1 = "\\\\192.168.1.45\\Output1\\ArmstrongandAHF\\" + armstrongSheet.PartNumber[1] + " - " + armstrongSheet.PartNumber[armstrongParse.BlLines.Count] + "\\";
                    //string output1 = "\\\\192.168.1.45\\Output1\\ArmstrongandAHF\\" + sheet.Cells[1, 0].StringValue.Trim() + " - " + sheet.Cells[armstrongParse.BlLines.Count, 0].StringValue.Trim() + "\\";
                    if (Directory.Exists(output1))
                    {
                        Directory.Delete(output1, true);
                    }
                    {
                        Directory.CreateDirectory(output1);
                    }
                    int blListCount = 0;
                    foreach (int l in armstrongParse.BlLines)  
                    {
                        string partNumber = armstrongSheet.PartNumber[l];
                        string armstrongPart = armstrongSheet.FileName[l];
                        int qty = int.Parse(armstrongSheet.Quantity[l]);
                        //string partNumber = sheet.Cells[l, 0].StringValue.Trim();
                        //string armstrongPart = sheet.Cells[l, 1].StringValue.Trim();
                        //int qty = int.Parse(sheet.Cells[l, 3].StringValue.Trim());

                        if (!Directory.Exists(output1 + "Qty " + qty.ToString() + "\\"))
                        {
                            Directory.CreateDirectory(output1 + "Qty " + qty.ToString() + "\\");
                        }
                        File.Copy(armstrongParse.BlList[blListCount], Settings.Default.XMF16x16 + "\\" + partNumber + ".pdf", true);
                        File.Copy(armstrongParse.BlList[blListCount], output1 + "Qty " + qty.ToString() + "\\" + partNumber + ".pdf", true);
                        preflightPdf.PreflightPdfPrint(mainForm, armstrongParse.BlList[blListCount]);
                        methodsMySQL.InsertPrepressLogAutomation(partNumber);
                        mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                        blListCount++;
                        /*
                        var c1 = partNumber.ToString();
                        var c2 = "";
                        var c3 = "";
                        var c4 = "X";
                        var newLine = string.Format("{0},{1},{2},{3}", c1, c2, c3, c4);
                        csv.AppendLine(newLine); 
                        */
                    }
                    //File.WriteAllText(filePath, csv.ToString());  
                    if (!Settings.Default.debugOn)
                    {
                        methodTicket.armstrongTicket(Settings.Default.armstrongHotfolder + "//" + workingFile, Path.GetFileNameWithoutExtension(workingFile), armstrongParse.BlLines, armstrongSheet);
                    }
                    if (Settings.Default.sendEmails == true)
                    {
                       //mail.sendMailArmstrongTeamCsv(workingFile, filePath);
                    }
                }
                if (armstrongParse.FlList.Count != 0)
                {
                    List<string> item = new List<string>();
                    List<int> itemQty = new List<int>();
                    List<string> itemPrint = new List<string>();
                    List<int> itemQtyPrint = new List<int>();
                    List<string> itemHold = new List<string>();
                    List<int> itemQtyHold = new List<int>();
                    List<string> itemTotal = new List<string>();
                    List<string> diffPerPage = new List<string>();
                    item = armstrongParse.FlList;
                    itemQty = armstrongParse.LabelQty;

                    int qtyCheck = 0;
                    while (item.Count > 0)
                    {
                        int itemListCount = 0;
                        itemHold.Clear();
                        itemQtyHold.Clear();
                        qtyCheck = itemQty[0];

                        foreach (int iqty in itemQty)
                        {
                            if (iqty != qtyCheck)
                            {
                                itemHold.Add(item[itemListCount]);
                                itemQtyHold.Add(iqty);
                            }
                            else
                            {
                                itemPrint.Add(item[itemListCount]);
                                itemQtyPrint.Add(iqty);
                            }
                            itemListCount++;
                        }
                        item = itemHold.ToList();
                        itemQty = itemQtyHold.ToList();

                        int printed = 0;

                        while (itemPrint.Count > 0)
                        {
                            if (itemPrint.Count() % 5 == 0)
                            {
                                itemTotal.Add(itemPrint[0]);
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                                itemTotal.Add(itemPrint[1]);
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                                itemTotal.Add(itemPrint[2]);
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                                itemTotal.Add(itemPrint[3]);
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                                itemTotal.Add(itemPrint[4]);
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                                itemPrint.RemoveRange(0, 5);
                                printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 7);
                                diffPerPage.Add("5 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                                itemQtyPrint.RemoveRange(0, 5);
                                
                            }
                            else
                            {
                                itemTotal.Add(itemPrint[0]);
                                itemTotal.Add(itemPrint[0]);
                                itemTotal.Add(itemPrint[0]);
                                itemTotal.Add(itemPrint[0]);
                                itemTotal.Add(itemPrint[0]);
                                itemPrint.RemoveAt(0);
                                printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 35);
                                diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                                itemQtyPrint.RemoveAt(0);
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                            }
                            
                        }
                    }
                    FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                    Document doc = new Document();
                    PdfWriter writer = PdfWriter.GetInstance(doc, fs);
                    writer.PdfVersion = PdfWriter.VERSION_1_3;
                    doc.Open();
                    doc.SetPageSize(new iTextSharp.text.Rectangle(408, 174));

                    while (itemTotal.Count() > 0)
                    {
                        doc.NewPage();
                        PdfReader readPDF = new PdfReader(itemTotal[0]);
                        PdfImportedPage readPDFPage = writer.GetImportedPage(readPDF, 1);
                        var trans = new System.Drawing.Drawing2D.Matrix();
                        writer.DirectContent.AddTemplate(readPDFPage, trans);
                        itemTotal.RemoveAt(0);
                    }
                    doc.Close();
                    File.Copy(path + "\\" + Path.GetFileNameWithoutExtension(workingFile) + ".pdf", Settings.Default.XMF5x1_75 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + ".pdf", true);
                    if (!Settings.Default.debugOn)
                    {
                        methodTicket.armstrongTicket6800(Settings.Default.armstrongHotfolder + "//" + workingFile, Path.GetFileNameWithoutExtension(workingFile), diffPerPage, armstrongParse.FlLines);
                    }
                }

                //Archive
                for (int i = 1; i < validCellsCheck; i++)
                {
                    string fileName = armstrongSheet.FileName[i];
                    string partNumber = armstrongSheet.PartNumber[i];
                    string qty = armstrongSheet.Quantity[i].ToString();
                    string specs = armstrongSheet.Stock[i];
                    string soNumber = armstrongSheet.SalesOrder[i];

                    //string partNumber = sheet.Cells[i, 0].StringValue;
                    //string qty = sheet.Cells[i, 3].StringValue;
                    //string specs = sheet.Cells[i, 4].StringValue;
                    //string soNumber = sheet.Cells[i, 5].StringValue;
                    string[] allFiles = Directory.GetFiles(pdfs, fileName.Trim() + ".pdf", SearchOption.AllDirectories);
                    string size = pdfProcessing.FormatGetSize(allFiles[0], "trim", 1);
                    methodsMySQL.InsertOrders(customer, Path.GetFileNameWithoutExtension(workingFile), fileName, partNumber, size, qty, "", soNumber, specs);
                }

                if (Settings.Default.sendEmails == true)
                {
                    mail.SendMailArmstrongTeam(workingFile, errorEmailList, false);
                }
                if (File.Exists(Settings.Default.armstrongArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + passedFile))
                {
                    File.Delete(Settings.Default.armstrongArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + passedFile);
                }
                Directory.CreateDirectory(Settings.Default.armstrongArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\");
                if (File.Exists(Settings.Default.armstrongHotfolder + "\\" + passedFile))
                {
                    File.Copy(Settings.Default.armstrongHotfolder + "\\" + passedFile, Settings.Default.armstrongArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + passedFile);
                }
                if (File.Exists(Settings.Default.armstrongHotfolder + "\\" + passedFile))
                {
                    File.Delete(Settings.Default.armstrongHotfolder + "\\" + passedFile);
                }
            }
        }        
    }
}
