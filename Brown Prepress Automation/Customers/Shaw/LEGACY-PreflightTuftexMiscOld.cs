using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Globalization;
using ExcelLibrary.SpreadSheet;
using Brown_Prepress_Automation.Properties;

namespace Brown_Prepress_Automation
{
    class PreflightTuftexMiscOld
    {
        AddPageMethodsMiscOld createPdfPageTuftexMiscOld = new AddPageMethodsMiscOld();
        MethodsTicket ticket = new MethodsTicket();
        MethodsCommon methods = new MethodsCommon();
        MethodsMail mail = new MethodsMail();
        //bool mailNotifications = true;
        int countPhotoArrayMethods = 0;
        int countImagesPerPdfPageMethods = 0;
        string[] styleCreateLayout = { "", "" };
        string[] styleNameCreateLayout = { "", "" };
        string[] colorCreateLayout = { "", "" };
        string[] seqCreateLayout = { "", "" };
        string[] woCreateLayout = { "", "" };
        string[] dateCreateLayout = { "", "" };
        string[] labelCreateLayout = { "", "" };
        //string xmlTuftexHotfolder = "";        
        List<string> errorList = new List<string>();
        List<string> exceptionErrorList = new List<string>();

        public void TuftexMisOldRun(FormMain mainForm, string checkList1)
        {
            try
            {
                List<string> checkList = new List<string>();
                checkList.Add(checkList1);
                while (checkList.Count > 0)
                {
                    Console.WriteLine("-------------------------------------------------------------");
                    string workingFile = checkList[0];
                    exceptionErrorList.Add(workingFile);
                    bool errorCheck = false;
                    string wonumberCheck = "";
                    string styleNumberCheck = "";
                    Workbook bookCheck = Workbook.Load(Settings.Default.shawHotfolder + "\\" + workingFile);
                    Worksheet sheetCheck = bookCheck.Worksheets[0];
                    int validCellsCheck = methods.countValidCells(Settings.Default.shawHotfolder + "\\" + workingFile, 6, 0, 0);
                    for (int i = 6; i < validCellsCheck; i++)
                    {
                        wonumberCheck = sheetCheck.Cells[i, 1].StringValue;
                        styleNumberCheck = sheetCheck.Cells[i, 2].StringValue;
                        string checkFile = Settings.Default.tuftexPdf + "\\" + styleNumberCheck.Trim() + ".pdf";
                        string checkFile2 = Settings.Default.tuftexJpg + "\\" + styleNumberCheck.Trim() + ".jpg";
                        if (File.Exists(checkFile) == false)
                        {                           
                           
                            if (File.Exists(checkFile2) == false)
                            {                                
                                errorList.Add("The image file for the style " + styleNumberCheck + " is Missing for " + wonumberCheck + ".\r\n");
                                errorCheck = true;
                            }
                            if (errorCheck == false)
                            {
                                Console.ForegroundColor = ConsoleColor.Yellow;
                                Console.WriteLine("PDF File does not exist for " + styleNumberCheck + " yet...using JPG instead for " + wonumberCheck + ".\r\n");
                                Console.ResetColor();
                            }
                        }
                        if ((File.Exists(checkFile2) == false) && (File.Exists(checkFile) == true))
                        {
                            System.IO.Directory.CreateDirectory(Settings.Default.tempDir + "\\jpgs\\");
                            methods.jpgCreate(Settings.Default.tuftexPdf + "\\" + styleNumberCheck + ".pdf", Settings.Default.tuftexJpg + "\\" + styleNumberCheck + ".jpg", 100, 600, 600, 1, 1);
                        }
                        
                    }
                    if (errorCheck == true)
                    {
                        checkList.Remove(workingFile);
                        using (System.IO.StreamWriter errorFile = new System.IO.StreamWriter(Settings.Default.shawErrorFolder + "\\" + workingFile + ".txt", true))
                            foreach (string error in errorList)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                errorFile.WriteLine(DateTime.Now + "| " + error);
                                Console.WriteLine(DateTime.Now + "| " + error);
                                Console.ResetColor();
                            }
                        if (File.Exists(Settings.Default.shawErrorFolder + "\\" + workingFile))
                        {
                            System.IO.File.Delete(Settings.Default.shawErrorFolder + "\\" + workingFile);
                        }
                        System.IO.File.Move(Settings.Default.shawHotfolder + "\\" + workingFile, Settings.Default.shawErrorFolder + "\\" + workingFile);
                        mail.SendMailTuftexTeam(workingFile, true);
                    }
                    else
                    {
                        string wonumberRun = "";
                        string styleNameRun = "";
                        string styleNumberRun = "";
                        string colorRun = "";
                        string sequenceNumberRun = "";
                        string styleLabelRun = "";
                        List<string> qtyCheck = new List<string>();
                        string dateRun = DateTime.Now.ToString("MM/dd/yy");
                        int qtyRun = 0;
                        Workbook bookRun = Workbook.Load(Settings.Default.shawHotfolder + "\\" + workingFile);
                        Worksheet sheetRun = bookRun.Worksheets[0];
                        string typeRun = "normal";
                        int validCellsRun = methods.countValidCells(Settings.Default.shawHotfolder + "\\" + workingFile, 6, 0, 0);
                        int ddpvalidcells = 0;
                        List<int> diffqtyCheck = new List<int>();
                        for (int i = 6; i < validCellsRun; i++)
                        {
                            diffqtyCheck.Add(Convert.ToInt32(sheetRun.Cells[i, 9].StringValue));
                        }
                        if (workingFile.ToLower().Contains("misc") || workingFile.ToLower().Contains("mill"))
                        {
                            typeRun = "misc";
                        }
                        if (workingFile.Contains("DDP"))
                        {
                            typeRun = "ddp";
                        }
                        if (workingFile.Contains("PVD01"))
                        {
                            typeRun = "PVD01";
                        }
                        if (workingFile.Contains("SEQ"))
                        {
                            typeRun = "SEQ";
                        }
                        //if ((workingFile.Contains("Misc") || workingFile.Contains("Mill")) && workingFile.Contains("QTY"))
                        //{
                        //    typeRun = "miscqty";
                        //}
                        if ((diffqtyCheck.Any(o => o != diffqtyCheck[0])) && (workingFile.ToLower().Contains("misc") || workingFile.ToLower().Contains("mill")))
                        {
                            typeRun = "miscqty";
                        }
                        FileStream fs = new FileStream(Settings.Default.tuftexMiscHotfolder + Path.GetFileNameWithoutExtension(workingFile) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                        Document doc = new Document();
                        PdfWriter writer = PdfWriter.GetInstance(doc, fs);
                        writer.PdfVersion = PdfWriter.VERSION_1_3;
                        doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
                        doc.SetMargins(0, 0, 0, 0);
                        doc.Open();
                        PdfContentByte cb = writer.DirectContent;
                        int progBar = validCellsRun;
                        if (validCellsRun % 2 != 0)
                        {
                            progBar = progBar + 1;
                        }
                        
                        int fileProgressStep = (int)Math.Ceiling(((double)100) / progBar);
                        for (int i = 6; i < validCellsRun; i++)
                        {
                            qtyRun = Convert.ToInt32(sheetRun.Cells[i, 9].StringValue);
                            wonumberRun = sheetRun.Cells[i, 1].StringValue;
                            styleNumberRun = sheetRun.Cells[i, 2].StringValue.Trim();
                            styleLabelRun = sheetRun.Cells[i, 7].StringValue.Trim();
                            //styleNameRun = sheetRun.Cells[i, 3].StringValue;
                            styleNameRun = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(sheetRun.Cells[i, 3].StringValue.ToLower());
                            if ((sheetRun.Cells[i, 5].StringValue == "") ||
                                (sheetRun.Cells[i, 5].StringValue == " "))
                            {
                                colorRun = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(sheetRun.Cells[i, 4].StringValue.ToLower());
                            }
                            else
                            {
                                colorRun = sheetRun.Cells[i, 5].StringValue.PadLeft(5, '0') + " " + CultureInfo.CurrentCulture.TextInfo.ToTitleCase(sheetRun.Cells[i, 4].StringValue.ToLower());
                            }
                            sequenceNumberRun = "DISPLAY SEQUENCE: " + sheetRun.Cells[i, 10].StringValue;
                            int qtyCounter = qtyRun;
                            if ((typeRun == "ddp") || (typeRun == "miscqty"))
                            {
                                while (qtyCounter > 0)
                                {
                                    createLayout(doc, cb, writer, validCellsRun, styleNumberRun, styleNameRun, colorRun, sequenceNumberRun, wonumberRun, dateRun, typeRun, styleLabelRun);
                                    qtyCounter--;
                                    ddpvalidcells++;
                                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                                }
                            }
                            else
                            {
                                createLayout(doc, cb, writer, validCellsRun, styleNumberRun, styleNameRun, colorRun, sequenceNumberRun, wonumberRun, dateRun, typeRun, styleLabelRun);
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                            }
                        }
                        if ((typeRun == "ddp") || (typeRun == "miscqty"))
                        {
                            if (ddpvalidcells % 2 != 0)
                            {
                                createLayout(doc, cb, writer, validCellsRun, styleNumberRun, styleNameRun, colorRun, sequenceNumberRun, wonumberRun, dateRun, typeRun, styleLabelRun);
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                            }
                        }
                        else
                        {
                            if (validCellsRun % 2 != 0)
                            {                            
                                createLayout(doc, cb, writer, validCellsRun, styleNumberRun, styleNameRun, colorRun, sequenceNumberRun, wonumberRun, dateRun, typeRun, styleLabelRun);
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                            }
                        }
                        if (typeRun == "ddp")
                        {
                            ticket.ddpTicket(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile));
                        }
                        //Close Document
                        doc.Close();
                        errorList.Clear();


                        //Reset Color Count
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine(DateTime.Now + "| " + workingFile + " Done");
                        Console.ResetColor();                        
                        checkList.Remove(workingFile);
                    }
                    Console.WriteLine("-------------------------------------------------------------");
                    exceptionErrorList.Remove(workingFile);
                }
            }
            catch (Exception ex)
            {
                foreach (string e in exceptionErrorList)
                {
                    using (System.IO.StreamWriter errorFile = new System.IO.StreamWriter(Settings.Default.shawErrorFolder + "\\" + e + ".txt", true))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        errorFile.WriteLine(DateTime.Now + "| " + ex.Message + "\r\nCheck your spreadsheet format. \r\nIf this is a misc or DDp tuftex order, it must contain \"Mill\" or \"DDP\" in the file name. Ask Matt with questions.");
                        Console.WriteLine(DateTime.Now + "| " + ex.Message + "\r\nCheck your spreadsheet format. \r\nIf this is a misc or DDP tuftex order, it must contain \"Mill\" or \"DDP\" in the file name. Ask Matt with questions.");
                        Console.ResetColor();
                        if (Settings.Default.sendEmails == true)
                        {
                            mail.SendMailTuftexTeam(e, true);
                        }
                    }
                    if (File.Exists(Settings.Default.shawErrorFolder + "\\" + e))
                    {
                        System.IO.File.Delete(Settings.Default.shawErrorFolder + "\\" + e);
                    }
                    System.IO.File.Move(Settings.Default.shawHotfolder + "\\" + e, Settings.Default.shawErrorFolder + "\\" + e);                    
                }
            }
            finally
            {                
                exceptionErrorList.Clear();
                errorList.Clear();
            }
        }

        public void createLayout(Document doc, PdfContentByte cb, PdfWriter writer, int validCells, string styleLayoutMethod, string styleNameLayoutMethod, string colorLayoutMethod, string seqLayoutMethod, string woNumberLayoutMethod, string dateLayoutMethod, string typeLayoutMethod, string labelLayoutMethod)
        {
            styleCreateLayout[countPhotoArrayMethods] = (styleLayoutMethod.Trim());
            styleNameCreateLayout[countPhotoArrayMethods] = (styleNameLayoutMethod.Trim());
            colorCreateLayout[countPhotoArrayMethods] = (colorLayoutMethod.Trim());
            seqCreateLayout[countPhotoArrayMethods] = (seqLayoutMethod.Trim());
            woCreateLayout[countPhotoArrayMethods] = (woNumberLayoutMethod.Trim());
            dateCreateLayout[countPhotoArrayMethods] = dateLayoutMethod;
            labelCreateLayout[countPhotoArrayMethods] = (labelLayoutMethod.Trim());
            countPhotoArrayMethods++;
            countImagesPerPdfPageMethods++;
            if (countImagesPerPdfPageMethods == 2)
            {
                createPdfPageTuftexMiscOld.AddPageNormal(doc, cb,
                    writer,
                    Settings.Default.tuftexJpg + "\\" + styleCreateLayout[0] + ".jpg", Settings.Default.tuftexJpg + "\\" + styleCreateLayout[1] + ".jpg",
                    woCreateLayout[0], woCreateLayout[1],
                    styleNameCreateLayout[0], styleNameCreateLayout[1],
                    colorCreateLayout[0], colorCreateLayout[1],
                    dateCreateLayout[0], dateCreateLayout[1],
                    seqCreateLayout[0], seqCreateLayout[1],
                    labelCreateLayout[0], labelCreateLayout[1],
                    typeLayoutMethod);
                countPhotoArrayMethods = 0;
                countImagesPerPdfPageMethods = 0;
                for (int z = 0; z <= 1; z++)
                {
                    styleCreateLayout[z] = "Blank";
                    woCreateLayout[z] = "";
                    styleNameCreateLayout[z] = "";
                    colorCreateLayout[z] = "";
                    seqCreateLayout[z] = "";
                    dateCreateLayout[z] = "";
                    labelCreateLayout[z] = "";
                }
            }
        }
    }
}

