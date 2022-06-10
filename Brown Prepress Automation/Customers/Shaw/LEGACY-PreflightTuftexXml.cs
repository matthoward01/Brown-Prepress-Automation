using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using Brown_Prepress_Automation.Properties;
using PdfToImage;

namespace Brown_Prepress_Automation
{
    class PreflightTuftexXml
    {
        List<string> errorList = new List<string>();
        List<string> exceptionErrorList = new List<string>();
        int colorCount = 0;
        bool repeat = false;
        //string jpgName = "";
        string pdfFileName = "";
        string labelWoNumber = "";
        AddPageMethodsTuftex createPdfPageTuftex = new AddPageMethodsTuftex();
        //bool mailNotifications = true;
        MethodsCommon methods = new MethodsCommon();
        MethodsMail mail = new MethodsMail();

        public void TuftexRun(FormMain mainForm, string passedXmlList)
        {
            bool sendEmails = Settings.Default.sendEmails;
            try
            {
                string tempDir = Settings.Default.tempDir;                   
                    

                List<string> xmlFileList = new List<string>();
                xmlFileList.Add(passedXmlList);

                while (xmlFileList.Count > 0)
                {
                    Console.WriteLine("-------------------------------------------------------------");
                    string workingFile = xmlFileList[0];                    
                    bool errorCheck = false;

                    //Check for Label Sku
                    exceptionErrorList.Add(workingFile);
                    XElement xml = XElement.Load(Settings.Default.shawHotfolder + "\\" + workingFile);
                    var xmlCheck = xml
                    .Descendants("main-label")
                    .Select(labelSku =>
                    {
                        return new
                        {
                            sku = labelSku.Attribute("sku").Value,
                        };
                    });
                    foreach (var labelSku in xmlCheck)
                    {                        
                        if (!File.Exists(Settings.Default.tuftexPdf + "\\" + labelSku.sku + ".pdf"))
                        {
                            if (!File.Exists(Settings.Default.tuftexJpg + "\\" + labelSku.sku + ".jpg"))
                            {
                                errorCheck = true;
                                errorList.Add("Label Base " + labelSku.sku + " does not exist. Resubmit after it is created.");
                            }
                        }
                    }
                    if (errorCheck == true)
                    {
                        xmlFileList.Remove(workingFile);
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
                        System.IO.File.Move(Settings.Default.shawHotfolder + workingFile, Settings.Default.shawErrorFolder + "\\" + workingFile);
                        mail.SendMailTuftexTeam(workingFile, true);
                    }
                    else
                    {
                        //No Error Found with Label Sku, so try to create label
                        var xmlColorCount = XDocument.Load(Settings.Default.shawHotfolder + workingFile);
                        var colorVariables = from c in xmlColorCount.Root.Descendants("color")
                                             where (string)c.Attribute("label-required") == "yes"
                                             select c.Element("name").Value;
                        foreach (string colorVariable in colorVariables)
                        {
                            colorCount++;
                        }
                        if (colorCount % 2 != 0)
                        {
                            repeat = true;
                        }
                        var labels = xml
                        .Descendants("color")
                        .Where(el => el.Attribute("label-required").Value == "yes")
                        .Select(labelItems =>
                        {
                            XElement labelBase = labelItems.Parent.Parent;
                            return new
                            {
                                woNumber = labelBase.Parent.Parent.Attribute("id").Value,
                                styleNumber = labelBase.Element("style-id").Value,
                                styleWidth = labelBase.Element("width").Value,
                                stylePatternRepeat = labelBase.Element("pattern-repeat").Value,
                                styleName = labelBase.Element("name").Value,
                                //styleName = labelBase.Element("base-name").Value,
                                styleDurability = labelBase.Element("durability").Value,
                                styleFiber = labelBase.Element("fiber-content").Value,
                                styleBase = labelBase.Element("labels").Element("main-label").Attribute("sku").Value,
                                bugs = labelBase.Element("bugs").Elements("bug").Select(b => b.Attribute("sku").Value).ToList(),
                                colorSequence = labelItems.Attribute("sequence").Value,
                                colorNumber = labelItems.Element("color-id").Value,
                                colorName = labelItems.Element("name").Value,
                                colorRequired = labelItems.Attribute("label-required").Value
                                //Infos = labelBase.Elements().Where(b => b.Name.LocalName.StartsWith("info")).Select(b => b.Value).ToList()
                            };
                        });
                        //FileStream fs = new FileStream(xmlTuftexXmfOutput + Path.GetFileNameWithoutExtension(workingFile) + " - TUFTEX.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                        FileStream fs = new FileStream(tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - TUFTEX.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                        Document doc = new Document();
                        PdfWriter writer = PdfWriter.GetInstance(doc, fs);
                        writer.PdfVersion = PdfWriter.VERSION_1_3;
                        doc.SetPageSize(new iTextSharp.text.Rectangle(864, 612));
                        doc.SetMargins(0, 0, 0, 0);
                        doc.Open();
                        PdfContentByte cb = writer.DirectContent;
                        int progBar = labels.Count();
                        if (labels.Count() % 2 != 0)
                        {
                            progBar = progBar + 1;
                        }

                        int fileProgressStep = (int)Math.Ceiling(((double)100) / progBar);
                        foreach (var label in labels)
                        {
                            labelWoNumber = label.woNumber;
                            if (label.bugs.Count() == 0)
                            {
                                label.bugs.Add("");
                            }
                            if (pdfFileName == "")
                            {
                                //jpgName = label.styleNumber;
                                pdfFileName = label.styleNumber;
                            }
                            if (label.colorRequired == "yes")
                            {
                                if (repeat == true)
                                {
                                    createPdfPageTuftex.AddPageNormal(
                                        doc,
                                        cb,
                                        writer,
                                        label.woNumber,
                                        label.styleNumber,
                                        label.styleDurability,
                                        label.styleWidth,
                                        CultureInfo.CurrentCulture.TextInfo.ToTitleCase(label.stylePatternRepeat.ToLower()).Replace("X", "x").Replace("In", "in"),
                                        CultureInfo.CurrentCulture.TextInfo.ToTitleCase(label.styleName.ToLower()),
                                        label.styleFiber,
                                        Settings.Default.tuftexJpg + "\\" + label.styleBase + ".jpg",
                                        label.bugs.ToArray(),
                                        label.colorSequence,
                                        label.colorNumber,
                                        CultureInfo.CurrentCulture.TextInfo.ToTitleCase(label.colorName.ToLower()).Replace("Cafe", "Café"));
                                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                                    repeat = false;
                                }
                                createPdfPageTuftex.AddPageNormal(
                                       doc,
                                       cb,
                                       writer,  
                                       label.woNumber,
                                       label.styleNumber,
                                       label.styleDurability,
                                       label.styleWidth,
                                       CultureInfo.CurrentCulture.TextInfo.ToTitleCase(label.stylePatternRepeat.ToLower()).Replace("X", "x").Replace("In", "in"),
                                       CultureInfo.CurrentCulture.TextInfo.ToTitleCase(label.styleName.ToLower()),
                                       label.styleFiber,
                                       Settings.Default.tuftexJpg + "\\" + label.styleBase + ".jpg",
                                       label.bugs.ToArray(),
                                       label.colorSequence,
                                       label.colorNumber,
                                       CultureInfo.CurrentCulture.TextInfo.ToTitleCase(label.colorName.ToLower()).Replace("Cafe", "Café"));
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));                                       
                            }
                        }
                        doc.Close();

                        FileStream fs2 = new FileStream(tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - STORE.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                        Document doc2 = new Document();
                        PdfWriter writer2 = PdfWriter.GetInstance(doc2, fs2);
                        writer2.PdfVersion = PdfWriter.VERSION_1_3;
                        doc2.SetPageSize(new iTextSharp.text.Rectangle(864, 612));
                        doc2.SetMargins(0, 0, 0, 0);
                        doc2.Open();
                        PdfContentByte cb2 = writer2.DirectContent;
                        foreach (var label in labels)
                        {
                            labelWoNumber = label.woNumber;
                            if (label.bugs.Count() == 0)
                            {
                                label.bugs.Add("");
                            }
                            if (pdfFileName == "")
                            {
                                //jpgName = label.styleNumber;
                                pdfFileName = label.styleNumber;
                            }
                            if (label.colorRequired == "yes")
                            {
                                    createPdfPageTuftex.AddPageNormal(
                                        doc2,
                                        cb2,
                                        writer2,
                                        "create",
                                        label.styleNumber,
                                        label.styleDurability,
                                        label.styleWidth,
                                        CultureInfo.CurrentCulture.TextInfo.ToTitleCase(label.stylePatternRepeat.ToLower()).Replace("X", "x").Replace("In", "in"),
                                        CultureInfo.CurrentCulture.TextInfo.ToTitleCase(label.styleName.ToLower()),
                                        label.styleFiber,
                                        Settings.Default.tuftexJpg + "\\" + label.styleBase + ".jpg",
                                        label.bugs.ToArray(),
                                        label.colorSequence,
                                        label.colorNumber,
                                        CultureInfo.CurrentCulture.TextInfo.ToTitleCase(label.colorName.ToLower()).Replace("Cafe", "Café"));                                                              
                            }
                        }
                        doc2.Close();
                        //System.IO.File.Copy(xmlTuftexPDFs + Path.GetFileNameWithoutExtension(workingFile) + ".pdf", xmlTuftexXmfOutput + Path.GetFileNameWithoutExtension(workingFile) + " - TUFTEX.pdf");
                        errorList.Clear();
                        //Reset Color Count
                        colorCount = 0;
                        repeat = false;

                        //Create JPG
                        //method.jpgCreate(tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - TUFTEX.pdf", "\\\\192.168.1.45\\Tuftex\\Labels\\Test\\" + jpgName + ".jpg");
                        //jpgName = "";

                        //Create PDF
                        FileStream fs1 = new FileStream(Settings.Default.tuftexPdf + "\\" + pdfFileName + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                        Document doc1 = new Document();
                        PdfReader inputFile = new PdfReader(tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - STORE.pdf");
                        PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                        writer1.PdfVersion = PdfWriter.VERSION_1_3;
                        doc1.SetPageSize(new iTextSharp.text.Rectangle(864f, 612f));
                        doc1.SetMargins(0, 0, 0, 0);
                        doc1.Open();
                        doc1.NewPage();
                        var imp = writer1.GetImportedPage(inputFile, 1);
                        var tm = new System.Drawing.Drawing2D.Matrix();
                        PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                        tm.Translate(0f, 0f);
                        writer1.DirectContent.AddTemplate(imp, tm);
                        doc1.Close();
                        pdfFileName = "";


                        //Copy pdf to Hotfolder
                        if (!labelWoNumber.ToLower().Contains("create")){
                            System.IO.File.Copy(tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - TUFTEX.pdf", Settings.Default.tuftexXmlHotfolder + Path.GetFileNameWithoutExtension(workingFile) + " - TUFTEX.pdf", true);
                        }
                        labelWoNumber = "";
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine(DateTime.Now + "| " + workingFile + " Done");
                        Console.ResetColor();                        
                        xmlFileList.Remove(workingFile);
                    }
                    //methods.SendToPrinter(xmlTuftexXmfOutput + Path.GetFileNameWithoutExtension(workingFile) + " - TUFTEX.pdf");
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
                        errorFile.WriteLine(DateTime.Now + "| " + ex.Message + "\r\nAsk Matt");
                        Console.WriteLine(DateTime.Now + "| " + ex.Message + "\r\nAsk Matt");
                        Console.ResetColor();
                        if (sendEmails == true)
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
    }
}
