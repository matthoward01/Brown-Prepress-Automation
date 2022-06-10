using System.Collections.Generic;
using System.Linq;
using Brown_Prepress_Automation.Properties;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Brown_Prepress_Automation
{
    class OutputHPShaw
    {
        MethodsCommon methods = new MethodsCommon();
        PdfProcessing pdfProcessing = new PdfProcessing();
        MethodsMail methodsMail = new MethodsMail();
        PreflightPdf preflightPdf = new PreflightPdf();

        public void HPShaw (List<string> files, List<string> partNumbers, FormMain mainForm)
        {
            string formattedSizeMedia = pdfProcessing.FormatGetSize(files[0], "media", 1);
            string formattedSizeTrim = pdfProcessing.FormatGetSize(files[0], "trim", 1);
            var sizes = formattedSizeTrim.Split('x');
            double width = double.Parse(sizes[0]);
            double height = double.Parse(sizes[1]);
            if (formattedSizeMedia == formattedSizeTrim)
            {
                width = width - 3;
                height = height - 3;
            }
            int partCount = 0;
            string size = width.ToString() + " x " + height.ToString();
            foreach (string f in files)
            {
                if (File.Exists(f))
                {
                    if (!Directory.Exists(Settings.Default.xmfHotfolders + "XMF " + size))
                    {
                        Directory.CreateDirectory(Settings.Default.xmfHotfolders + "XMF " + size);
                        methodsMail.SendMailCreateHotFolder("XMF " + size);
                    }
                    File.Copy(f, Settings.Default.xmfHotfolders + "XMF " + size + "\\" + partNumbers[partCount] + ".pdf", true);

                    preflightPdf.PreflightPdfPrint(mainForm, f);
                }
                partCount++;
            }
        }

        public string[] pdf1upBoard(string fileName, string cover, string liner, string frontLabel, string backLabel)
        {
            List<string> frontLabelSizeList = new List<string>();
            List<string> backLabelWidthList = new List<string>();
            List<string> backLabelHeightList = new List<string>();
            List<string> frontLabelList = new List<string>();
            string[] frontLabelArray = { frontLabel };
            frontLabelList = frontLabelArray.ToList();
            //frontLabelList = frontLabel.ToList();
            List<string> createdCoverList = new List<string>();
            List<string> createdLinerList = new List<string>();
            List<string> coverLinerList = new List<string>();
            coverLinerList.Clear();
            List<string> backLabelList = new List<string>();
            string[] backLabelArray = { backLabel };
            backLabelList = backLabelArray.ToList();
            //backLabelList = backLabel.ToList();
            int coverCount = 0;
            int linerCount = 0;
            string[] coverArray = { cover };
            string[] linerArray = { liner };

            foreach (string file in frontLabelArray)
            {
                if (Path.GetFileNameWithoutExtension(file) == "none")
                {
                    frontLabelSizeList.Add("none");
                }
                else
                {
                    FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                    Document doc1 = new Document();
                    PdfReader inputFile = new PdfReader(file);
                    PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                    writer1.PdfVersion = PdfWriter.VERSION_1_3;
                    //PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                    Rectangle fileSize = inputFile.GetBoxSize(1, "trim");
                    if (fileSize.Height < 250)
                    {
                        doc1.SetPageSize(new Rectangle(fileSize.Width, fileSize.Height));
                        doc1.SetMargins(0, 0, 0, 0);
                        doc1.Open();
                        doc1.NewPage();
                        var imp = writer1.GetImportedPage(inputFile, 1);
                        var tm = new System.Drawing.Drawing2D.Matrix();

                        tm.Translate(-24.12f, -24.12f);

                        writer1.DirectContent.AddTemplate(imp, tm);
                        frontLabelSizeList.Add("half");
                    }
                    else
                    {
                        doc1.SetPageSize(new Rectangle(fileSize.Width, fileSize.Height));
                        doc1.SetMargins(0, 0, 0, 0);
                        doc1.Open();
                        doc1.NewPage();
                        var imp = writer1.GetImportedPage(inputFile, 1);
                        var tm = new System.Drawing.Drawing2D.Matrix();

                        tm.Translate(-24.12f, -24.12f);

                        writer1.DirectContent.AddTemplate(imp, tm);
                        frontLabelSizeList.Add("full");
                    }
                    doc1.Close();
                }
            }

            foreach (string file in backLabelArray)
            {
                if (Path.GetFileNameWithoutExtension(file) == "none")
                {
                    backLabelWidthList.Add("none");
                    backLabelHeightList.Add("none");
                }
                else
                {
                    FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                    Document doc1 = new Document();
                    PdfReader inputFile = new PdfReader(file);
                    PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                    writer1.PdfVersion = PdfWriter.VERSION_1_3;
                    //PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                    Rectangle fileSize = inputFile.GetBoxSize(1, "trim");

                    doc1.SetPageSize(new Rectangle(fileSize.Width, fileSize.Height));
                    doc1.SetMargins(0, 0, 0, 0);
                    doc1.Open();
                    doc1.NewPage();

                    PdfReader whiteBGLabelFile = new PdfReader(FormMain.Globals.appDir + "Images\\White Background.pdf");
                    PdfImportedPage whiteBGLabelPage = writer1.GetImportedPage(whiteBGLabelFile, 1);
                    var whiteBGLabelPDF = new System.Drawing.Drawing2D.Matrix();
                    writer1.DirectContent.AddTemplate(whiteBGLabelPage, whiteBGLabelPDF);

                    var imp = writer1.GetImportedPage(inputFile, 1);
                    var tm = new System.Drawing.Drawing2D.Matrix();

                    tm.Translate(-24.12f, -24.12f);

                    writer1.DirectContent.AddTemplate(imp, tm);
                    backLabelWidthList.Add(fileSize.Width.ToString());
                    backLabelHeightList.Add(fileSize.Height.ToString());

                    doc1.Close();
                }
            }

            foreach (string file in coverArray)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + fileName + " Cover.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                Rectangle fileSize = inputFile.GetBoxSize(1, "trim");
                doc1.SetPageSize(new Rectangle(fileSize.Width, fileSize.Height));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                writer1.DirectContent.AddTemplate(imp, tm);

                //Front Label
                if (frontLabelSizeList[0] != "none")
                {
                    PdfReader frontLabelFile = new PdfReader(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(frontLabelList[0]) + ".pdf");
                    PdfImportedPage frontLabelPage = writer1.GetImportedPage(frontLabelFile, 1);
                    var frontLabelPDF = new System.Drawing.Drawing2D.Matrix();
                    if (frontLabelSizeList[0] == "half")
                    {
                        frontLabelPDF.Translate(1404.167f, 1602.167f);
                    }
                    if (frontLabelSizeList[0] == "full")
                    {
                        frontLabelPDF.Translate(1404.25f, 1403.5f);
                    }
                    writer1.DirectContent.AddTemplate(frontLabelPage, frontLabelPDF);
                }
                doc1.Close();
                frontLabelList.RemoveAt(0);
                frontLabelSizeList.RemoveAt(0);

                createdCoverList.Add(fileName + " Cover.pdf");
                coverCount++;
            }

            foreach (string file in linerArray)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + fileName + " Liner.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                Rectangle fileSize = inputFile.GetBoxSize(1, "trim");
                doc1.SetPageSize(new Rectangle(fileSize.Width, fileSize.Height));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                writer1.DirectContent.AddTemplate(imp, tm);

                //Back Label
                if (backLabelWidthList[0] != "none")
                {
                    float labelWidth = (fileSize.Width - float.Parse(backLabelWidthList[0])) / 2;
                    float labelHeight = ((fileSize.Height - float.Parse(backLabelHeightList[0])) / 2) - 63f;

                    PdfReader backLabelFile = new PdfReader(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(backLabelList[0]) + ".pdf");
                    PdfImportedPage backLabelPage = writer1.GetImportedPage(backLabelFile, 1);
                    var backLabelPDF = new System.Drawing.Drawing2D.Matrix();
                    backLabelPDF.Translate(labelWidth, labelHeight);
                    writer1.DirectContent.AddTemplate(backLabelPage, backLabelPDF);
                }
                doc1.Close();

                backLabelList.RemoveAt(0);
                backLabelWidthList.RemoveAt(0);
                backLabelHeightList.RemoveAt(0);

                createdLinerList.Add(fileName + " Liner.pdf");
                linerCount++;
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + fileName + " HP 40x56 - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new Rectangle(2880, 4032));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            while (createdCoverList.Count() > 0)
            {
                doc.NewPage();
                PdfReader coverFile = new PdfReader(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(createdCoverList[0]) + ".pdf");
                PdfImportedPage coverPage = writer.GetImportedPage(coverFile, 1);
                PdfReader linerFile = new PdfReader(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(createdLinerList[0]) + ".pdf");
                PdfImportedPage linerPage = writer.GetImportedPage(linerFile, 1);
                Rectangle coverFileSize = coverFile.GetBoxSize(1, "media");
                Rectangle linerFileSize = linerFile.GetBoxSize(1, "media");

                float coverXPlacement = ((doc.PageSize.Width - coverFileSize.Height) / 2) + coverFileSize.Height;
                float coverYPlacement = ((doc.PageSize.Height / 2) - coverFileSize.Width) / 2;
                float linerXPlacement = ((doc.PageSize.Width - coverFileSize.Height) / 2) + coverFileSize.Height;
                float linerYPlacement = (doc.PageSize.Height - linerFileSize.Width) - (((doc.PageSize.Height / 2) - linerFileSize.Width) / 2);

                //Cover
                var coverPDF = new System.Drawing.Drawing2D.Matrix();
                coverPDF.Translate(coverXPlacement, coverYPlacement);
                coverPDF.Rotate(90);
                writer.DirectContent.AddTemplate(coverPage, coverPDF);

                //Liner            
                var linerPDF = new System.Drawing.Drawing2D.Matrix();
                linerPDF.Translate(linerXPlacement, linerYPlacement);
                linerPDF.Rotate(90);
                writer.DirectContent.AddTemplate(linerPage, linerPDF);

                //Control
                pdfProcessing.PdfPlacement(writer, FormMain.Globals.appDir + "Images\\Shaw Control Bar.pdf", linerXPlacement - linerFileSize.Height, linerYPlacement, 90, 1);

                coverLinerList.Add(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(createdCoverList[0]) + ".pdf");
                coverLinerList.Add(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(createdLinerList[0]) + ".pdf");

                createdCoverList.RemoveAt(0);
                createdLinerList.RemoveAt(0);

            }
            doc.Close();

            coverCount = 0;
            linerCount = 0;

            return coverLinerList.ToArray();
        }
    }
}
