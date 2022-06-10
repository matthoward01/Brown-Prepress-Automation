using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text.pdf;
using System.IO;
using iTextSharp.text;
using Brown_Prepress_Automation.Properties;
using System.Drawing;

namespace Brown_Prepress_Automation
{
    class PreflightPdf
    {
        MethodsCommon commonMethods = new MethodsCommon();
        PdfProcessing pdfProcessing = new PdfProcessing();
        Indigo5600DistCalc indigo5600DistCalc = new Indigo5600DistCalc();

        public void PreflightPdfLayoutCombined(FormMain mainForm, string[] passedFile, string outputFile, float stockWidth, float stockHeight, string press)
        {
            stockWidth = stockWidth * 72;
            stockHeight = stockHeight * 72;

            FileStream fs = new FileStream(outputFile, FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(stockWidth, stockHeight));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            foreach (string s in passedFile)
            {
                List<float> sizes = new List<float>();
                List<int> resultsCalculate = new List<int>();
                List<float> resultsStartDraw = new List<float>();
                float bleed = pdfProcessing.PdfResize(s);
                sizes = pdfProcessing.GetSize(Settings.Default.tempDir + "\\" + Path.GetFileName(s), bleed);
                Console.WriteLine("Number of Pages " + sizes[0]);
                sizes.RemoveAt(0);
                while (sizes.Count > 0)
                {
                    resultsCalculate = pdfProcessing.Calculate(stockWidth.ToString(), stockHeight.ToString(), (sizes[0] - bleed).ToString(), (sizes[1] - bleed).ToString(), sizes[2].ToString());
                    Console.WriteLine(sizes[0] + " x " + sizes[1]);
                    /////////////////
                    Console.WriteLine("# Up: " + resultsCalculate[0]);
                    Console.WriteLine("# Horizontal: " + resultsCalculate[1]);
                    Console.WriteLine("# Vertical: " + resultsCalculate[2]);
                    Console.WriteLine("Rotation: " + resultsCalculate[3]);
                    ///////////////////
                    if (resultsCalculate[0] <= 0)
                    {
                        throw new Exception(Path.GetFileName(s) + " will not fit on the Indigo");
                    }
                    else
                    {
                        doc.NewPage();                        
                        resultsStartDraw = pdfProcessing.PdfStartDraw(stockWidth, stockHeight, sizes[0], sizes[1], resultsCalculate[1], resultsCalculate[2], resultsCalculate[3]);
                        float hStep = 0;
                        float vStep = 0;
                        int h = resultsCalculate[1];
                        int v = resultsCalculate[2];

                        while (v != 0)
                        {
                            while (h != 0)
                            {
                                pdfProcessing.PdfPlacement(writer, Settings.Default.tempDir + "\\" + Path.GetFileName(s), resultsStartDraw[0] + hStep, resultsStartDraw[1] + vStep, resultsCalculate[3], 1);

                                if (resultsCalculate[3] == 0)
                                {
                                    hStep = hStep + sizes[0];
                                }
                                else
                                {
                                    hStep = hStep + sizes[1];
                                }
                                h--;
                            }

                            if (resultsCalculate[3] == 0)
                            {
                                vStep = vStep + sizes[1];
                            }
                            else
                            {
                                vStep = vStep + sizes[0];
                            }
                            //Reset Horizontal Values to allow for stepping vertically
                            hStep = 0;
                            h = resultsCalculate[1];

                            v--;
                        }
                        hStep = 0;
                        vStep = 0;

                        pdfProcessing.PdfDrawCropMarks(writer, stockWidth, stockHeight, sizes[0], sizes[1], resultsCalculate[1], resultsCalculate[2], resultsCalculate[3], sizes[2], press);                   
                        
                    }
                    sizes.RemoveRange(0, 3);
                }
                resultsCalculate.Clear();
            }
            doc.Close();
            
        }

        public List<string> PreflightPdfLayoutCombinedNew(FormMain mainForm, string[] passedFile, string outputFile, float stockWidth, float stockHeight, string press, int[] qty)
        {
            stockWidth = stockWidth * 72;
            stockHeight = stockHeight * 72;

            FileStream fs = new FileStream(outputFile, FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(stockWidth, stockHeight));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            List<string> returnList = new List<string>();
            int i = 0;
            //while (!isDone)
            //{                
                List<float> sizes = new List<float>();
                List<int> resultsCalculate = new List<int>();
                List<float> resultsStartDraw = new List<float>();
                float bleed = pdfProcessing.PdfResize(passedFile[i]);
                sizes = pdfProcessing.GetSize(Settings.Default.tempDir + "\\" + Path.GetFileName(passedFile[0]), bleed);
                Console.WriteLine("Number of Pages " + sizes[0]);
                sizes.RemoveAt(0);
                while (sizes.Count > 0)
                {
                    resultsCalculate = pdfProcessing.Calculate(stockWidth.ToString(), stockHeight.ToString(), (sizes[0] - bleed).ToString(), (sizes[1] - bleed).ToString(), sizes[2].ToString());
                    var listOfPdfDist = indigo5600DistCalc.Indigo5600Dist(mainForm, passedFile.ToList(), qty.ToList(), resultsCalculate[0]);
                    returnList = listOfPdfDist.Item2;
                    Console.WriteLine(sizes[0] + " x " + sizes[1]);
                    /////////////////
                    Console.WriteLine("# Up: " + resultsCalculate[0]);
                    Console.WriteLine("# Horizontal: " + resultsCalculate[1]);
                    Console.WriteLine("# Vertical: " + resultsCalculate[2]);
                    Console.WriteLine("Rotation: " + resultsCalculate[3]);
                    //////////////////
                    if (resultsCalculate[0] <= 0)
                    {
                        throw new Exception(Path.GetFileName(listOfPdfDist.Item1[0]) + " will not fit on the Indigo");
                    }
                    else
                    {
                        while (i < listOfPdfDist.Item1.Count())
                        {
                            doc.NewPage();
                            resultsStartDraw = pdfProcessing.PdfStartDraw(stockWidth, stockHeight, sizes[0], sizes[1], resultsCalculate[1], resultsCalculate[2], resultsCalculate[3]);

                            float hStep = 0;
                            float vStep = 0;
                            int h = resultsCalculate[1];
                            int v = resultsCalculate[2];

                            while (v != 0)
                            {
                                while (h != 0)
                                {
                                    pdfProcessing.PdfPlacement(writer, Settings.Default.tempDir + "\\" + Path.GetFileName(listOfPdfDist.Item1[i]), resultsStartDraw[0] + hStep, resultsStartDraw[1] + vStep, resultsCalculate[3], 1);
                                    i++;
                                    if (i == (listOfPdfDist.Item1.Count - 1))
                                    {
                                        //isDone = true;
                                    }
                                    if (resultsCalculate[3] == 0)
                                    {
                                        hStep = hStep + sizes[0];
                                    }
                                    else
                                    {
                                        hStep = hStep + sizes[1];
                                    }
                                    h--;
                                }

                                if (resultsCalculate[3] == 0)
                                {
                                    vStep = vStep + sizes[1];
                                }
                                else
                                {
                                    vStep = vStep + sizes[0];
                                }
                                //Reset Horizontal Values to allow for stepping vertically
                                hStep = 0;
                                h = resultsCalculate[1];

                                v--;
                            }
                            hStep = 0;
                            vStep = 0;

                        PdfContentByte cb = writer.DirectContent;
                        cb.SetCMYKColorFill(0, 0, 0, 255);
                        cb.Circle(22.5f, 72f, 9f);
                        cb.Circle(stockWidth - 22.5f, 72f, 9f);
                        cb.Circle(22.5f, stockHeight - 72f, 9f);
                        cb.Circle(stockWidth - 22.5f, stockHeight - 72f, 9f);
                        cb.Fill();
                        pdfProcessing.PdfDrawCropMarks(writer, stockWidth, stockHeight, sizes[0], sizes[1], resultsCalculate[1], resultsCalculate[2], resultsCalculate[3], sizes[2], press);
                        }
                    }
                    sizes.RemoveRange(0, 3);
                }
                resultsCalculate.Clear();                
            //}
            doc.Close();
            return returnList;
        }

        public void PreflightPdfPrint(FormMain mainForm, string passedFile)
        {
            //passedFile = Settings.Default.hotFolder + "\\" + passedFile;         
            string gillRFont = "Fonts\\GIL_____.TTF";
            BaseFont GillSansR = BaseFont.CreateFont(gillRFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileName(passedFile), FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc1 = new Document();
            PdfReader inputFile = new PdfReader(passedFile);
            PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
            writer1.PdfVersion = PdfWriter.VERSION_1_3;
            int pageCount = inputFile.NumberOfPages;
            int fileProgressStep = (int)Math.Ceiling(((double)100) / pageCount);
            bool tabloid = true;
            float paperWidth = 792;
            float paperHeight = 1224;
            doc1.Open();
            for (int i = 1; i <= pageCount; i++)
            {
                PdfImportedPage page = writer1.GetImportedPage(inputFile, i);
                iTextSharp.text.Rectangle fileSize = inputFile.GetBoxSize(i, "media");
                doc1.SetMargins(0, 0, 0, 0);
                bool rotate = false;
                float scalePercent = 1;
                if (((fileSize.Width * fileSize.Height) < 414720) || ((fileSize.Width == 612) && fileSize.Height == 792) || ((fileSize.Width == 792) && fileSize.Height == 612))
                {
                    tabloid = false;
                    //mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + tabloid +"\r\n", Color.Red, FontStyle.Regular); }));
                    paperWidth = 612;
                    paperHeight = 792;
                    doc1.SetPageSize(new iTextSharp.text.Rectangle(paperWidth, paperHeight));
                }
                else
                {
                    paperWidth = 792;
                    paperHeight = 1224;
                    doc1.SetPageSize(new iTextSharp.text.Rectangle(paperWidth, paperHeight));
                }

                float widthScale = fileSize.Width;
                float heightScale = fileSize.Height;
                float xPosition = (paperWidth - fileSize.Width) / 2;
                float yPosition = (paperHeight - fileSize.Height) / 2;

                if (fileSize.Width > fileSize.Height)
                {
                    rotate = true;
                    xPosition = (((paperWidth - fileSize.Height) / 2) + fileSize.Height);
                    yPosition = (paperHeight - fileSize.Width) / 2;
                }
                if ((fileSize.Width > paperWidth || fileSize.Height > paperHeight) && !rotate)
                {
                    widthScale = paperWidth / fileSize.Width;
                    heightScale = paperHeight / fileSize.Height;

                    if (widthScale < heightScale)
                    {
                        scalePercent = widthScale;
                    }
                    else
                    {
                        scalePercent = heightScale;
                    }
                    scalePercent = scalePercent - .05f;
                    xPosition = ((paperWidth - (fileSize.Width * scalePercent)) / 2) / scalePercent;
                    yPosition = ((paperHeight - (fileSize.Height * scalePercent)) / 2) / scalePercent;
                }
                if ((fileSize.Width > paperHeight || fileSize.Height > paperWidth) && rotate)
                {
                    widthScale = paperHeight / fileSize.Width;
                    heightScale = paperWidth / fileSize.Height;
                    if (widthScale < heightScale)
                    {
                        scalePercent = widthScale;
                    }
                    else
                    {
                        scalePercent = heightScale;
                    }
                    scalePercent = scalePercent - .05f;
                    xPosition = (((paperWidth - (fileSize.Height * scalePercent)) / 2) + (fileSize.Height * scalePercent)) / scalePercent;
                    yPosition = ((paperHeight - (fileSize.Width * scalePercent)) / 2) / scalePercent;
                    //xPosition = (paperHeight - (fileSizeTrim.Height * scalePercent));
                    //yPosition = (paperWidth - (fileSizeTrim.Width * scalePercent));
                }
                doc1.NewPage();
                PdfReader pdfFile = new PdfReader(passedFile);
                PdfImportedPage pdfPage = writer1.GetImportedPage(pdfFile, i);
                PdfContentByte cb = writer1.DirectContent;
                var placePdf = new System.Drawing.Drawing2D.Matrix();
                placePdf.Scale(scalePercent, scalePercent);
                placePdf.Translate(xPosition, yPosition);
                if (rotate)
                {
                    placePdf.Rotate(90);
                }
                writer1.DirectContent.AddTemplate(pdfPage, placePdf);
                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                cb.BeginText();
                cb.SetFontAndSize(GillSansR, 12);
                cb.SetTextMatrix(36, 36);
                cb.ShowText(Path.GetFileNameWithoutExtension(passedFile) + " - Pg: " + i);
                cb.EndText();
            }
            doc1.Close();
            if (!Settings.Default.debugOn)
            {
                commonMethods.SendToPrinter(Settings.Default.tempDir + "\\" + Path.GetFileName(passedFile), false, tabloid);
            }
        }        
    }
}