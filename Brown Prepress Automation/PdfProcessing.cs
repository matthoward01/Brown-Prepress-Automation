using System;
using System.Collections.Generic;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Globalization;
using System.Windows.Forms;
using Brown_Prepress_Automation.Properties;
using System.Text;
using iTextSharp.text.pdf.parser;

namespace Brown_Prepress_Automation
{
    class PdfProcessing
    {
        public List<float> GetSize(string file, float bleed)
        {
            List<float> sizeList = new List<float>();
            PdfReader inputFile = new PdfReader(file);
            sizeList.Add(inputFile.NumberOfPages);
            for (int i = 1; i <= inputFile.NumberOfPages; i++)
            {
                
                Rectangle fileSizeMedia = inputFile.GetBoxSize(i, "media");

                sizeList.Add((fileSizeMedia.Width));
                sizeList.Add((fileSizeMedia.Height));
                sizeList.Add((fileSizeMedia.Width + bleed) - fileSizeMedia.Width);
            }

            //Number of Pages -> Width1 -> Height1 -> Width2 -> Height2 -> etc
            return sizeList;
        }

        public string FormatGetSize(string file, string pdfBox, int page)
        {            
            PdfReader inputFile = new PdfReader(file);

            Rectangle fileSizeTrim = inputFile.GetBoxSize(page, pdfBox);
                
            string formattedSize = (fileSizeTrim.Width / 72).ToString("0.00") + " x " + (fileSizeTrim.Height / 72).ToString("0.00");

            return formattedSize;
        }
        
        public float PdfResize(string file)
        {
            FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc1 = new Document();

            //Reading File
            PdfReader inputFile = new PdfReader(file);
            Rectangle fileSizeTrim = inputFile.GetBoxSize(1, "trim");
            Rectangle fileSizeBleed = inputFile.GetBoxSize(1, "bleed");
            Rectangle fileSizeMedia = inputFile.GetBoxSize(1, "media");
            float fileSizeBleedWidth = fileSizeBleed.Width;
            float fileSizeBleedHeight = fileSizeBleed.Height;
            //Writing File
            PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
            writer1.PdfVersion = PdfWriter.VERSION_1_3;
            if (((fileSizeTrim.Width * fileSizeTrim.Height) <= 1119744) && (fileSizeBleed.Height == fileSizeTrim.Height))
            {
                fileSizeBleedWidth = fileSizeBleedWidth + 18;
                fileSizeBleedHeight = fileSizeBleedHeight + 18;
            }
            doc1.SetPageSize(new Rectangle(fileSizeBleedWidth, fileSizeBleedHeight));           
            doc1.SetMargins(0, 0, 0, 0);
            doc1.Open();
            doc1.NewPage();
            var imp = writer1.GetImportedPage(inputFile, 1);
            var tm = new System.Drawing.Drawing2D.Matrix();
            PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
            if (fileSizeMedia.Height != fileSizeBleedHeight)
            {
                tm.Translate(((fileSizeBleedWidth - fileSizeMedia.Width)/2), ((fileSizeBleedHeight - fileSizeMedia.Height)/2));
            }
            else
            {
                tm.Translate(0f, 0f);
            }
            writer1.DirectContent.AddTemplate(imp, tm);
            doc1.Close();
            float bleed = fileSizeBleedWidth - fileSizeTrim.Width;
            /*if ((fileSizeTrim.Width * fileSizeTrim.Height) <= 216)
            {
                bleed = 18;
            }*/

            return bleed;
        }

        public void SplitPdf(string file, string destPath)
        {
            PdfReader inputFile = new PdfReader(file);
            for (int pageNumber = 1; pageNumber <= inputFile.NumberOfPages; pageNumber++)
            {
                string fileName = Path.GetFileNameWithoutExtension(file) + " - " + pageNumber + ".pdf";

                Document document = new Document();
                PdfCopy pdfCopy = new PdfCopy(document, new FileStream(destPath + "\\" + fileName, FileMode.Create));

                document.Open();

                pdfCopy.AddPage(pdfCopy.GetImportedPage(inputFile, pageNumber));

                document.Close();
            }
        }

        public void SplitPdf2Set(string file, string destPath)
        {
            PdfReader inputFile = new PdfReader(file);
            for (int pageNumber = 1; pageNumber <= inputFile.NumberOfPages; pageNumber++)
            {
                string fileName = Path.GetFileNameWithoutExtension(file) + " - " + pageNumber + ".pdf";

                Document document = new Document();
                PdfCopy pdfCopy = new PdfCopy(document, new FileStream(destPath + "\\" + fileName, FileMode.Create));

                document.Open();

                pdfCopy.AddPage(pdfCopy.GetImportedPage(inputFile, pageNumber));
                if (pageNumber != inputFile.NumberOfPages)
                {
                    pageNumber++;
                    pdfCopy.AddPage(pdfCopy.GetImportedPage(inputFile, pageNumber));
                }

                document.Close();
            }
        }

        public int GetPdfTotalPages(string file)
        {
            PdfReader inputFile = new PdfReader(file);

            return inputFile.NumberOfPages;
        }

        public void SplitPdfCleanup(string file)
        {
            if (File.Exists(file))
            {
                try
                {
                    File.Delete(file);
                }

                catch (IOException ex)
                {
                    Console.WriteLine(ex.Message);
                    //rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
                    //rtMain.AppendText(DateTime.Now + " | " + ex.Message, Color.Red, FontStyle.Regular);
                    //rtMain.AppendText("-------------------------------------------------------------\r\n", Color.Black, FontStyle.Regular);
                }
            }
        }

        public string pdfToText(string file)
        {
            StringBuilder pdfSB= new StringBuilder();
            try
            {
                if (File.Exists(file))
                {
                    PdfReader reader = new PdfReader(file);

                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        //ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        //string currentText = PdfTextExtractor.GetTextFromPage(reader, i, strategy);
                        string currentText = PdfTextExtractor.GetTextFromPage(reader, i);

                        pdfSB.Append(currentText);
                    }
                    reader.Close();
                }
            }
            catch (Exception ex)
            {

            }
            return pdfSB.ToString();
        }

        public List<int> Calculate(string stockWidth, string stockHeight, string fileWidth, string fileHeight, string gutter)
        {
            List<int> results = new List<int>();
            string qty = "0";
            string guttersH = gutter;
            string guttersV = gutter;
            string reservedH = ".375";
            string reservedV = ".75";

            try
            {
                if (qty == "")
                {
                    qty = "0";
                }
                float qtyFloat = float.Parse(qty, CultureInfo.InvariantCulture);
                float stockWidthFloat = float.Parse(stockWidth, CultureInfo.InvariantCulture);
                float stockHeightFloat = float.Parse(stockHeight, CultureInfo.InvariantCulture);
                float fileWidthFloat = float.Parse(fileWidth, CultureInfo.InvariantCulture);
                float fileHeightFloat = float.Parse(fileHeight, CultureInfo.InvariantCulture);
                float gutterHFloat = float.Parse(guttersH, CultureInfo.InvariantCulture);
                float gutterVFloat = float.Parse(guttersV, CultureInfo.InvariantCulture);
                float edgeH = float.Parse(reservedH, CultureInfo.InvariantCulture);
                float edgeV = float.Parse(reservedV, CultureInfo.InvariantCulture);
                edgeH = edgeH * 72;
                edgeV = edgeV * 72;

                //No Rotation
                float resultsH1 = (stockWidthFloat - edgeH) / (fileWidthFloat + gutterHFloat);
                float resultsV1 = (stockHeightFloat - edgeV) / (fileHeightFloat + gutterVFloat);
                int nUp1 = (int)resultsH1 * (int)resultsV1;

                //Rotated
                float resultsH2 = (stockWidthFloat - edgeH) / (fileHeightFloat + gutterHFloat);
                float resultsV2 = (stockHeightFloat - edgeV) / (fileWidthFloat + gutterVFloat);
                int nUp2 = (int)resultsH2 * (int)resultsV2;

                if (nUp1 >= nUp2)
                {
                    results.Add(nUp1);
                    results.Add((int)resultsH1);
                    results.Add((int)resultsV1);
                    results.Add(0);
                    results.Add((int)Math.Ceiling(qtyFloat / nUp1));
                }
                else
                {
                    results.Add(nUp2);
                    results.Add((int)resultsH2);
                    results.Add((int)resultsV2);
                    results.Add(90);
                    results.Add((int)Math.Ceiling(qtyFloat / nUp2));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //Number UP -> Horizonal -> Veritcal -> Rotation -> Qty
            return results;
        }

        public List<float> PdfStartDraw(float stockWidth, float stockHeight, float width, float height, int across, int down, int rotation)
        {
            float hPlace;
            float vPlace;
            List<float> results = new List<float>();
            if (rotation == 0)
            {
                hPlace = (stockWidth - (width * across)) / 2;
                vPlace = (stockHeight - (height * down)) / 2;
            }
            else
            {
                hPlace = ((stockWidth - (height * across)) / 2) + height;
                vPlace = (stockHeight - (width * down)) / 2;
            }

            results.Add(hPlace);
            results.Add(vPlace);

            return results;
        }

        public void PdfDrawCropMarks(PdfWriter writer, float stockWidth, float stockHeight, float width, float height, int across, int down, int rotation, float gutter, string press)
        {
            PdfContentByte cb = writer.DirectContentUnder;
            float h1;
            float v1;
            float h2;
            float v2;
            int hLineCount;
            int vLineCount;
            float offset;
            float hStep = 0;
            float vStep = 0;

            if (gutter != 0)
            {
                hLineCount = down * 2;
                vLineCount = across * 2;
                offset = gutter / 2;
            }
            else
            {
                hLineCount = down + 1;
                vLineCount = across + 1;
                offset = 0;
            }

            if (rotation == 0)
            {
                h1 = (stockWidth - (width * across)) / 2;
                v1 = (stockHeight - (height * down)) / 2;
                h2 = stockWidth - h1;
                v2 = stockHeight - v1;
            }
            else
            {
                h1 = (stockWidth - (height * across)) / 2;
                v1 = (stockHeight - (width * down)) / 2;
                h2 = stockWidth - h1;
                v2 = stockHeight - v1;
            }

            if (press != "6800")
            {
                for (int v = 0; v < vLineCount; v++)
                {
                    cb.MoveTo((h1 + offset) + hStep, v1 - 18f);
                    cb.LineTo((h1 + offset) + hStep, v2 + 18f);
                    cb.Stroke();
                    if (rotation == 0)
                    {
                        if ((v % 2 != 0) && (v != 0) && (gutter != 0))
                        {
                            hStep = hStep + gutter;
                        }
                        else
                        {
                            hStep = hStep + width - gutter;
                        }
                    }
                    else
                    {
                        if ((v % 2 != 0) && (v != 0) && (gutter != 0))
                        {
                            hStep = hStep + gutter;
                        }
                        else
                        {
                            hStep = hStep + height - gutter;
                        }
                    }
                }
            }

            if (press == "6800")
            {
                cb.SetLineWidth(18f);
            }
            for (int h = 0; h < hLineCount; h++)
            {
                cb.MoveTo(h1 - 18f, (v1 + offset) + vStep);
                cb.LineTo(h2 + 18f, (v1 + offset) + vStep);
                cb.Stroke();
                if (rotation == 0)
                {
                    if ((h % 2 != 0) && (h != 0) && (gutter != 0))
                    {
                        vStep = vStep + gutter;
                    }
                    else
                    {
                        vStep = vStep + height - gutter;
                    }
                }
                else
                {
                    if ((h % 2 != 0) && (h != 0) && (gutter != 0))
                    {
                        vStep = vStep + gutter;
                    }
                    else
                    {
                        vStep = vStep + width - gutter;
                    }
                }
            }

            
            cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
            cb.MoveTo(h1, v1);
            cb.LineTo(h2, v1);
            cb.LineTo(h2, v2);
            cb.LineTo(h1, v2);
            cb.Fill();
        }

        public void PdfPlacement(PdfWriter writer, string pdfName, float hPlacement, float vPlacement, float rotation, int pdfPageNumber)
        {
            PdfReader pdfFile = new PdfReader(pdfName);
            PdfImportedPage pdfImportedPage = writer.GetImportedPage(pdfFile, pdfPageNumber);
            var pdfDrawing = new System.Drawing.Drawing2D.Matrix();
            pdfDrawing.Translate(hPlacement, vPlacement);
            pdfDrawing.Rotate(rotation);
            writer.DirectContent.AddTemplate(pdfImportedPage, pdfDrawing);
        }
    }
}
