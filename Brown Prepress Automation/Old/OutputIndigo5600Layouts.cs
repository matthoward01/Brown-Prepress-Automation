using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Brown_Prepress_Automation.Properties;

namespace Brown_Prepress_Automation
{
    class OutputIndigo5600LayoutsOld
    {
        PdfProcessing pdfProcessing = new PdfProcessing();
        ////////////////////////////
        //////////5600 LAYOUTS////////
        ////////////////////////////

        public void pdf1x1_5(string filename, string[] art, int[] qty)
        {
            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            item = art.ToList();
            itemQty = qty.ToList();

            foreach (string i in item)
            {
                System.IO.File.Copy(i, Settings.Default.Oldprintable5600_1x1_5 + Path.GetFileNameWithoutExtension(i) + " - Printable.pdf", true);
            }
        }

        public void pdf3_5x1_5(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(270f, 126f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 126f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();

                //methods.jpgCreate(Settings.Default.tempDir+ "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", "\\temp\\jpgs\\" + Path.GetFileNameWithoutExtension(file) + ".jpg");
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();
            while (item.Count() > 0)
            {
                int count = itemQty[0] + 5;
                while (count > 0)
                {
                    itemTotal.Add(item[0]);
                    count--;
                }
                item.RemoveAt(0);
                itemQty.RemoveAt(0);
            }
            while (itemTotal.Count() % 27 != 0)
            {
                itemTotal.Add("Blank");
            }

            string path = Settings.Default.tempDir;
            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                float stepDistance = 0;

                //Minus 18 to remove for Bleed
                cb.MoveTo(18, 90);
                cb.LineTo(846, 90);
                cb.Stroke();

                for (int i = 1; i <= 9; i++)
                {
                    //Row 1
                    PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                    PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                    var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                    var R1C1 = new System.Drawing.Drawing2D.Matrix();
                    R1C1.Translate(27f, 81f + stepDistance);
                    writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                    PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                    PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                    var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                    var R1C2 = new System.Drawing.Drawing2D.Matrix();
                    R1C2.Translate(297f, 81f + stepDistance);
                    writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                    PdfReader R1C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                    PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                    var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                    var R1C3 = new System.Drawing.Drawing2D.Matrix();
                    R1C3.Translate(567f, 81f + stepDistance);
                    writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                    stepDistance = stepDistance + 126;

                    //Use Original step for Gutter Mark
                    cb.MoveTo(18f, (90f + stepDistance - 18));
                    cb.LineTo(846f, (90f + stepDistance - 18));
                    cb.Stroke();

                    if (i == 9)
                    {
                        //Use Original step for Gutter Mark
                        cb.MoveTo(18f, (90f + stepDistance));
                        cb.LineTo(846f, (90f + stepDistance));
                        cb.Stroke();
                    }

                    

                    itemTotal.RemoveRange(0, 3);
                }

                stepDistance = 0;
                
                //Cropmark Vertical
                cb.MoveTo(36, 72);
                cb.LineTo(36, 1224);
                cb.Stroke();
                cb.MoveTo(288, 72);
                cb.LineTo(288, 1224);
                cb.Stroke();
                cb.MoveTo(306, 72);
                cb.LineTo(306, 1224);
                cb.Stroke();
                cb.MoveTo(558, 72);
                cb.LineTo(558, 1224);
                cb.Stroke();
                cb.MoveTo(576, 72);
                cb.LineTo(576, 1224);
                cb.Stroke();
                cb.MoveTo(828, 72);
                cb.LineTo(828, 1224);
                cb.Stroke();                              

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27, 81);
                cb.LineTo(837, 81);
                cb.LineTo(837, 1215);
                cb.LineTo(27, 1215);
                cb.Fill();
            }
            doc.Close();
        }

        public void pdf3_5x1_75(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(270f, 144f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 144f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();

                //methods.jpgCreate(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", "\\temp\\jpgs\\" + Path.GetFileNameWithoutExtension(file) + ".jpg");
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();
            while (item.Count() > 0)
            {
                int count = itemQty[0] + 5;
                while (count > 0)
                {
                    itemTotal.Add(item[0]);
                    count--;
                }
                item.RemoveAt(0);
                itemQty.RemoveAt(0);
            }
            while (itemTotal.Count() % 24 != 0)
            {
                itemTotal.Add("Blank");
            }

            string path = Settings.Default.tempDir;
            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(27f, 72f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(297f, 72f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(567f, 72f);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(27f, 216f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(297f, 216f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(567f, 216f);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(27f, 360f);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(297f, 360f);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(567f, 360f);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                //Row 4
                PdfReader R4C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(27f, 504f);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(297f, 504f);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(567f, 504f);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);

                //Row 5
                PdfReader R5C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[12]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(27f, 648f);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[13]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(297f, 648f);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                PdfReader R5C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[14]) + ".pdf");
                PdfImportedPage R5C3Page = writer.GetImportedPage(R5C3File, 1);
                var R5C3PDF = writer.GetImportedPage(R5C3File, 1);
                var R5C3 = new System.Drawing.Drawing2D.Matrix();
                R5C3.Translate(567f, 648f);
                writer.DirectContent.AddTemplate(R5C3Page, R5C3);

                //Row 6
                PdfReader R6C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[15]) + ".pdf");
                PdfImportedPage R6C1Page = writer.GetImportedPage(R6C1File, 1);
                var R6C1PDF = writer.GetImportedPage(R6C1File, 1);
                var R6C1 = new System.Drawing.Drawing2D.Matrix();
                R6C1.Translate(27f, 792f);
                writer.DirectContent.AddTemplate(R6C1Page, R6C1);

                PdfReader R6C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[16]) + ".pdf");
                PdfImportedPage R6C2Page = writer.GetImportedPage(R6C2File, 1);
                var R6C2PDF = writer.GetImportedPage(R6C2File, 1);
                var R6C2 = new System.Drawing.Drawing2D.Matrix();
                R6C2.Translate(297f, 792f);
                writer.DirectContent.AddTemplate(R6C2Page, R6C2);

                PdfReader R6C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[17]) + ".pdf");
                PdfImportedPage R6C3Page = writer.GetImportedPage(R6C3File, 1);
                var R6C3PDF = writer.GetImportedPage(R6C3File, 1);
                var R6C3 = new System.Drawing.Drawing2D.Matrix();
                R6C3.Translate(567f, 792f);
                writer.DirectContent.AddTemplate(R6C3Page, R6C3);

                //Row 7
                PdfReader R7C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[18]) + ".pdf");
                PdfImportedPage R7C1Page = writer.GetImportedPage(R7C1File, 1);
                var R7C1PDF = writer.GetImportedPage(R7C1File, 1);
                var R7C1 = new System.Drawing.Drawing2D.Matrix();
                R7C1.Translate(27f, 936f);
                writer.DirectContent.AddTemplate(R7C1Page, R7C1);

                PdfReader R7C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[19]) + ".pdf");
                PdfImportedPage R7C2Page = writer.GetImportedPage(R7C2File, 1);
                var R7C2PDF = writer.GetImportedPage(R7C2File, 1);
                var R7C2 = new System.Drawing.Drawing2D.Matrix();
                R7C2.Translate(297f, 936f);
                writer.DirectContent.AddTemplate(R7C2Page, R7C2);

                PdfReader R7C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[20]) + ".pdf");
                PdfImportedPage R7C3Page = writer.GetImportedPage(R7C3File, 1);
                var R7C3PDF = writer.GetImportedPage(R7C3File, 1);
                var R7C3 = new System.Drawing.Drawing2D.Matrix();
                R7C3.Translate(567f, 936f);
                writer.DirectContent.AddTemplate(R7C3Page, R7C3);

                //Row 8
                PdfReader R8C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[21]) + ".pdf");
                PdfImportedPage R8C1Page = writer.GetImportedPage(R8C1File, 1);
                var R8C1PDF = writer.GetImportedPage(R8C1File, 1);
                var R8C1 = new System.Drawing.Drawing2D.Matrix();
                R8C1.Translate(27f, 1080f);
                writer.DirectContent.AddTemplate(R8C1Page, R8C1);

                PdfReader R8C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[22]) + ".pdf");
                PdfImportedPage R8C2Page = writer.GetImportedPage(R8C2File, 1);
                var R8C2PDF = writer.GetImportedPage(R8C2File, 1);
                var R8C2 = new System.Drawing.Drawing2D.Matrix();
                R8C2.Translate(297f, 1080f);
                writer.DirectContent.AddTemplate(R8C2Page, R8C2);

                PdfReader R8C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[23]) + ".pdf");
                PdfImportedPage R8C3Page = writer.GetImportedPage(R8C3File, 1);
                var R8C3PDF = writer.GetImportedPage(R8C3File, 1);
                var R8C3 = new System.Drawing.Drawing2D.Matrix();
                R8C3.Translate(567f, 1080f);
                writer.DirectContent.AddTemplate(R8C3Page, R8C3);

                //Cropmark Vertical
                cb.MoveTo(36, 63);
                cb.LineTo(36, 1233);
                cb.Stroke();
                cb.MoveTo(288, 63);
                cb.LineTo(288, 1233);
                cb.Stroke();
                cb.MoveTo(306, 63);
                cb.LineTo(306, 1233);
                cb.Stroke();
                cb.MoveTo(558, 63);
                cb.LineTo(558, 1233);
                cb.Stroke();
                cb.MoveTo(576, 63);
                cb.LineTo(576, 1233);
                cb.Stroke();
                cb.MoveTo(828, 63);
                cb.LineTo(828, 1233);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(18, 81);
                cb.LineTo(846, 81);
                cb.Stroke();
                cb.MoveTo(18, 207);
                cb.LineTo(846, 207);
                cb.Stroke();
                cb.MoveTo(18, 225);
                cb.LineTo(846, 225);
                cb.Stroke();
                cb.MoveTo(18, 351);
                cb.LineTo(846, 351);
                cb.Stroke();
                cb.MoveTo(18, 369);
                cb.LineTo(846, 369);
                cb.Stroke();
                cb.MoveTo(18, 495);
                cb.LineTo(846, 495);
                cb.Stroke();
                cb.MoveTo(18, 513);
                cb.LineTo(846, 513);
                cb.Stroke();
                cb.MoveTo(18, 639);
                cb.LineTo(846, 639);
                cb.Stroke();
                cb.MoveTo(18, 657);
                cb.LineTo(846, 657);
                cb.Stroke();
                cb.MoveTo(18, 783);
                cb.LineTo(846, 783);
                cb.Stroke();
                cb.MoveTo(18, 801);
                cb.LineTo(846, 801);
                cb.Stroke();
                cb.MoveTo(18, 927);
                cb.LineTo(846, 927);
                cb.Stroke();
                cb.MoveTo(18, 945);
                cb.LineTo(846, 945);
                cb.Stroke();
                cb.MoveTo(18, 1071);
                cb.LineTo(846, 1071);
                cb.Stroke();
                cb.MoveTo(18, 1089);
                cb.LineTo(846, 1089);
                cb.Stroke();
                cb.MoveTo(18, 1215);
                cb.LineTo(846, 1215);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27, 72);
                cb.LineTo(837, 72);
                cb.LineTo(837, 1224);
                cb.LineTo(27, 1224);
                cb.Fill();

                itemTotal.RemoveRange(0, 24);
            }
            doc.Close();
        }

        public void pdf4_5x1_25(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(342f, 108f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 108f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();

                //methods.jpgCreate(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", "\\temp\\jpgs\\" + Path.GetFileNameWithoutExtension(file) + ".jpg");
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();
            while (item.Count() > 0)
            {
                int count = itemQty[0]+5;
                while (count > 0)
                {
                    itemTotal.Add(item[0]);
                    count--;
                }
                item.RemoveAt(0);
                itemQty.RemoveAt(0);
            }
            while (itemTotal.Count() % 22 != 0)
            {
                itemTotal.Add("Blank");
            }

            string path = Settings.Default.tempDir;
            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                float stepIncrement = 0;
                int z = 0;
                for (int i = 0; i < 11; i++)
                {
                    pdfProcessing.pdfPlacement(writer, path + Path.GetFileNameWithoutExtension(itemTotal[z]) + ".pdf", 90, 54+stepIncrement, 0, 1);
                    z++;
                    pdfProcessing.pdfPlacement(writer, path + Path.GetFileNameWithoutExtension(itemTotal[z]) + ".pdf", 432, 54+stepIncrement, 0, 1);
                    z++;
                    stepIncrement = stepIncrement + 108f;
                }
                stepIncrement = 0;

                pdfProcessing.pdfDrawCropMarks(writer, 864, 1296, 342, 108, 2, 11, 0, 18, "indigo");

                itemTotal.RemoveRange(0, 22);
            }
            doc.Close();
        }

        public void pdf3_5x2(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(270f, 162f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 162f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();

                //methods.jpgCreate(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", "\\temp\\jpgs\\" + Path.GetFileNameWithoutExtension(file) + ".jpg");
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();
            while (item.Count() > 0)
            {
                int count = itemQty[0] + 5;
                while (count > 0)
                {
                    itemTotal.Add(item[0]);
                    count--;
                }
                item.RemoveAt(0);
                itemQty.RemoveAt(0);
            }
            while (itemTotal.Count() % 24 != 0)
            {
                itemTotal.Add("Blank");
            }

            string path = Settings.Default.tempDir;
            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(27f, 81f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(297f, 81f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(567f, 81f);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(27f, 243f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(297f, 243f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(567f, 243f);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(27f, 405f);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(297f, 405f);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(567f, 405f);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                //Row 4
                PdfReader R4C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(27f, 567f);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(297f, 567f);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(567f, 567f);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);

                //Row 5
                PdfReader R5C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[12]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(27f, 729f);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[13]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(297f, 729f);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                PdfReader R5C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[14]) + ".pdf");
                PdfImportedPage R5C3Page = writer.GetImportedPage(R5C3File, 1);
                var R5C3PDF = writer.GetImportedPage(R5C3File, 1);
                var R5C3 = new System.Drawing.Drawing2D.Matrix();
                R5C3.Translate(567f, 729f);
                writer.DirectContent.AddTemplate(R5C3Page, R5C3);

                //Row 6
                PdfReader R6C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[15]) + ".pdf");
                PdfImportedPage R6C1Page = writer.GetImportedPage(R6C1File, 1);
                var R6C1PDF = writer.GetImportedPage(R6C1File, 1);
                var R6C1 = new System.Drawing.Drawing2D.Matrix();
                R6C1.Translate(27f, 891f);
                writer.DirectContent.AddTemplate(R6C1Page, R6C1);

                PdfReader R6C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[16]) + ".pdf");
                PdfImportedPage R6C2Page = writer.GetImportedPage(R6C2File, 1);
                var R6C2PDF = writer.GetImportedPage(R6C2File, 1);
                var R6C2 = new System.Drawing.Drawing2D.Matrix();
                R6C2.Translate(297f, 891f);
                writer.DirectContent.AddTemplate(R6C2Page, R6C2);

                PdfReader R6C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[17]) + ".pdf");
                PdfImportedPage R6C3Page = writer.GetImportedPage(R6C3File, 1);
                var R6C3PDF = writer.GetImportedPage(R6C3File, 1);
                var R6C3 = new System.Drawing.Drawing2D.Matrix();
                R6C3.Translate(567f, 891f);
                writer.DirectContent.AddTemplate(R6C3Page, R6C3);

                //Row 7
                PdfReader R7C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[18]) + ".pdf");
                PdfImportedPage R7C1Page = writer.GetImportedPage(R7C1File, 1);
                var R7C1PDF = writer.GetImportedPage(R7C1File, 1);
                var R7C1 = new System.Drawing.Drawing2D.Matrix();
                R7C1.Translate(27f, 1053f);
                writer.DirectContent.AddTemplate(R7C1Page, R7C1);

                PdfReader R7C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[19]) + ".pdf");
                PdfImportedPage R7C2Page = writer.GetImportedPage(R7C2File, 1);
                var R7C2PDF = writer.GetImportedPage(R7C2File, 1);
                var R7C2 = new System.Drawing.Drawing2D.Matrix();
                R7C2.Translate(297f, 1053f);
                writer.DirectContent.AddTemplate(R7C2Page, R7C2);

                PdfReader R7C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[20]) + ".pdf");
                PdfImportedPage R7C3Page = writer.GetImportedPage(R7C3File, 1);
                var R7C3PDF = writer.GetImportedPage(R7C3File, 1);
                var R7C3 = new System.Drawing.Drawing2D.Matrix();
                R7C3.Translate(567f, 1053f);
                writer.DirectContent.AddTemplate(R7C3Page, R7C3);

                //Cropmark Vertical
                cb.MoveTo(36, 72);
                cb.LineTo(36, 1224);
                cb.Stroke();
                cb.MoveTo(288, 72);
                cb.LineTo(288, 1224);
                cb.Stroke();
                cb.MoveTo(306, 72);
                cb.LineTo(306, 1224);
                cb.Stroke();
                cb.MoveTo(558, 72);
                cb.LineTo(558, 1224);
                cb.Stroke();
                cb.MoveTo(576, 72);
                cb.LineTo(576, 1224);
                cb.Stroke();
                cb.MoveTo(828, 72);
                cb.LineTo(828, 1224);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(18, 90);
                cb.LineTo(846, 90);
                cb.Stroke();
                cb.MoveTo(18, 234);
                cb.LineTo(846, 234);
                cb.Stroke();
                cb.MoveTo(18, 252);
                cb.LineTo(846, 252);
                cb.Stroke();
                cb.MoveTo(18, 396);
                cb.LineTo(846, 396);
                cb.Stroke();
                cb.MoveTo(18, 414);
                cb.LineTo(846, 414);
                cb.Stroke();
                cb.MoveTo(18, 558);
                cb.LineTo(846, 558);
                cb.Stroke();
                cb.MoveTo(18, 576);
                cb.LineTo(846, 576);
                cb.Stroke();
                cb.MoveTo(18, 720);
                cb.LineTo(846, 720);
                cb.Stroke();
                cb.MoveTo(18, 738);
                cb.LineTo(846, 738);
                cb.Stroke();
                cb.MoveTo(18, 882);
                cb.LineTo(846, 882);
                cb.Stroke();
                cb.MoveTo(18, 900);
                cb.LineTo(846, 900);
                cb.Stroke();
                cb.MoveTo(18, 1044);
                cb.LineTo(846, 1044);
                cb.Stroke();
                cb.MoveTo(18, 1062);
                cb.LineTo(846, 1062);
                cb.Stroke();
                cb.MoveTo(18, 1206);
                cb.LineTo(846, 1206);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27, 81);
                cb.LineTo(837, 81);
                cb.LineTo(837, 1215);
                cb.LineTo(27, 1215);
                cb.Fill();

                itemTotal.RemoveRange(0, 24);
            }
            doc.Close();
        }

        public void pdf3_5x2_5(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(270f, 198f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 198f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();

                //methods.jpgCreate(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", "\\temp\\jpgs\\" + Path.GetFileNameWithoutExtension(file) + ".jpg");
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;
            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(27f, 54f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(297f, 54f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(567f, 54f);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(27f, 252f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(297f, 252f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(567f, 252f);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(27f, 450f);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(297f, 450f);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(567f, 450f);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                //Row 4
                PdfReader R4C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(27f, 648f);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(297f, 648f);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(567f, 648f);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);

                //Row 5
                PdfReader R5C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(27f, 846f);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(297f, 846f);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                PdfReader R5C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C3Page = writer.GetImportedPage(R5C3File, 1);
                var R5C3PDF = writer.GetImportedPage(R5C3File, 1);
                var R5C3 = new System.Drawing.Drawing2D.Matrix();
                R5C3.Translate(567f, 846f);
                writer.DirectContent.AddTemplate(R5C3Page, R5C3);

                //Row 6
                PdfReader R6C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R6C1Page = writer.GetImportedPage(R6C1File, 1);
                var R6C1PDF = writer.GetImportedPage(R6C1File, 1);
                var R6C1 = new System.Drawing.Drawing2D.Matrix();
                R6C1.Translate(27f, 1044f);
                writer.DirectContent.AddTemplate(R6C1Page, R6C1);

                PdfReader R6C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R6C2Page = writer.GetImportedPage(R6C2File, 1);
                var R6C2PDF = writer.GetImportedPage(R6C2File, 1);
                var R6C2 = new System.Drawing.Drawing2D.Matrix();
                R6C2.Translate(297f, 1044f);
                writer.DirectContent.AddTemplate(R6C2Page, R6C2);

                PdfReader R6C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R6C3Page = writer.GetImportedPage(R6C3File, 1);
                var R6C3PDF = writer.GetImportedPage(R6C3File, 1);
                var R6C3 = new System.Drawing.Drawing2D.Matrix();
                R6C3.Translate(567f, 1044f);
                writer.DirectContent.AddTemplate(R6C3Page, R6C3);


                //Cropmark Vertical
                cb.MoveTo(36, 45);
                cb.LineTo(36, 1251);
                cb.Stroke();
                cb.MoveTo(288, 45);
                cb.LineTo(288, 1251);
                cb.Stroke();
                cb.MoveTo(306, 45);
                cb.LineTo(306, 1251);
                cb.Stroke();
                cb.MoveTo(558, 45);
                cb.LineTo(558, 1251);
                cb.Stroke();
                cb.MoveTo(576, 45);
                cb.LineTo(576, 1251);
                cb.Stroke();
                cb.MoveTo(828, 45);
                cb.LineTo(828, 1251);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(18, 63f);
                cb.LineTo(846, 63f);
                cb.Stroke();
                cb.MoveTo(18, 243);
                cb.LineTo(846, 243);
                cb.Stroke();
                cb.MoveTo(18, 261);
                cb.LineTo(846, 261);
                cb.Stroke();
                cb.MoveTo(18, 441);
                cb.LineTo(846, 441);
                cb.Stroke();
                cb.MoveTo(18, 459);
                cb.LineTo(846, 459);
                cb.Stroke();
                cb.MoveTo(18, 639);
                cb.LineTo(846, 639);
                cb.Stroke();
                cb.MoveTo(18, 657);
                cb.LineTo(846, 657);
                cb.Stroke();
                cb.MoveTo(18, 837);
                cb.LineTo(846, 837);
                cb.Stroke();
                cb.MoveTo(18, 855);
                cb.LineTo(846, 855);
                cb.Stroke();
                cb.MoveTo(18, 1035);
                cb.LineTo(846, 1035);
                cb.Stroke();
                cb.MoveTo(18, 1053);
                cb.LineTo(846, 1053);
                cb.Stroke();
                cb.MoveTo(18, 1233);
                cb.LineTo(846, 1233);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27, 54);
                cb.LineTo(837, 54);
                cb.LineTo(837, 1242);
                cb.LineTo(27, 1242);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf3_5x5(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(270f, 378f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 378f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(27f, 81f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(297f, 81f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(567f, 81f);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(27f, 459f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(297f, 459f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(567f, 459f);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(27f, 837f);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(297f, 837f);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(567f, 837f);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);


                //Cropmark Vertical
                cb.MoveTo(36, 72);
                cb.LineTo(36, 1224);
                cb.Stroke();
                cb.MoveTo(288, 72);
                cb.LineTo(288, 1224);
                cb.Stroke();
                cb.MoveTo(306, 72);
                cb.LineTo(306, 1224);
                cb.Stroke();
                cb.MoveTo(558, 72);
                cb.LineTo(558, 1224);
                cb.Stroke();
                cb.MoveTo(576, 72);
                cb.LineTo(576, 1224);
                cb.Stroke();
                cb.MoveTo(828, 72);
                cb.LineTo(828, 1224);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(18, 90);
                cb.LineTo(846, 90);
                cb.Stroke();
                cb.MoveTo(18, 450);
                cb.LineTo(846, 450);
                cb.Stroke();
                cb.MoveTo(18, 468);
                cb.LineTo(846, 468);
                cb.Stroke();
                cb.MoveTo(18, 828);
                cb.LineTo(846, 828);
                cb.Stroke();
                cb.MoveTo(18, 846);
                cb.LineTo(846, 846);
                cb.Stroke();
                cb.MoveTo(18, 1206);
                cb.LineTo(846, 1206);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27, 108);
                cb.LineTo(837, 108);
                cb.LineTo(837, 1188);
                cb.LineTo(27, 1188);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf3_25x1_75(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(252f, 144f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 144f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();
            while (item.Count() > 0)
            {
                int count = itemQty[0] + 5;
                while (count > 0)
                {
                    itemTotal.Add(item[0]);
                    count--;
                }
                item.RemoveAt(0);
                itemQty.RemoveAt(0);
            }
            while (itemTotal.Count() % 24 != 0)
            {
                itemTotal.Add("Blank");
            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(54f, 72f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(306f, 72f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(558f, 72f);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(54f, 216f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(306f, 216f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(558f, 216f);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(54f, 360f);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(306f, 360f);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(558f, 360f);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                //Row 4
                PdfReader R4C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(54f, 504f);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(306f, 504f);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(558f, 504f);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);

                //Row 5
                PdfReader R5C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[12]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(54f, 648f);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[13]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(306f, 648f);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                PdfReader R5C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[14]) + ".pdf");
                PdfImportedPage R5C3Page = writer.GetImportedPage(R5C3File, 1);
                var R5C3PDF = writer.GetImportedPage(R5C3File, 1);
                var R5C3 = new System.Drawing.Drawing2D.Matrix();
                R5C3.Translate(558f, 648f);
                writer.DirectContent.AddTemplate(R5C3Page, R5C3);

                //Row 6
                PdfReader R6C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[15]) + ".pdf");
                PdfImportedPage R6C1Page = writer.GetImportedPage(R6C1File, 1);
                var R6C1PDF = writer.GetImportedPage(R6C1File, 1);
                var R6C1 = new System.Drawing.Drawing2D.Matrix();
                R6C1.Translate(54f, 792f);
                writer.DirectContent.AddTemplate(R6C1Page, R6C1);

                PdfReader R6C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[16]) + ".pdf");
                PdfImportedPage R6C2Page = writer.GetImportedPage(R6C2File, 1);
                var R6C2PDF = writer.GetImportedPage(R6C2File, 1);
                var R6C2 = new System.Drawing.Drawing2D.Matrix();
                R6C2.Translate(306f, 792f);
                writer.DirectContent.AddTemplate(R6C2Page, R6C2);

                PdfReader R6C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[17]) + ".pdf");
                PdfImportedPage R6C3Page = writer.GetImportedPage(R6C3File, 1);
                var R6C3PDF = writer.GetImportedPage(R6C3File, 1);
                var R6C3 = new System.Drawing.Drawing2D.Matrix();
                R6C3.Translate(558f, 792f);
                writer.DirectContent.AddTemplate(R6C3Page, R6C3);

                //Row 7
                PdfReader R7C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[18]) + ".pdf");
                PdfImportedPage R7C1Page = writer.GetImportedPage(R7C1File, 1);
                var R7C1PDF = writer.GetImportedPage(R7C1File, 1);
                var R7C1 = new System.Drawing.Drawing2D.Matrix();
                R7C1.Translate(54f, 936f);
                writer.DirectContent.AddTemplate(R7C1Page, R7C1);

                PdfReader R7C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[19]) + ".pdf");
                PdfImportedPage R7C2Page = writer.GetImportedPage(R7C2File, 1);
                var R7C2PDF = writer.GetImportedPage(R7C2File, 1);
                var R7C2 = new System.Drawing.Drawing2D.Matrix();
                R7C2.Translate(306f, 936f);
                writer.DirectContent.AddTemplate(R7C2Page, R7C2);

                PdfReader R7C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[20]) + ".pdf");
                PdfImportedPage R7C3Page = writer.GetImportedPage(R7C3File, 1);
                var R7C3PDF = writer.GetImportedPage(R7C3File, 1);
                var R7C3 = new System.Drawing.Drawing2D.Matrix();
                R7C3.Translate(558f, 936f);
                writer.DirectContent.AddTemplate(R7C3Page, R7C3);

                //Row 8
                PdfReader R8C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[21]) + ".pdf");
                PdfImportedPage R8C1Page = writer.GetImportedPage(R8C1File, 1);
                var R8C1PDF = writer.GetImportedPage(R8C1File, 1);
                var R8C1 = new System.Drawing.Drawing2D.Matrix();
                R8C1.Translate(54f, 1080f);
                writer.DirectContent.AddTemplate(R8C1Page, R8C1);

                PdfReader R8C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[22]) + ".pdf");
                PdfImportedPage R8C2Page = writer.GetImportedPage(R8C2File, 1);
                var R8C2PDF = writer.GetImportedPage(R8C2File, 1);
                var R8C2 = new System.Drawing.Drawing2D.Matrix();
                R8C2.Translate(306f, 1080f);
                writer.DirectContent.AddTemplate(R8C2Page, R8C2);

                PdfReader R8C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[23]) + ".pdf");
                PdfImportedPage R8C3Page = writer.GetImportedPage(R8C3File, 1);
                var R8C3PDF = writer.GetImportedPage(R8C3File, 1);
                var R8C3 = new System.Drawing.Drawing2D.Matrix();
                R8C3.Translate(558f, 1080f);
                writer.DirectContent.AddTemplate(R8C3Page, R8C3);

                //Cropmark Vertical
                cb.MoveTo(63, 63);
                cb.LineTo(63, 1233);
                cb.Stroke();
                cb.MoveTo(297, 63);
                cb.LineTo(297, 1233);
                cb.Stroke();
                cb.MoveTo(315, 63);
                cb.LineTo(315, 1233);
                cb.Stroke();
                cb.MoveTo(549, 63);
                cb.LineTo(549, 1233);
                cb.Stroke();
                cb.MoveTo(567, 63);
                cb.LineTo(567, 1233);
                cb.Stroke();
                cb.MoveTo(801, 63);
                cb.LineTo(801, 1233);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(45, 81);
                cb.LineTo(819, 81);
                cb.Stroke();
                cb.MoveTo(45, 207);
                cb.LineTo(819, 207);
                cb.Stroke();
                cb.MoveTo(45, 225);
                cb.LineTo(819, 225);
                cb.Stroke();
                cb.MoveTo(45, 351);
                cb.LineTo(819, 351);
                cb.Stroke();
                cb.MoveTo(45, 369);
                cb.LineTo(819, 369);
                cb.Stroke();
                cb.MoveTo(45, 495);
                cb.LineTo(819, 495);
                cb.Stroke();
                cb.MoveTo(45, 513);
                cb.LineTo(819, 513);
                cb.Stroke();
                cb.MoveTo(45, 639);
                cb.LineTo(819, 639);
                cb.Stroke();
                cb.MoveTo(45, 657);
                cb.LineTo(819, 657);
                cb.Stroke();
                cb.MoveTo(45, 783);
                cb.LineTo(819, 783);
                cb.Stroke();
                cb.MoveTo(45, 801);
                cb.LineTo(819, 801);
                cb.Stroke();
                cb.MoveTo(45, 927);
                cb.LineTo(819, 927);
                cb.Stroke();
                cb.MoveTo(45, 945);
                cb.LineTo(819, 945);
                cb.Stroke();
                cb.MoveTo(45, 1071);
                cb.LineTo(819, 1071);
                cb.Stroke();
                cb.MoveTo(45, 1089);
                cb.LineTo(819, 1089);
                cb.Stroke();
                cb.MoveTo(45, 1215);
                cb.LineTo(819, 1215);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(54, 72);
                cb.LineTo(810, 72);
                cb.LineTo(810, 1224);
                cb.LineTo(54, 1224);
                cb.Fill();

                itemTotal.RemoveRange(0, 24);
            }
            doc.Close();
        }

        public void pdf3_25x3_25(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(252f, 252f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 252f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(54f, 144f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(306f, 144f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(558f, 144f);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(54f, 396);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(306f, 396);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(558f, 396);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(54f, 648);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(306f, 648);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(558f, 648);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                //Row 4
                PdfReader R4C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(54f, 900);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(306f, 900);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(558f, 900);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);

                //Cropmark Vertical
                cb.MoveTo(63, 135);
                cb.LineTo(63, 1161);
                cb.Stroke();
                cb.MoveTo(297, 135);
                cb.LineTo(297, 1161);
                cb.Stroke();
                cb.MoveTo(315, 135);
                cb.LineTo(315, 1161);
                cb.Stroke();
                cb.MoveTo(549, 135);
                cb.LineTo(549, 1161);
                cb.Stroke();
                cb.MoveTo(567, 135);
                cb.LineTo(567, 1161);
                cb.Stroke();
                cb.MoveTo(801, 135);
                cb.LineTo(801, 1161);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(45, 153);
                cb.LineTo(819, 153);
                cb.Stroke();
                cb.MoveTo(45, 387);
                cb.LineTo(819, 387);
                cb.Stroke();
                cb.MoveTo(45, 405);
                cb.LineTo(819, 405);
                cb.Stroke();
                cb.MoveTo(45, 639);
                cb.LineTo(819, 639);
                cb.Stroke();
                cb.MoveTo(45, 657);
                cb.LineTo(819, 657);
                cb.Stroke();
                cb.MoveTo(45, 891);
                cb.LineTo(819, 891);
                cb.Stroke();
                cb.MoveTo(45, 909);
                cb.LineTo(819, 909);
                cb.Stroke();
                cb.MoveTo(45, 1143);
                cb.LineTo(819, 1143);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(54, 144);
                cb.LineTo(810, 144);
                cb.LineTo(810, 1152);
                cb.LineTo(54, 1152);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf3_25x3_5(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(252f, 270f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 270f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(54f, 108f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(306f, 108f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(558f, 108f);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(54f, 378);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(306f, 378);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(558f, 378);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(54f, 648);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(306f, 648);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(558f, 648);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                //Row 4
                PdfReader R4C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(54f, 918);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(306f, 918);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(558f, 918);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);

                //Cropmark Vertical
                cb.MoveTo(63, 99);
                cb.LineTo(63, 1197);
                cb.Stroke();
                cb.MoveTo(297, 99);
                cb.LineTo(297, 1197);
                cb.Stroke();
                cb.MoveTo(315, 99);
                cb.LineTo(315, 1197);
                cb.Stroke();
                cb.MoveTo(549, 99);
                cb.LineTo(549, 1197);
                cb.Stroke();
                cb.MoveTo(567, 99);
                cb.LineTo(567, 1197);
                cb.Stroke();
                cb.MoveTo(801, 99);
                cb.LineTo(801, 1197);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(45, 117);
                cb.LineTo(819, 117);
                cb.Stroke();
                cb.MoveTo(45, 369);
                cb.LineTo(819, 369);
                cb.Stroke();
                cb.MoveTo(45, 387);
                cb.LineTo(819, 387);
                cb.Stroke();
                cb.MoveTo(45, 639);
                cb.LineTo(819, 639);
                cb.Stroke();
                cb.MoveTo(45, 657);
                cb.LineTo(819, 657);
                cb.Stroke();
                cb.MoveTo(45, 909);
                cb.LineTo(819, 909);
                cb.Stroke();
                cb.MoveTo(45, 927);
                cb.LineTo(819, 927);
                cb.Stroke();
                cb.MoveTo(45, 1179);
                cb.LineTo(819, 1179);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(54, 108);
                cb.LineTo(810, 108);
                cb.LineTo(810, 1188);
                cb.LineTo(54, 1188);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf3_25x4(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(252f, 306f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 306f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(54f, 36f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(306f, 36f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(558f, 36f);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(54f, 342);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(306f, 342);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(558f, 342);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(54f, 648);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(306f, 648);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(558f, 648);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                //Row 4
                PdfReader R4C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(54f, 954);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(306f, 954);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(558f, 954);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);

                //Cropmark Vertical
                cb.MoveTo(63, 27);
                cb.LineTo(63, 1269);
                cb.Stroke();
                cb.MoveTo(297, 27);
                cb.LineTo(297, 1269);
                cb.Stroke();
                cb.MoveTo(315, 27);
                cb.LineTo(315, 1269);
                cb.Stroke();
                cb.MoveTo(549, 27);
                cb.LineTo(549, 1269);
                cb.Stroke();
                cb.MoveTo(567, 27);
                cb.LineTo(567, 1269);
                cb.Stroke();
                cb.MoveTo(801, 27);
                cb.LineTo(801, 1269);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(45, 45);
                cb.LineTo(819, 45);
                cb.Stroke();
                cb.MoveTo(45, 333);
                cb.LineTo(819, 333);
                cb.Stroke();
                cb.MoveTo(45, 351);
                cb.LineTo(819, 351);
                cb.Stroke();
                cb.MoveTo(45, 639);
                cb.LineTo(819, 639);
                cb.Stroke();
                cb.MoveTo(45, 657);
                cb.LineTo(819, 657);
                cb.Stroke();
                cb.MoveTo(45, 945);
                cb.LineTo(819, 945);
                cb.Stroke();
                cb.MoveTo(45, 963);
                cb.LineTo(819, 963);
                cb.Stroke();
                cb.MoveTo(45, 1251);
                cb.LineTo(819, 1251);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(54, 144);
                cb.LineTo(810, 144);
                cb.LineTo(810, 1152);
                cb.LineTo(54, 1152);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf4x3(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(306f, 234f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 234f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(315, 36);
                R1C1.Rotate(90);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(549, 36);
                R1C2.Rotate(90);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(783, 36);
                R1C3.Rotate(90);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(315, 342);
                R2C1.Rotate(90);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(549, 342);
                R2C2.Rotate(90);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(783, 342);
                R2C3.Rotate(90);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(315, 648);
                R3C1.Rotate(90);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(549, 648);
                R3C2.Rotate(90);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(783, 648);
                R3C3.Rotate(90);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                //Row 4
                PdfReader R4C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(315, 954);
                R4C1.Rotate(90);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(549, 954);
                R4C2.Rotate(90);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(783, 954);
                R4C3.Rotate(90);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);


                //Cropmark Vertical
                cb.MoveTo(90, 27);
                cb.LineTo(90, 1269);
                cb.Stroke();
                cb.MoveTo(306, 27);
                cb.LineTo(306, 1269);
                cb.Stroke();
                cb.MoveTo(324, 27);
                cb.LineTo(324, 1269);
                cb.Stroke();
                cb.MoveTo(540, 27);
                cb.LineTo(540, 1269);
                cb.Stroke();
                cb.MoveTo(558, 27);
                cb.LineTo(558, 1269);
                cb.Stroke();
                cb.MoveTo(774, 27);
                cb.LineTo(774, 1269);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(72, 45);
                cb.LineTo(792, 45);
                cb.Stroke();
                cb.MoveTo(72, 333);
                cb.LineTo(792, 333);
                cb.Stroke();
                cb.MoveTo(72, 351);
                cb.LineTo(792, 351);
                cb.Stroke();
                cb.MoveTo(72, 639);
                cb.LineTo(792, 639);
                cb.Stroke();
                cb.MoveTo(72, 657);
                cb.LineTo(792, 657);
                cb.Stroke();
                cb.MoveTo(72, 945);
                cb.LineTo(792, 945);
                cb.Stroke();
                cb.MoveTo(72, 963);
                cb.LineTo(792, 963);
                cb.Stroke();
                cb.MoveTo(72, 1251);
                cb.LineTo(792, 1251);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(81, 36);
                cb.LineTo(783, 36);
                cb.LineTo(783, 1260);
                cb.LineTo(81, 1260);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf4_5x1_625(string filename, string[] art, int[] qty)
        {
            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            item = art.ToList();
            itemQty = qty.ToList();

            foreach (string i in item)
            {
                System.IO.File.Copy(i, Settings.Default.Oldprintable5600_4_5x1_625 + Path.GetFileNameWithoutExtension(i) + " - Printable.pdf", true);
            }
        }

        public void pdf4_5x4_5(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(342f, 342f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 342f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(90f, 135f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(432f, 135f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(90f, 477f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(432f, 477f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(90f, 819f);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(432f, 819f);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                //Cropmark Vertical
                cb.MoveTo(99, 127);
                cb.LineTo(99, 1169);
                cb.Stroke();
                cb.MoveTo(423, 127);
                cb.LineTo(423, 1169);
                cb.Stroke();
                cb.MoveTo(441, 127);
                cb.LineTo(441, 1169);
                cb.Stroke();
                cb.MoveTo(765, 127);
                cb.LineTo(765, 1169);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(81, 144);
                cb.LineTo(783, 144);
                cb.Stroke();
                cb.MoveTo(81, 468);
                cb.LineTo(783, 468);
                cb.Stroke();
                cb.MoveTo(81, 486);
                cb.LineTo(783, 486);
                cb.Stroke();
                cb.MoveTo(81, 810);
                cb.LineTo(783, 810);
                cb.Stroke();
                cb.MoveTo(81, 828);
                cb.LineTo(783, 828);
                cb.Stroke();
                cb.MoveTo(81, 1152);
                cb.LineTo(783, 1152);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(90, 135);
                cb.LineTo(774, 135);
                cb.LineTo(774, 1161);
                cb.LineTo(90, 1161);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf4_5x6(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(342f, 450f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 450f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(90f, 198f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(432f, 198f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(90f, 648f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(432f, 648f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                //Cropmark Vertical
                cb.MoveTo(99, 189);
                cb.LineTo(99, 1107);
                cb.Stroke();
                cb.MoveTo(423, 189);
                cb.LineTo(423, 1107);
                cb.Stroke();
                cb.MoveTo(441, 189);
                cb.LineTo(441, 1107);
                cb.Stroke();
                cb.MoveTo(765, 189);
                cb.LineTo(765, 1107);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(81, 207);
                cb.LineTo(783, 207);
                cb.Stroke();
                cb.MoveTo(81, 639);
                cb.LineTo(783, 639);
                cb.Stroke();
                cb.MoveTo(81, 657);
                cb.LineTo(783, 657);
                cb.Stroke();
                cb.MoveTo(81, 1089);
                cb.LineTo(783, 1089);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(90, 198);
                cb.LineTo(774, 198);
                cb.LineTo(774, 1098);
                cb.LineTo(90, 1098);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf4_75x6(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(360f, 450f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 450f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(72f, 198f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(432f, 198f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(72f, 648f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(432f, 648f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                //Cropmark Vertical
                cb.MoveTo(81f, 189);
                cb.LineTo(81f, 1107);
                cb.Stroke();
                cb.MoveTo(423, 189);
                cb.LineTo(423, 1107);
                cb.Stroke();
                cb.MoveTo(441, 189);
                cb.LineTo(441, 1107);
                cb.Stroke();
                cb.MoveTo(783, 189);
                cb.LineTo(783, 1107);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(63, 207);
                cb.LineTo(801, 207);
                cb.Stroke();
                cb.MoveTo(63, 639);
                cb.LineTo(801, 639);
                cb.Stroke();
                cb.MoveTo(63, 657);
                cb.LineTo(801, 657);
                cb.Stroke();
                cb.MoveTo(63, 1089);
                cb.LineTo(801, 1089);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(72, 198);
                cb.LineTo(792, 198);
                cb.LineTo(792, 1098);
                cb.LineTo(72, 1098);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf4x6(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(306f, 450f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 450f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(126f, 198f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(432f, 198f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(126f, 648f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(432f, 648f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                //Cropmark Vertical
                cb.MoveTo(135, 189);
                cb.LineTo(135, 1107);
                cb.Stroke();
                cb.MoveTo(423, 189);
                cb.LineTo(423, 1107);
                cb.Stroke();
                cb.MoveTo(441, 189);
                cb.LineTo(441, 1107);
                cb.Stroke();
                cb.MoveTo(729, 189);
                cb.LineTo(729, 1107);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(117, 207);
                cb.LineTo(747, 207);
                cb.Stroke();
                cb.MoveTo(117, 639);
                cb.LineTo(747, 639);
                cb.Stroke();
                cb.MoveTo(117, 657);
                cb.LineTo(747, 657);
                cb.Stroke();
                cb.MoveTo(117, 1089);
                cb.LineTo(747, 1089);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(126, 198);
                cb.LineTo(738, 198);
                cb.LineTo(738, 1098);
                cb.LineTo(126, 1098);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf4x11(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(305.76f, 809.76f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 809.76f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(836.88f, 36.48f);
                R1C1.Rotate(90);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(836.88f, 342.24f);
                R2C1.Rotate(90);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(836.88f, 648f);
                R3C1.Rotate(90);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                //Row 4
                PdfReader R4C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(836.88f, 953.76f);
                R4C1.Rotate(90);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                //Cropmark Vertical
                cb.MoveTo(36, 27f);
                cb.LineTo(36, 1269);
                cb.Stroke();
                cb.MoveTo(828, 27);
                cb.LineTo(828, 1269);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(18, 44.52f);
                cb.LineTo(846, 44.52f);
                cb.Stroke();
                cb.MoveTo(18, 332.52f);
                cb.LineTo(846, 332.52f);
                cb.Stroke();
                cb.MoveTo(18, 350.52f);
                cb.LineTo(846, 350.52f);
                cb.Stroke();
                cb.MoveTo(18, 638.52f);
                cb.LineTo(846, 638.52f);
                cb.Stroke();
                cb.MoveTo(18, 656.52f);
                cb.LineTo(846, 656.52f);
                cb.Stroke();
                cb.MoveTo(18, 944.52f);
                cb.LineTo(846, 944.52f);
                cb.Stroke();
                cb.MoveTo(18, 962.52f);
                cb.LineTo(846, 962.52f);
                cb.Stroke();
                cb.MoveTo(18, 1250.52f);
                cb.LineTo(846, 1250.52f);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27.12f, 36.48f);
                cb.LineTo(836.88f, 36.48f);
                cb.LineTo(836.88f, 1259.52f);
                cb.LineTo(27.12f, 1259.52f);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf5_5x2_75(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(414f, 216f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 216f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(18f, 108f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(432f, 108f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(18f, 324f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(432f, 324f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(18f, 540f);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(432f, 540f);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                //Row 4
                PdfReader R4C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(18f, 758f);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(432f, 758f);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                //Row 5
                PdfReader R5C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(18f, 974f);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(432f, 974f);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                //Cropmark Vertical
                cb.MoveTo(27, 99);
                cb.LineTo(27, 1197);
                cb.Stroke();
                cb.MoveTo(423, 99);
                cb.LineTo(423, 1197);
                cb.Stroke();
                cb.MoveTo(441, 99);
                cb.LineTo(441, 1197);
                cb.Stroke();
                cb.MoveTo(837, 99);
                cb.LineTo(837, 1197);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(9, 117);
                cb.LineTo(855, 117);
                cb.Stroke();
                cb.MoveTo(9, 315);
                cb.LineTo(855, 315);
                cb.Stroke();
                cb.MoveTo(9, 333);
                cb.LineTo(855, 333);
                cb.Stroke();
                cb.MoveTo(9, 531);
                cb.LineTo(855, 531);
                cb.Stroke();
                cb.MoveTo(9, 549);
                cb.LineTo(855, 549);
                cb.Stroke();
                cb.MoveTo(9, 747);
                cb.LineTo(855, 747);
                cb.Stroke();
                cb.MoveTo(9, 765);
                cb.LineTo(855, 765);
                cb.Stroke();
                cb.MoveTo(9, 963);
                cb.LineTo(855, 963);
                cb.Stroke();
                cb.MoveTo(9, 981);
                cb.LineTo(855, 981);
                cb.Stroke();
                cb.MoveTo(9, 1179);
                cb.LineTo(855, 1179);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(18, 108);
                cb.LineTo(846, 108);
                cb.LineTo(846, 1188);
                cb.LineTo(18, 1188);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf5_5x5_5(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(414f, 414f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 414f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(18f, 27f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(432f, 27f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(18f, 441f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(432f, 441f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(18f, 855f);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(432f, 855f);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                //Cropmark Vertical
                cb.MoveTo(27, 18);
                cb.LineTo(27, 1278);
                cb.Stroke();
                cb.MoveTo(423, 18);
                cb.LineTo(423, 1278);
                cb.Stroke();
                cb.MoveTo(441, 18);
                cb.LineTo(441, 1278);
                cb.Stroke();
                cb.MoveTo(837, 18);
                cb.LineTo(837, 1278);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(9, 36);
                cb.LineTo(855, 36);
                cb.Stroke();
                cb.MoveTo(9, 432);
                cb.LineTo(855, 432);
                cb.Stroke();
                cb.MoveTo(9, 450);
                cb.LineTo(855, 450);
                cb.Stroke();
                cb.MoveTo(9, 846);
                cb.LineTo(855, 846);
                cb.Stroke();
                cb.MoveTo(9, 864);
                cb.LineTo(855, 864);
                cb.Stroke();
                cb.MoveTo(9, 1260);
                cb.LineTo(855, 1260);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(18, 27);
                cb.LineTo(846, 27);
                cb.LineTo(846, 1269);
                cb.LineTo(18, 1269);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf5_75x5_75(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(432f, 432f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 432f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                pdfProcessing.pdfPlacement(writer, path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf", 216f, 216f, 0, 1);

                //Row 2
                pdfProcessing.pdfPlacement(writer, path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf", 216f, 648f, 0, 1);

                //Cropmark Vertical

                pdfProcessing.pdfDrawCropMarks(writer, 864, 1296, 432, 432, 1, 2, 0, 18, "indigo");
                
                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf6x3(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(450f, 234f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 234f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(315f, 198f);
                R1C1.Rotate(90);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(549f, 198f);
                R1C2.Rotate(90);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(783f, 198f);
                R1C3.Rotate(90);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(315f, 648f);
                R2C1.Rotate(90);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(549f, 648f);
                R2C2.Rotate(90);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(783f, 648f);
                R2C3.Rotate(90);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                //Cropmark Vertical
                cb.MoveTo(90, 189);
                cb.LineTo(90, 1107);
                cb.Stroke();
                cb.MoveTo(306, 189);
                cb.LineTo(306, 1107);
                cb.Stroke();
                cb.MoveTo(324, 189);
                cb.LineTo(324, 1107);
                cb.Stroke();
                cb.MoveTo(540, 189);
                cb.LineTo(540, 1107);
                cb.Stroke();
                cb.MoveTo(558, 189);
                cb.LineTo(558, 1107);
                cb.Stroke();
                cb.MoveTo(774, 189);
                cb.LineTo(774, 1107);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(72, 207);
                cb.LineTo(792, 207);
                cb.Stroke();
                cb.MoveTo(72, 639);
                cb.LineTo(792, 639);
                cb.Stroke();
                cb.MoveTo(72, 657);
                cb.LineTo(792, 657);
                cb.Stroke();
                cb.MoveTo(72, 1089);
                cb.LineTo(792, 1089);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(81, 198);
                cb.LineTo(783, 198);
                cb.LineTo(783, 1098);
                cb.LineTo(81, 1098);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf6_5x7(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(486f, 522f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 522f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(693f, 200f);
                R1C1.Rotate(90);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(693f, 686f);
                R2C1.Rotate(90);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                //Cropmark Vertical
                cb.MoveTo(180, 191);
                cb.LineTo(180, 1181);
                cb.Stroke();
                cb.MoveTo(684, 191);
                cb.LineTo(684, 1181);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(162, 209.5f);
                cb.LineTo(702, 209.5f);
                cb.Stroke();
                cb.MoveTo(162, 677.5f);
                cb.LineTo(702, 677.5f);
                cb.Stroke();
                cb.MoveTo(162, 695.5f);
                cb.LineTo(702, 695.5f);
                cb.Stroke();
                cb.MoveTo(162, 1163.5f);
                cb.LineTo(702, 1163.5f);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(171, 200);
                cb.LineTo(693, 200);
                cb.LineTo(693, 1172);
                cb.LineTo(171, 1172);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf7x11(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(522f, 810f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 810f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(837f, 126f);
                R1C1.Rotate(90);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(837f, 648f);
                R2C1.Rotate(90);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                //Cropmark Vertical
                cb.MoveTo(36, 117);
                cb.LineTo(36, 1179);
                cb.Stroke();
                cb.MoveTo(828, 117);
                cb.LineTo(828, 1179);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(18, 135);
                cb.LineTo(846, 135);
                cb.Stroke();
                cb.MoveTo(18, 639);
                cb.LineTo(846, 639);
                cb.Stroke();
                cb.MoveTo(18, 657);
                cb.LineTo(846, 657);
                cb.Stroke();
                cb.MoveTo(18, 1161);
                cb.LineTo(846, 1161);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27, 126);
                cb.LineTo(837, 126);
                cb.LineTo(837, 1170);
                cb.LineTo(27, 1170);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf11x5_5(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(810f, 414f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 414f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;


            while (itemTotal.Count() > 0)
            {
                doc.NewPage();

                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(27f, 27f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(27f, 441f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(27f, 855f);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                //Cropmark Vertical
                cb.MoveTo(36, 18);
                cb.LineTo(36, 1278);
                cb.Stroke();
                cb.MoveTo(828, 18);
                cb.LineTo(828, 1278);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(18, 36);
                cb.LineTo(846, 36);
                cb.Stroke();
                cb.MoveTo(18, 432);
                cb.LineTo(846, 432);
                cb.Stroke();
                cb.MoveTo(18, 450);
                cb.LineTo(846, 450);
                cb.Stroke();
                cb.MoveTo(18, 846);
                cb.LineTo(846, 846);
                cb.Stroke();
                cb.MoveTo(18, 864);
                cb.LineTo(846, 864);
                cb.Stroke();
                cb.MoveTo(18, 1260);
                cb.LineTo(846, 1260);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27, 27);
                cb.LineTo(837, 27);
                cb.LineTo(837, 1269);
                cb.LineTo(27, 1269);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }

        public void pdf1up5600(string fileName, string[] art)
        {
            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContent;

            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();

            while (itemTotal.Count() > 0)
            {
                PdfReader inputFile = new PdfReader(itemTotal[0]);
                var imp = writer.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                doc.NewPage();
                iTextSharp.text.Rectangle fileSize = inputFile.GetBoxSize(1, "media");
                float labelXPlacement = ((doc.PageSize.Width - fileSize.Width) / 2);
                float labelYPlacement = (doc.PageSize.Height - fileSize.Height) / 2;
                tm.Translate(labelXPlacement, labelYPlacement);
                writer.DirectContent.AddTemplate(imp, tm);
                itemTotal.RemoveAt(0);

            }
            doc.Close();
        }

        public void pdf1up5600_13x19(string fileName, string[] art)
        {
            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(936, 1368));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContent;

            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();

            while (itemTotal.Count() > 0)
            {
                PdfReader inputFile = new PdfReader(itemTotal[0]);
                var imp = writer.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                doc.NewPage();
                iTextSharp.text.Rectangle fileSize = inputFile.GetBoxSize(1, "media");
                float labelXPlacement = ((doc.PageSize.Width - fileSize.Width) / 2);
                float labelYPlacement = (doc.PageSize.Height - fileSize.Height) / 2;
                tm.Translate(labelXPlacement, labelYPlacement);
                writer.DirectContent.AddTemplate(imp, tm);
                itemTotal.RemoveAt(0);

            }
            doc.Close();
        }

        public void pdf1up5600Rotated(string fileName, string[] art)
        {
            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContent;
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();

            while (itemTotal.Count() > 0)
            {

                PdfReader inputFile = new PdfReader(itemTotal[0]);
                var imp = writer.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                doc.NewPage();
                PdfImportedPage page = writer.GetImportedPage(inputFile, 1);
                iTextSharp.text.Rectangle fileSize = inputFile.GetBoxSize(1, "media");
                float labelXPlacement = ((doc.PageSize.Width - fileSize.Height) / 2) + fileSize.Height;
                float labelYPlacement = (doc.PageSize.Height - fileSize.Width) / 2;
                tm.Translate(labelXPlacement, labelYPlacement);
                tm.Rotate(90);
                writer.DirectContent.AddTemplate(imp, tm);
                itemTotal.RemoveAt(0);

            }
            doc.Close();
        }

        public void pdf17x5_5(string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(1242f, 414f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 414f)
                {
                    tm.Translate(-15.12f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + " - Printable.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(864, 1296));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemTotal = new List<string>();
            itemTotal = art.ToList();
            itemQty = qty.ToList();

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(432f, 27f);
                R1C1.Rotate(90);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(846f, 27f);
                R1C2.Rotate(90);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                //Cropmark Vertical
                cb.MoveTo(27, 18);
                cb.LineTo(27, 1278);
                cb.Stroke();
                cb.MoveTo(423, 18);
                cb.LineTo(423, 1278);
                cb.Stroke();
                cb.MoveTo(441, 18);
                cb.LineTo(441, 1278);
                cb.Stroke();
                cb.MoveTo(837, 18);
                cb.LineTo(837, 1278);
                cb.Stroke();

                //Cropmarks Horizontal
                cb.MoveTo(9, 36);
                cb.LineTo(855, 36);
                cb.Stroke();
                cb.MoveTo(9, 1260);
                cb.LineTo(855, 1260);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(18, 27);
                cb.LineTo(846, 27);
                cb.LineTo(846, 1269);
                cb.LineTo(18, 1269);
                cb.Fill();

                itemTotal.RemoveAt(0);
            }
            doc.Close();
        }    
    }
}
