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
    class OutputIndigo6800Shaw
    {
        ////////////////////////////////
        //////////6800 LAYOUTS//////////
        ////////////////////////////////
        PdfProcessing pdfProcessing = new PdfProcessing();

        public List<string> pdf6800Across5(FormMain mainForm, string filename, List<string> art, List<int> qty)
        {
            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art;
            itemQty = qty;

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
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[4]);
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
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }
            string formattedSizeTrim = pdfProcessing.FormatGetSize(itemTotal[0], "media", 1);
            var sizes = formattedSizeTrim.Split('x');
            double width = double.Parse(sizes[0]);
            double height = double.Parse(sizes[1]);
            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(filename) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.Open();
            doc.SetPageSize(new iTextSharp.text.Rectangle(((float)width) * 72, ((float)height) * 72));

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

            return diffPerPage;
        }

        public List<string> pdf0_5x0_5_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(45f, 45f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 45f)
                {
                    tm.Translate(-19.62f, -19.62f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2700));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    itemTotal.Add(itemPrint[0]);
                    itemTotal.Add(itemPrint[0]);
                    itemTotal.Add(itemPrint[0]);
                    itemTotal.Add(itemPrint[0]);
                    itemTotal.Add(itemPrint[0]);
                    itemTotal.Add(itemPrint[0]);
                    itemTotal.Add(itemPrint[0]);
                    itemTotal.Add(itemPrint[0]);
                    itemTotal.Add(itemPrint[0]);
                    itemTotal.Add(itemPrint[0]);
                    itemTotal.Add(itemPrint[0]);
                    itemPrint.RemoveAt(0);
                    printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 660);
                    diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                    itemQtyPrint.RemoveAt(0);
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                float stepDistance = 0;
                for (int i = 1; i <= 60; i++)
                {
                    PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                    PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                    var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                    var R1C1 = new System.Drawing.Drawing2D.Matrix();
                    R1C1.Translate(202.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                    PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                    PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                    var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                    var R1C2 = new System.Drawing.Drawing2D.Matrix();
                    R1C2.Translate(247.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                    PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                    PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                    var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                    var R1C3 = new System.Drawing.Drawing2D.Matrix();
                    R1C3.Translate(292.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                    PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                    PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                    var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                    var R1C4 = new System.Drawing.Drawing2D.Matrix();
                    R1C4.Translate(337.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                    PdfReader R1C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                    PdfImportedPage R1C5Page = writer.GetImportedPage(R1C5File, 1);
                    var R1C5PDF = writer.GetImportedPage(R1C5File, 1);
                    var R1C5 = new System.Drawing.Drawing2D.Matrix();
                    R1C5.Translate(382.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C5Page, R1C5);

                    PdfReader R1C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                    PdfImportedPage R1C6Page = writer.GetImportedPage(R1C6File, 1);
                    var R1C6PDF = writer.GetImportedPage(R1C6File, 1);
                    var R1C6 = new System.Drawing.Drawing2D.Matrix();
                    R1C6.Translate(427.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C6Page, R1C6);

                    PdfReader R1C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                    PdfImportedPage R1C7Page = writer.GetImportedPage(R1C7File, 1);
                    var R1C7PDF = writer.GetImportedPage(R1C7File, 1);
                    var R1C7 = new System.Drawing.Drawing2D.Matrix();
                    R1C7.Translate(472.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C7Page, R1C7);

                    PdfReader R1C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                    PdfImportedPage R1C8Page = writer.GetImportedPage(R1C8File, 1);
                    var R1C8PDF = writer.GetImportedPage(R1C8File, 1);
                    var R1C8 = new System.Drawing.Drawing2D.Matrix();
                    R1C8.Translate(517.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C8Page, R1C8);

                    PdfReader R1C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                    PdfImportedPage R1C9Page = writer.GetImportedPage(R1C9File, 1);
                    var R1C9PDF = writer.GetImportedPage(R1C9File, 1);
                    var R1C9 = new System.Drawing.Drawing2D.Matrix();
                    R1C9.Translate(562.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C9Page, R1C9);

                    PdfReader R1C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                    PdfImportedPage R1C10Page = writer.GetImportedPage(R1C10File, 1);
                    var R1C10PDF = writer.GetImportedPage(R1C10File, 1);
                    var R1C10 = new System.Drawing.Drawing2D.Matrix();
                    R1C10.Translate(607.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C10Page, R1C10);

                    PdfReader R1C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                    PdfImportedPage R1C11Page = writer.GetImportedPage(R1C11File, 1);
                    var R1C11PDF = writer.GetImportedPage(R1C11File, 1);
                    var R1C11 = new System.Drawing.Drawing2D.Matrix();
                    R1C11.Translate(652.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C11Page, R1C11);

                    stepDistance = stepDistance + 45;
                }
                stepDistance = 0;

                itemTotal.RemoveRange(0, 11);


                cb.SetLineWidth(18f);

                for (int i = 1; i <= 11; i++)
                {
                    //Cropmarks Horizontal
                    cb.MoveTo(180f, stepDistance);
                    cb.LineTo(720f, stepDistance);
                    cb.Stroke();
                    stepDistance = stepDistance + (45*6);
                }

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(198f, 0);
                cb.LineTo(702f, 0);
                cb.LineTo(702f, 2700);
                cb.LineTo(198f, 2700);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf1_5x1_5_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(126f, 126f));
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
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2646f));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 4 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemPrint.RemoveRange(0, 4);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 21);
                        diffPerPage.Add("4 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 4);
                    }
                    else if (itemPrint.Count() % 2 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemPrint.RemoveRange(0, 2);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 42);
                        diffPerPage.Add("2 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 2);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 84);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                float stepDistance = 0;
                for (int i = 1; i <= 21; i++)
                {
                    PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                    PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                    var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                    var R1C1 = new System.Drawing.Drawing2D.Matrix();
                    R1C1.Translate(198f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                    PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                    PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                    var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                    var R1C2 = new System.Drawing.Drawing2D.Matrix();
                    R1C2.Translate(324f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                    PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                    PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                    var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                    var R1C3 = new System.Drawing.Drawing2D.Matrix();
                    R1C3.Translate(450f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                    PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                    PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                    var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                    var R1C4 = new System.Drawing.Drawing2D.Matrix();
                    R1C4.Translate(576f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                    stepDistance = stepDistance + 126;
                }
                stepDistance = 0;

                itemTotal.RemoveRange(0, 4);


                cb.SetLineWidth(18f);

                for (int i = 1; i <= 8; i++)
                {
                    //Cropmarks Horizontal
                    cb.MoveTo(180f, stepDistance);
                    cb.LineTo(720f, stepDistance);
                    cb.Stroke();
                    stepDistance = stepDistance + (126 * 3);
                }

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(198f, 0);
                cb.LineTo(702f, 0);
                cb.LineTo(702f, 2646f);
                cb.LineTo(198f, 2646f);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf1_5x0_375_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(126f, 45f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 45f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2520f));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 4 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemPrint.RemoveRange(0, 4);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 56);
                        diffPerPage.Add("4 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 4);
                    }
                    else if (itemPrint.Count() % 2 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemPrint.RemoveRange(0, 2);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 112);
                        diffPerPage.Add("2 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 2);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 224);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                float stepDistance = 0;
                for (int i = 1; i <= 56; i++)
                {
                    PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                    PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                    var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                    var R1C1 = new System.Drawing.Drawing2D.Matrix();
                    R1C1.Translate(198f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                    PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                    PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                    var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                    var R1C2 = new System.Drawing.Drawing2D.Matrix();
                    R1C2.Translate(324f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                    PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                    PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                    var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                    var R1C3 = new System.Drawing.Drawing2D.Matrix();
                    R1C3.Translate(450f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                    PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                    PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                    var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                    var R1C4 = new System.Drawing.Drawing2D.Matrix();
                    R1C4.Translate(576f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                    stepDistance = stepDistance + 45f;
                }
                stepDistance = 0;

                itemTotal.RemoveRange(0, 4);


                cb.SetLineWidth(18f);

                for (int i = 1; i <= 8; i++)
                {
                    //Cropmarks Horizontal
                    cb.MoveTo(180f, stepDistance);
                    cb.LineTo(720f, stepDistance);
                    cb.Stroke();
                    stepDistance = stepDistance + (45f * 8);
                }

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(198f, 0);
                cb.LineTo(702f, 0);
                cb.LineTo(702f, 2520f);
                cb.LineTo(198f, 2520f);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf3_25x1_75_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(243f, 135f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 135f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2673));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 6 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[4]);
                        itemTotal.Add(itemPrint[5]);
                        itemPrint.RemoveRange(0, 6);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 11);
                        diffPerPage.Add("6 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 6);

                    }
                    else if (itemPrint.Count() % 3 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[2]);
                        itemPrint.RemoveRange(0, 3);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 22);
                        diffPerPage.Add("3 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 3);
                    }
                    else if (itemPrint.Count() % 2 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemPrint.RemoveRange(0, 2);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 33);
                        diffPerPage.Add("2 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 2);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 66);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }

            }

            /*while (itemTotal.Count() % 2 != 0)
            {
                itemTotal.Insert(0, itemTotal[0]);
            }*/

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                float stepDistance = -4.5f;
                for (int i = 1; i <= 11; i++)
                {
                    //Row 1
                    PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                    PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                    var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                    var R1C1 = new System.Drawing.Drawing2D.Matrix();
                    R1C1.Translate(184.5f, stepDistance);
                    R1C1.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                    PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                    PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                    var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                    var R1C2 = new System.Drawing.Drawing2D.Matrix();
                    R1C2.Translate(319.5f, stepDistance);
                    R1C2.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                    PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                    PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                    var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                    var R1C3 = new System.Drawing.Drawing2D.Matrix();
                    R1C3.Translate(454.5f, stepDistance);
                    R1C3.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                    PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                    PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                    var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                    var R1C4 = new System.Drawing.Drawing2D.Matrix();
                    R1C4.Translate(589.5f, stepDistance);
                    R1C4.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                    PdfReader R1C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                    PdfImportedPage R1C5Page = writer.GetImportedPage(R1C5File, 1);
                    var R1C5PDF = writer.GetImportedPage(R1C5File, 1);
                    var R1C5 = new System.Drawing.Drawing2D.Matrix();
                    R1C5.Translate(724.5f, stepDistance);
                    R1C5.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C5Page, R1C5);

                    PdfReader R1C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                    PdfImportedPage R1C6Page = writer.GetImportedPage(R1C6File, 1);
                    var R1C6PDF = writer.GetImportedPage(R1C6File, 1);
                    var R1C6 = new System.Drawing.Drawing2D.Matrix();
                    R1C6.Translate(859.5f, stepDistance);
                    R1C6.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C6Page, R1C6);

                    stepDistance = stepDistance + 243;
                }
                stepDistance = 0;

                itemTotal.RemoveRange(0, 6);

                cb.SetLineWidth(18f);

                //Cropmarks Horizontal
                for (int i = 1; i <= 12; i++)
                {
                    cb.MoveTo(22.5f, stepDistance);
                    cb.LineTo(877.5f, stepDistance);
                    cb.Stroke();
                    stepDistance = stepDistance + 243;
                }
                stepDistance = 0;

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(40.5f, 0);
                cb.LineTo(859.5f, 0);
                cb.LineTo(859.5f, 2673);
                cb.LineTo(40.5f, 2673);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf2x0_5_Short(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(162f, 54f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 54f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 1656));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 8 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[4]);
                        itemTotal.Add(itemPrint[4]);
                        itemTotal.Add(itemPrint[4]);
                        itemTotal.Add(itemPrint[5]);
                        itemTotal.Add(itemPrint[5]);
                        itemTotal.Add(itemPrint[5]);
                        itemTotal.Add(itemPrint[6]);
                        itemTotal.Add(itemPrint[6]);
                        itemTotal.Add(itemPrint[6]);
                        itemTotal.Add(itemPrint[7]);
                        itemTotal.Add(itemPrint[7]);
                        itemTotal.Add(itemPrint[7]);
                        itemPrint.RemoveRange(0, 8);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 15);
                        diffPerPage.Add("8 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 8);

                    }
                    else if (itemPrint.Count() % 4 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[3]);
                        itemPrint.RemoveRange(0, 4);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 30);
                        diffPerPage.Add("4 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 4);
                    }
                    else if (itemPrint.Count() % 2 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemPrint.RemoveRange(0, 2);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 60);
                        diffPerPage.Add("2 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 2);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 120);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }

            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();

                //Row 1
                PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(153f, 9f);
                R1C1.Rotate(90);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(207f, 9f);
                R1C2.Rotate(90);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(261f, 9f);
                R1C3.Rotate(90);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                var R1C4 = new System.Drawing.Drawing2D.Matrix();
                R1C4.Translate(333f, 9f);
                R1C4.Rotate(90);
                writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                PdfReader R1C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R1C5Page = writer.GetImportedPage(R1C5File, 1);
                var R1C5PDF = writer.GetImportedPage(R1C5File, 1);
                var R1C5 = new System.Drawing.Drawing2D.Matrix();
                R1C5.Translate(387f, 9f);
                R1C5.Rotate(90);
                writer.DirectContent.AddTemplate(R1C5Page, R1C5);

                PdfReader R1C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R1C6Page = writer.GetImportedPage(R1C6File, 1);
                var R1C6PDF = writer.GetImportedPage(R1C6File, 1);
                var R1C6 = new System.Drawing.Drawing2D.Matrix();
                R1C6.Translate(441f, 9f);
                R1C6.Rotate(90);
                writer.DirectContent.AddTemplate(R1C6Page, R1C6);

                PdfReader R1C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R1C7Page = writer.GetImportedPage(R1C7File, 1);
                var R1C7PDF = writer.GetImportedPage(R1C7File, 1);
                var R1C7 = new System.Drawing.Drawing2D.Matrix();
                R1C7.Translate(513f, 9f);
                R1C7.Rotate(90);
                writer.DirectContent.AddTemplate(R1C7Page, R1C7);

                PdfReader R1C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R1C8Page = writer.GetImportedPage(R1C8File, 1);
                var R1C8PDF = writer.GetImportedPage(R1C8File, 1);
                var R1C8 = new System.Drawing.Drawing2D.Matrix();
                R1C8.Translate(567f, 9f);
                R1C8.Rotate(90);
                writer.DirectContent.AddTemplate(R1C8Page, R1C8);

                PdfReader R1C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R1C9Page = writer.GetImportedPage(R1C9File, 1);
                var R1C9PDF = writer.GetImportedPage(R1C9File, 1);
                var R1C9 = new System.Drawing.Drawing2D.Matrix();
                R1C9.Translate(621f, 9f);
                R1C9.Rotate(90);
                writer.DirectContent.AddTemplate(R1C9Page, R1C9);

                PdfReader R1C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R1C10Page = writer.GetImportedPage(R1C10File, 1);
                var R1C10PDF = writer.GetImportedPage(R1C10File, 1);
                var R1C10 = new System.Drawing.Drawing2D.Matrix();
                R1C10.Translate(693f, 9f);
                R1C10.Rotate(90);
                writer.DirectContent.AddTemplate(R1C10Page, R1C10);

                PdfReader R1C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R1C11Page = writer.GetImportedPage(R1C11File, 1);
                var R1C11PDF = writer.GetImportedPage(R1C11File, 1);
                var R1C11 = new System.Drawing.Drawing2D.Matrix();
                R1C11.Translate(747f, 9f);
                R1C11.Rotate(90);
                writer.DirectContent.AddTemplate(R1C11Page, R1C11);

                PdfReader R1C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R1C12Page = writer.GetImportedPage(R1C12File, 1);
                var R1C12PDF = writer.GetImportedPage(R1C12File, 1);
                var R1C12 = new System.Drawing.Drawing2D.Matrix();
                R1C12.Translate(801f, 9f);
                R1C12.Rotate(90);
                writer.DirectContent.AddTemplate(R1C12Page, R1C12);


                //Row 2
                PdfReader R2C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(153f, 171f);
                R2C1.Rotate(90);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(207f, 171f);
                R2C2.Rotate(90);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(261f, 171f);
                R2C3.Rotate(90);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                PdfReader R2C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R2C4Page = writer.GetImportedPage(R2C4File, 1);
                var R2C4PDF = writer.GetImportedPage(R2C4File, 1);
                var R2C4 = new System.Drawing.Drawing2D.Matrix();
                R2C4.Translate(333f, 171f);
                R2C4.Rotate(90);
                writer.DirectContent.AddTemplate(R2C4Page, R2C4);

                PdfReader R2C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R2C5Page = writer.GetImportedPage(R2C5File, 1);
                var R2C5PDF = writer.GetImportedPage(R2C5File, 1);
                var R2C5 = new System.Drawing.Drawing2D.Matrix();
                R2C5.Translate(387f, 171f);
                R2C5.Rotate(90);
                writer.DirectContent.AddTemplate(R2C5Page, R2C5);

                PdfReader R2C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R2C6Page = writer.GetImportedPage(R2C6File, 1);
                var R2C6PDF = writer.GetImportedPage(R2C6File, 1);
                var R2C6 = new System.Drawing.Drawing2D.Matrix();
                R2C6.Translate(441f, 171f);
                R2C6.Rotate(90);
                writer.DirectContent.AddTemplate(R2C6Page, R2C6);

                PdfReader R2C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R2C7Page = writer.GetImportedPage(R2C7File, 1);
                var R2C7PDF = writer.GetImportedPage(R2C7File, 1);
                var R2C7 = new System.Drawing.Drawing2D.Matrix();
                R2C7.Translate(513f, 171f);
                R2C7.Rotate(90);
                writer.DirectContent.AddTemplate(R2C7Page, R2C7);

                PdfReader R2C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R2C8Page = writer.GetImportedPage(R2C8File, 1);
                var R2C8PDF = writer.GetImportedPage(R2C8File, 1);
                var R2C8 = new System.Drawing.Drawing2D.Matrix();
                R2C8.Translate(567f, 171f);
                R2C8.Rotate(90);
                writer.DirectContent.AddTemplate(R2C8Page, R2C8);

                PdfReader R2C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R2C9Page = writer.GetImportedPage(R2C9File, 1);
                var R2C9PDF = writer.GetImportedPage(R2C9File, 1);
                var R2C9 = new System.Drawing.Drawing2D.Matrix();
                R2C9.Translate(621f, 171f);
                R2C9.Rotate(90);
                writer.DirectContent.AddTemplate(R2C9Page, R2C9);

                PdfReader R2C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R2C10Page = writer.GetImportedPage(R2C10File, 1);
                var R2C10PDF = writer.GetImportedPage(R2C10File, 1);
                var R2C10 = new System.Drawing.Drawing2D.Matrix();
                R2C10.Translate(693f, 171f);
                R2C10.Rotate(90);
                writer.DirectContent.AddTemplate(R2C10Page, R2C10);

                PdfReader R2C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R2C11Page = writer.GetImportedPage(R2C11File, 1);
                var R2C11PDF = writer.GetImportedPage(R2C11File, 1);
                var R2C11 = new System.Drawing.Drawing2D.Matrix();
                R2C11.Translate(747f, 171f);
                R2C11.Rotate(90);
                writer.DirectContent.AddTemplate(R2C11Page, R2C11);

                PdfReader R2C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R2C12Page = writer.GetImportedPage(R2C12File, 1);
                var R2C12PDF = writer.GetImportedPage(R2C12File, 1);
                var R2C12 = new System.Drawing.Drawing2D.Matrix();
                R2C12.Translate(801f, 171f);
                R2C12.Rotate(90);
                writer.DirectContent.AddTemplate(R2C12Page, R2C12);


                //Row 3
                PdfReader R3C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(153f, 333f);
                R3C1.Rotate(90);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(207f, 333f);
                R3C2.Rotate(90);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(261f, 333f);
                R3C3.Rotate(90);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                PdfReader R3C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R3C4Page = writer.GetImportedPage(R3C4File, 1);
                var R3C4PDF = writer.GetImportedPage(R3C4File, 1);
                var R3C4 = new System.Drawing.Drawing2D.Matrix();
                R3C4.Translate(333f, 333f);
                R3C4.Rotate(90);
                writer.DirectContent.AddTemplate(R3C4Page, R3C4);

                PdfReader R3C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R3C5Page = writer.GetImportedPage(R3C5File, 1);
                var R3C5PDF = writer.GetImportedPage(R3C5File, 1);
                var R3C5 = new System.Drawing.Drawing2D.Matrix();
                R3C5.Translate(387f, 333f);
                R3C5.Rotate(90);
                writer.DirectContent.AddTemplate(R3C5Page, R3C5);

                PdfReader R3C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R3C6Page = writer.GetImportedPage(R3C6File, 1);
                var R3C6PDF = writer.GetImportedPage(R3C6File, 1);
                var R3C6 = new System.Drawing.Drawing2D.Matrix();
                R3C6.Translate(441f, 333f);
                R3C6.Rotate(90);
                writer.DirectContent.AddTemplate(R3C6Page, R3C6);

                PdfReader R3C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R3C7Page = writer.GetImportedPage(R3C7File, 1);
                var R3C7PDF = writer.GetImportedPage(R3C7File, 1);
                var R3C7 = new System.Drawing.Drawing2D.Matrix();
                R3C7.Translate(513f, 333f);
                R3C7.Rotate(90);
                writer.DirectContent.AddTemplate(R3C7Page, R3C7);

                PdfReader R3C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R3C8Page = writer.GetImportedPage(R3C8File, 1);
                var R3C8PDF = writer.GetImportedPage(R3C8File, 1);
                var R3C8 = new System.Drawing.Drawing2D.Matrix();
                R3C8.Translate(567f, 333f);
                R3C8.Rotate(90);
                writer.DirectContent.AddTemplate(R3C8Page, R3C8);

                PdfReader R3C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R3C9Page = writer.GetImportedPage(R3C9File, 1);
                var R3C9PDF = writer.GetImportedPage(R3C9File, 1);
                var R3C9 = new System.Drawing.Drawing2D.Matrix();
                R3C9.Translate(621f, 333f);
                R3C9.Rotate(90);
                writer.DirectContent.AddTemplate(R3C9Page, R3C9);

                PdfReader R3C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R3C10Page = writer.GetImportedPage(R3C10File, 1);
                var R3C10PDF = writer.GetImportedPage(R3C10File, 1);
                var R3C10 = new System.Drawing.Drawing2D.Matrix();
                R3C10.Translate(693f, 333f);
                R3C10.Rotate(90);
                writer.DirectContent.AddTemplate(R3C10Page, R3C10);

                PdfReader R3C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R3C11Page = writer.GetImportedPage(R3C11File, 1);
                var R3C11PDF = writer.GetImportedPage(R3C11File, 1);
                var R3C11 = new System.Drawing.Drawing2D.Matrix();
                R3C11.Translate(747f, 333f);
                R3C11.Rotate(90);
                writer.DirectContent.AddTemplate(R3C11Page, R3C11);

                PdfReader R3C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R3C12Page = writer.GetImportedPage(R3C12File, 1);
                var R3C12PDF = writer.GetImportedPage(R3C12File, 1);
                var R3C12 = new System.Drawing.Drawing2D.Matrix();
                R3C12.Translate(801f, 333f);
                R3C12.Rotate(90);
                writer.DirectContent.AddTemplate(R3C12Page, R3C12);


                //Row 4
                PdfReader R4C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(153f, 495f);
                R4C1.Rotate(90);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(207f, 495f);
                R4C2.Rotate(90);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(261f, 495f);
                R4C3.Rotate(90);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);

                PdfReader R4C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R4C4Page = writer.GetImportedPage(R4C4File, 1);
                var R4C4PDF = writer.GetImportedPage(R4C4File, 1);
                var R4C4 = new System.Drawing.Drawing2D.Matrix();
                R4C4.Translate(333f, 495f);
                R4C4.Rotate(90);
                writer.DirectContent.AddTemplate(R4C4Page, R4C4);

                PdfReader R4C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R4C5Page = writer.GetImportedPage(R4C5File, 1);
                var R4C5PDF = writer.GetImportedPage(R4C5File, 1);
                var R4C5 = new System.Drawing.Drawing2D.Matrix();
                R4C5.Translate(387f, 495f);
                R4C5.Rotate(90);
                writer.DirectContent.AddTemplate(R4C5Page, R4C5);

                PdfReader R4C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R4C6Page = writer.GetImportedPage(R4C6File, 1);
                var R4C6PDF = writer.GetImportedPage(R4C6File, 1);
                var R4C6 = new System.Drawing.Drawing2D.Matrix();
                R4C6.Translate(441f, 495f);
                R4C6.Rotate(90);
                writer.DirectContent.AddTemplate(R4C6Page, R4C6);

                PdfReader R4C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R4C7Page = writer.GetImportedPage(R4C7File, 1);
                var R4C7PDF = writer.GetImportedPage(R4C7File, 1);
                var R4C7 = new System.Drawing.Drawing2D.Matrix();
                R4C7.Translate(513f, 495f);
                R4C7.Rotate(90);
                writer.DirectContent.AddTemplate(R4C7Page, R4C7);

                PdfReader R4C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R4C8Page = writer.GetImportedPage(R4C8File, 1);
                var R4C8PDF = writer.GetImportedPage(R4C8File, 1);
                var R4C8 = new System.Drawing.Drawing2D.Matrix();
                R4C8.Translate(567f, 495f);
                R4C8.Rotate(90);
                writer.DirectContent.AddTemplate(R4C8Page, R4C8);

                PdfReader R4C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R4C9Page = writer.GetImportedPage(R4C9File, 1);
                var R4C9PDF = writer.GetImportedPage(R4C9File, 1);
                var R4C9 = new System.Drawing.Drawing2D.Matrix();
                R4C9.Translate(621f, 495f);
                R4C9.Rotate(90);
                writer.DirectContent.AddTemplate(R4C9Page, R4C9);

                PdfReader R4C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R4C10Page = writer.GetImportedPage(R4C10File, 1);
                var R4C10PDF = writer.GetImportedPage(R4C10File, 1);
                var R4C10 = new System.Drawing.Drawing2D.Matrix();
                R4C10.Translate(693f, 495f);
                R4C10.Rotate(90);
                writer.DirectContent.AddTemplate(R4C10Page, R4C10);

                PdfReader R4C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R4C11Page = writer.GetImportedPage(R4C11File, 1);
                var R4C11PDF = writer.GetImportedPage(R4C11File, 1);
                var R4C11 = new System.Drawing.Drawing2D.Matrix();
                R4C11.Translate(747f, 495f);
                R4C11.Rotate(90);
                writer.DirectContent.AddTemplate(R4C11Page, R4C11);

                PdfReader R4C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R4C12Page = writer.GetImportedPage(R4C12File, 1);
                var R4C12PDF = writer.GetImportedPage(R4C12File, 1);
                var R4C12 = new System.Drawing.Drawing2D.Matrix();
                R4C12.Translate(801f, 495f);
                R4C12.Rotate(90);
                writer.DirectContent.AddTemplate(R4C12Page, R4C12);

                //Row 5
                PdfReader R5C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(153f, 657f);
                R5C1.Rotate(90);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(207f, 657f);
                R5C2.Rotate(90);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                PdfReader R5C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R5C3Page = writer.GetImportedPage(R5C3File, 1);
                var R5C3PDF = writer.GetImportedPage(R5C3File, 1);
                var R5C3 = new System.Drawing.Drawing2D.Matrix();
                R5C3.Translate(261f, 657f);
                R5C3.Rotate(90);
                writer.DirectContent.AddTemplate(R5C3Page, R5C3);

                PdfReader R5C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R5C4Page = writer.GetImportedPage(R5C4File, 1);
                var R5C4PDF = writer.GetImportedPage(R5C4File, 1);
                var R5C4 = new System.Drawing.Drawing2D.Matrix();
                R5C4.Translate(333f, 657f);
                R5C4.Rotate(90);
                writer.DirectContent.AddTemplate(R5C4Page, R5C4);

                PdfReader R5C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R5C5Page = writer.GetImportedPage(R5C5File, 1);
                var R5C5PDF = writer.GetImportedPage(R5C5File, 1);
                var R5C5 = new System.Drawing.Drawing2D.Matrix();
                R5C5.Translate(387f, 657f);
                R5C5.Rotate(90);
                writer.DirectContent.AddTemplate(R5C5Page, R5C5);

                PdfReader R5C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R5C6Page = writer.GetImportedPage(R5C6File, 1);
                var R5C6PDF = writer.GetImportedPage(R5C6File, 1);
                var R5C6 = new System.Drawing.Drawing2D.Matrix();
                R5C6.Translate(441f, 657f);
                R5C6.Rotate(90);
                writer.DirectContent.AddTemplate(R5C6Page, R5C6);

                PdfReader R5C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R5C7Page = writer.GetImportedPage(R5C7File, 1);
                var R5C7PDF = writer.GetImportedPage(R5C7File, 1);
                var R5C7 = new System.Drawing.Drawing2D.Matrix();
                R5C7.Translate(513f, 657f);
                R5C7.Rotate(90);
                writer.DirectContent.AddTemplate(R5C7Page, R5C7);

                PdfReader R5C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R5C8Page = writer.GetImportedPage(R5C8File, 1);
                var R5C8PDF = writer.GetImportedPage(R5C8File, 1);
                var R5C8 = new System.Drawing.Drawing2D.Matrix();
                R5C8.Translate(567f, 657f);
                R5C8.Rotate(90);
                writer.DirectContent.AddTemplate(R5C8Page, R5C8);

                PdfReader R5C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R5C9Page = writer.GetImportedPage(R5C9File, 1);
                var R5C9PDF = writer.GetImportedPage(R5C9File, 1);
                var R5C9 = new System.Drawing.Drawing2D.Matrix();
                R5C9.Translate(621f, 657f);
                R5C9.Rotate(90);
                writer.DirectContent.AddTemplate(R5C9Page, R5C9);

                PdfReader R5C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R5C10Page = writer.GetImportedPage(R5C10File, 1);
                var R5C10PDF = writer.GetImportedPage(R5C10File, 1);
                var R5C10 = new System.Drawing.Drawing2D.Matrix();
                R5C10.Translate(693f, 657f);
                R5C10.Rotate(90);
                writer.DirectContent.AddTemplate(R5C10Page, R5C10);

                PdfReader R5C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R5C11Page = writer.GetImportedPage(R5C11File, 1);
                var R5C11PDF = writer.GetImportedPage(R5C11File, 1);
                var R5C11 = new System.Drawing.Drawing2D.Matrix();
                R5C11.Translate(747f, 657f);
                R5C11.Rotate(90);
                writer.DirectContent.AddTemplate(R5C11Page, R5C11);

                PdfReader R5C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R5C12Page = writer.GetImportedPage(R5C12File, 1);
                var R5C12PDF = writer.GetImportedPage(R5C12File, 1);
                var R5C12 = new System.Drawing.Drawing2D.Matrix();
                R5C12.Translate(801f, 657f);
                R5C12.Rotate(90);
                writer.DirectContent.AddTemplate(R5C12Page, R5C12);


                //Row 6
                PdfReader R6C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[12]) + ".pdf");
                PdfImportedPage R6C1Page = writer.GetImportedPage(R6C1File, 1);
                var R6C1PDF = writer.GetImportedPage(R6C1File, 1);
                var R6C1 = new System.Drawing.Drawing2D.Matrix();
                R6C1.Translate(153f, 837f);
                R6C1.Rotate(90);
                writer.DirectContent.AddTemplate(R6C1Page, R6C1);

                PdfReader R6C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[13]) + ".pdf");
                PdfImportedPage R6C2Page = writer.GetImportedPage(R6C2File, 1);
                var R6C2PDF = writer.GetImportedPage(R6C2File, 1);
                var R6C2 = new System.Drawing.Drawing2D.Matrix();
                R6C2.Translate(207f, 837f);
                R6C2.Rotate(90);
                writer.DirectContent.AddTemplate(R6C2Page, R6C2);

                PdfReader R6C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[14]) + ".pdf");
                PdfImportedPage R6C3Page = writer.GetImportedPage(R6C3File, 1);
                var R6C3PDF = writer.GetImportedPage(R6C3File, 1);
                var R6C3 = new System.Drawing.Drawing2D.Matrix();
                R6C3.Translate(261f, 837f);
                R6C3.Rotate(90);
                writer.DirectContent.AddTemplate(R6C3Page, R6C3);

                PdfReader R6C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[15]) + ".pdf");
                PdfImportedPage R6C4Page = writer.GetImportedPage(R6C4File, 1);
                var R6C4PDF = writer.GetImportedPage(R6C4File, 1);
                var R6C4 = new System.Drawing.Drawing2D.Matrix();
                R6C4.Translate(333f, 837f);
                R6C4.Rotate(90);
                writer.DirectContent.AddTemplate(R6C4Page, R6C4);

                PdfReader R6C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[16]) + ".pdf");
                PdfImportedPage R6C5Page = writer.GetImportedPage(R6C5File, 1);
                var R6C5PDF = writer.GetImportedPage(R6C5File, 1);
                var R6C5 = new System.Drawing.Drawing2D.Matrix();
                R6C5.Translate(387f, 837f);
                R6C5.Rotate(90);
                writer.DirectContent.AddTemplate(R6C5Page, R6C5);

                PdfReader R6C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[17]) + ".pdf");
                PdfImportedPage R6C6Page = writer.GetImportedPage(R6C6File, 1);
                var R6C6PDF = writer.GetImportedPage(R6C6File, 1);
                var R6C6 = new System.Drawing.Drawing2D.Matrix();
                R6C6.Translate(441f, 837f);
                R6C6.Rotate(90);
                writer.DirectContent.AddTemplate(R6C6Page, R6C6);

                PdfReader R6C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[18]) + ".pdf");
                PdfImportedPage R6C7Page = writer.GetImportedPage(R6C7File, 1);
                var R6C7PDF = writer.GetImportedPage(R6C7File, 1);
                var R6C7 = new System.Drawing.Drawing2D.Matrix();
                R6C7.Translate(513f, 837f);
                R6C7.Rotate(90);
                writer.DirectContent.AddTemplate(R6C7Page, R6C7);

                PdfReader R6C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[19]) + ".pdf");
                PdfImportedPage R6C8Page = writer.GetImportedPage(R6C8File, 1);
                var R6C8PDF = writer.GetImportedPage(R6C8File, 1);
                var R6C8 = new System.Drawing.Drawing2D.Matrix();
                R6C8.Translate(567f, 837f);
                R6C8.Rotate(90);
                writer.DirectContent.AddTemplate(R6C8Page, R6C8);

                PdfReader R6C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[20]) + ".pdf");
                PdfImportedPage R6C9Page = writer.GetImportedPage(R6C9File, 1);
                var R6C9PDF = writer.GetImportedPage(R6C9File, 1);
                var R6C9 = new System.Drawing.Drawing2D.Matrix();
                R6C9.Translate(621f, 837f);
                R6C9.Rotate(90);
                writer.DirectContent.AddTemplate(R6C9Page, R6C9);

                PdfReader R6C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[21]) + ".pdf");
                PdfImportedPage R6C10Page = writer.GetImportedPage(R6C10File, 1);
                var R6C10PDF = writer.GetImportedPage(R6C10File, 1);
                var R6C10 = new System.Drawing.Drawing2D.Matrix();
                R6C10.Translate(693f, 837f);
                R6C10.Rotate(90);
                writer.DirectContent.AddTemplate(R6C10Page, R6C10);

                PdfReader R6C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[22]) + ".pdf");
                PdfImportedPage R6C11Page = writer.GetImportedPage(R6C11File, 1);
                var R6C11PDF = writer.GetImportedPage(R6C11File, 1);
                var R6C11 = new System.Drawing.Drawing2D.Matrix();
                R6C11.Translate(747f, 837f);
                R6C11.Rotate(90);
                writer.DirectContent.AddTemplate(R6C11Page, R6C11);

                PdfReader R6C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[23]) + ".pdf");
                PdfImportedPage R6C12Page = writer.GetImportedPage(R6C12File, 1);
                var R6C12PDF = writer.GetImportedPage(R6C12File, 1);
                var R6C12 = new System.Drawing.Drawing2D.Matrix();
                R6C12.Translate(801f, 837f);
                R6C12.Rotate(90);
                writer.DirectContent.AddTemplate(R6C12Page, R6C12);


                //Row 7
                PdfReader R7C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[12]) + ".pdf");
                PdfImportedPage R7C1Page = writer.GetImportedPage(R7C1File, 1);
                var R7C1PDF = writer.GetImportedPage(R7C1File, 1);
                var R7C1 = new System.Drawing.Drawing2D.Matrix();
                R7C1.Translate(153f, 999f);
                R7C1.Rotate(90);
                writer.DirectContent.AddTemplate(R7C1Page, R7C1);

                PdfReader R7C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[13]) + ".pdf");
                PdfImportedPage R7C2Page = writer.GetImportedPage(R7C2File, 1);
                var R7C2PDF = writer.GetImportedPage(R7C2File, 1);
                var R7C2 = new System.Drawing.Drawing2D.Matrix();
                R7C2.Translate(207f, 999f);
                R7C2.Rotate(90);
                writer.DirectContent.AddTemplate(R7C2Page, R7C2);

                PdfReader R7C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[14]) + ".pdf");
                PdfImportedPage R7C3Page = writer.GetImportedPage(R7C3File, 1);
                var R7C3PDF = writer.GetImportedPage(R7C3File, 1);
                var R7C3 = new System.Drawing.Drawing2D.Matrix();
                R7C3.Translate(261f, 999f);
                R7C3.Rotate(90);
                writer.DirectContent.AddTemplate(R7C3Page, R7C3);

                PdfReader R7C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[15]) + ".pdf");
                PdfImportedPage R7C4Page = writer.GetImportedPage(R7C4File, 1);
                var R7C4PDF = writer.GetImportedPage(R7C4File, 1);
                var R7C4 = new System.Drawing.Drawing2D.Matrix();
                R7C4.Translate(333f, 999f);
                R7C4.Rotate(90);
                writer.DirectContent.AddTemplate(R7C4Page, R7C4);

                PdfReader R7C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[16]) + ".pdf");
                PdfImportedPage R7C5Page = writer.GetImportedPage(R7C5File, 1);
                var R7C5PDF = writer.GetImportedPage(R7C5File, 1);
                var R7C5 = new System.Drawing.Drawing2D.Matrix();
                R7C5.Translate(387f, 999f);
                R7C5.Rotate(90);
                writer.DirectContent.AddTemplate(R7C5Page, R7C5);

                PdfReader R7C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[17]) + ".pdf");
                PdfImportedPage R7C6Page = writer.GetImportedPage(R7C6File, 1);
                var R7C6PDF = writer.GetImportedPage(R7C6File, 1);
                var R7C6 = new System.Drawing.Drawing2D.Matrix();
                R7C6.Translate(441f, 999f);
                R7C6.Rotate(90);
                writer.DirectContent.AddTemplate(R7C6Page, R7C6);

                PdfReader R7C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[18]) + ".pdf");
                PdfImportedPage R7C7Page = writer.GetImportedPage(R7C7File, 1);
                var R7C7PDF = writer.GetImportedPage(R7C7File, 1);
                var R7C7 = new System.Drawing.Drawing2D.Matrix();
                R7C7.Translate(513f, 999f);
                R7C7.Rotate(90);
                writer.DirectContent.AddTemplate(R7C7Page, R7C7);

                PdfReader R7C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[19]) + ".pdf");
                PdfImportedPage R7C8Page = writer.GetImportedPage(R7C8File, 1);
                var R7C8PDF = writer.GetImportedPage(R7C8File, 1);
                var R7C8 = new System.Drawing.Drawing2D.Matrix();
                R7C8.Translate(567f, 999f);
                R7C8.Rotate(90);
                writer.DirectContent.AddTemplate(R7C8Page, R7C8);

                PdfReader R7C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[20]) + ".pdf");
                PdfImportedPage R7C9Page = writer.GetImportedPage(R7C9File, 1);
                var R7C9PDF = writer.GetImportedPage(R7C9File, 1);
                var R7C9 = new System.Drawing.Drawing2D.Matrix();
                R7C9.Translate(621f, 999f);
                R7C9.Rotate(90);
                writer.DirectContent.AddTemplate(R7C9Page, R7C9);

                PdfReader R7C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[21]) + ".pdf");
                PdfImportedPage R7C10Page = writer.GetImportedPage(R7C10File, 1);
                var R7C10PDF = writer.GetImportedPage(R7C10File, 1);
                var R7C10 = new System.Drawing.Drawing2D.Matrix();
                R7C10.Translate(693f, 999f);
                R7C10.Rotate(90);
                writer.DirectContent.AddTemplate(R7C10Page, R7C10);

                PdfReader R7C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[22]) + ".pdf");
                PdfImportedPage R7C11Page = writer.GetImportedPage(R7C11File, 1);
                var R7C11PDF = writer.GetImportedPage(R7C11File, 1);
                var R7C11 = new System.Drawing.Drawing2D.Matrix();
                R7C11.Translate(747f, 999f);
                R7C11.Rotate(90);
                writer.DirectContent.AddTemplate(R7C11Page, R7C11);

                PdfReader R7C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[23]) + ".pdf");
                PdfImportedPage R7C12Page = writer.GetImportedPage(R7C12File, 1);
                var R7C12PDF = writer.GetImportedPage(R7C12File, 1);
                var R7C12 = new System.Drawing.Drawing2D.Matrix();
                R7C12.Translate(801f, 999f);
                R7C12.Rotate(90);
                writer.DirectContent.AddTemplate(R7C12Page, R7C12);


                //Row 8
                PdfReader R8C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[12]) + ".pdf");
                PdfImportedPage R8C1Page = writer.GetImportedPage(R8C1File, 1);
                var R8C1PDF = writer.GetImportedPage(R8C1File, 1);
                var R8C1 = new System.Drawing.Drawing2D.Matrix();
                R8C1.Translate(153f, 1161f);
                R8C1.Rotate(90);
                writer.DirectContent.AddTemplate(R8C1Page, R8C1);

                PdfReader R8C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[13]) + ".pdf");
                PdfImportedPage R8C2Page = writer.GetImportedPage(R8C2File, 1);
                var R8C2PDF = writer.GetImportedPage(R8C2File, 1);
                var R8C2 = new System.Drawing.Drawing2D.Matrix();
                R8C2.Translate(207f, 1161f);
                R8C2.Rotate(90);
                writer.DirectContent.AddTemplate(R8C2Page, R8C2);

                PdfReader R8C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[14]) + ".pdf");
                PdfImportedPage R8C3Page = writer.GetImportedPage(R8C3File, 1);
                var R8C3PDF = writer.GetImportedPage(R8C3File, 1);
                var R8C3 = new System.Drawing.Drawing2D.Matrix();
                R8C3.Translate(261f, 1161f);
                R8C3.Rotate(90);
                writer.DirectContent.AddTemplate(R8C3Page, R8C3);

                PdfReader R8C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[15]) + ".pdf");
                PdfImportedPage R8C4Page = writer.GetImportedPage(R8C4File, 1);
                var R8C4PDF = writer.GetImportedPage(R8C4File, 1);
                var R8C4 = new System.Drawing.Drawing2D.Matrix();
                R8C4.Translate(333f, 1161f);
                R8C4.Rotate(90);
                writer.DirectContent.AddTemplate(R8C4Page, R8C4);

                PdfReader R8C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[16]) + ".pdf");
                PdfImportedPage R8C5Page = writer.GetImportedPage(R8C5File, 1);
                var R8C5PDF = writer.GetImportedPage(R8C5File, 1);
                var R8C5 = new System.Drawing.Drawing2D.Matrix();
                R8C5.Translate(387f, 1161f);
                R8C5.Rotate(90);
                writer.DirectContent.AddTemplate(R8C5Page, R8C5);

                PdfReader R8C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[17]) + ".pdf");
                PdfImportedPage R8C6Page = writer.GetImportedPage(R8C6File, 1);
                var R8C6PDF = writer.GetImportedPage(R8C6File, 1);
                var R8C6 = new System.Drawing.Drawing2D.Matrix();
                R8C6.Translate(441f, 1161f);
                R8C6.Rotate(90);
                writer.DirectContent.AddTemplate(R8C6Page, R8C6);

                PdfReader R8C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[18]) + ".pdf");
                PdfImportedPage R8C7Page = writer.GetImportedPage(R8C7File, 1);
                var R8C7PDF = writer.GetImportedPage(R8C7File, 1);
                var R8C7 = new System.Drawing.Drawing2D.Matrix();
                R8C7.Translate(513f, 1161f);
                R8C7.Rotate(90);
                writer.DirectContent.AddTemplate(R8C7Page, R8C7);

                PdfReader R8C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[19]) + ".pdf");
                PdfImportedPage R8C8Page = writer.GetImportedPage(R8C8File, 1);
                var R8C8PDF = writer.GetImportedPage(R8C8File, 1);
                var R8C8 = new System.Drawing.Drawing2D.Matrix();
                R8C8.Translate(567f, 1161f);
                R8C8.Rotate(90);
                writer.DirectContent.AddTemplate(R8C8Page, R8C8);

                PdfReader R8C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[20]) + ".pdf");
                PdfImportedPage R8C9Page = writer.GetImportedPage(R8C9File, 1);
                var R8C9PDF = writer.GetImportedPage(R8C9File, 1);
                var R8C9 = new System.Drawing.Drawing2D.Matrix();
                R8C9.Translate(621f, 1161f);
                R8C9.Rotate(90);
                writer.DirectContent.AddTemplate(R8C9Page, R8C9);

                PdfReader R8C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[21]) + ".pdf");
                PdfImportedPage R8C10Page = writer.GetImportedPage(R8C10File, 1);
                var R8C10PDF = writer.GetImportedPage(R8C10File, 1);
                var R8C10 = new System.Drawing.Drawing2D.Matrix();
                R8C10.Translate(693f, 1161f);
                R8C10.Rotate(90);
                writer.DirectContent.AddTemplate(R8C10Page, R8C10);

                PdfReader R8C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[22]) + ".pdf");
                PdfImportedPage R8C11Page = writer.GetImportedPage(R8C11File, 1);
                var R8C11PDF = writer.GetImportedPage(R8C11File, 1);
                var R8C11 = new System.Drawing.Drawing2D.Matrix();
                R8C11.Translate(747f, 1161f);
                R8C11.Rotate(90);
                writer.DirectContent.AddTemplate(R8C11Page, R8C11);

                PdfReader R8C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[23]) + ".pdf");
                PdfImportedPage R8C12Page = writer.GetImportedPage(R8C12File, 1);
                var R8C12PDF = writer.GetImportedPage(R8C12File, 1);
                var R8C12 = new System.Drawing.Drawing2D.Matrix();
                R8C12.Translate(801f, 1161f);
                R8C12.Rotate(90);
                writer.DirectContent.AddTemplate(R8C12Page, R8C12);


                //Row 9
                PdfReader R9C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[12]) + ".pdf");
                PdfImportedPage R9C1Page = writer.GetImportedPage(R9C1File, 1);
                var R9C1PDF = writer.GetImportedPage(R9C1File, 1);
                var R9C1 = new System.Drawing.Drawing2D.Matrix();
                R9C1.Translate(153f, 1323f);
                R9C1.Rotate(90);
                writer.DirectContent.AddTemplate(R9C1Page, R9C1);

                PdfReader R9C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[13]) + ".pdf");
                PdfImportedPage R9C2Page = writer.GetImportedPage(R9C2File, 1);
                var R9C2PDF = writer.GetImportedPage(R9C2File, 1);
                var R9C2 = new System.Drawing.Drawing2D.Matrix();
                R9C2.Translate(207f, 1323f);
                R9C2.Rotate(90);
                writer.DirectContent.AddTemplate(R9C2Page, R9C2);

                PdfReader R9C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[14]) + ".pdf");
                PdfImportedPage R9C3Page = writer.GetImportedPage(R9C3File, 1);
                var R9C3PDF = writer.GetImportedPage(R9C3File, 1);
                var R9C3 = new System.Drawing.Drawing2D.Matrix();
                R9C3.Translate(261f, 1323f);
                R9C3.Rotate(90);
                writer.DirectContent.AddTemplate(R9C3Page, R9C3);

                PdfReader R9C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[15]) + ".pdf");
                PdfImportedPage R9C4Page = writer.GetImportedPage(R9C4File, 1);
                var R9C4PDF = writer.GetImportedPage(R9C4File, 1);
                var R9C4 = new System.Drawing.Drawing2D.Matrix();
                R9C4.Translate(333f, 1323f);
                R9C4.Rotate(90);
                writer.DirectContent.AddTemplate(R9C4Page, R9C4);

                PdfReader R9C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[16]) + ".pdf");
                PdfImportedPage R9C5Page = writer.GetImportedPage(R9C5File, 1);
                var R9C5PDF = writer.GetImportedPage(R9C5File, 1);
                var R9C5 = new System.Drawing.Drawing2D.Matrix();
                R9C5.Translate(387f, 1323f);
                R9C5.Rotate(90);
                writer.DirectContent.AddTemplate(R9C5Page, R9C5);

                PdfReader R9C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[17]) + ".pdf");
                PdfImportedPage R9C6Page = writer.GetImportedPage(R9C6File, 1);
                var R9C6PDF = writer.GetImportedPage(R9C6File, 1);
                var R9C6 = new System.Drawing.Drawing2D.Matrix();
                R9C6.Translate(441f, 1323f);
                R9C6.Rotate(90);
                writer.DirectContent.AddTemplate(R9C6Page, R9C6);

                PdfReader R9C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[18]) + ".pdf");
                PdfImportedPage R9C7Page = writer.GetImportedPage(R9C7File, 1);
                var R9C7PDF = writer.GetImportedPage(R9C7File, 1);
                var R9C7 = new System.Drawing.Drawing2D.Matrix();
                R9C7.Translate(513f, 1323f);
                R9C7.Rotate(90);
                writer.DirectContent.AddTemplate(R9C7Page, R9C7);

                PdfReader R9C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[19]) + ".pdf");
                PdfImportedPage R9C8Page = writer.GetImportedPage(R9C8File, 1);
                var R9C8PDF = writer.GetImportedPage(R9C8File, 1);
                var R9C8 = new System.Drawing.Drawing2D.Matrix();
                R9C8.Translate(567f, 1323f);
                R9C8.Rotate(90);
                writer.DirectContent.AddTemplate(R9C8Page, R9C8);

                PdfReader R9C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[20]) + ".pdf");
                PdfImportedPage R9C9Page = writer.GetImportedPage(R9C9File, 1);
                var R9C9PDF = writer.GetImportedPage(R9C9File, 1);
                var R9C9 = new System.Drawing.Drawing2D.Matrix();
                R9C9.Translate(621f, 1323f);
                R9C9.Rotate(90);
                writer.DirectContent.AddTemplate(R9C9Page, R9C9);

                PdfReader R9C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[21]) + ".pdf");
                PdfImportedPage R9C10Page = writer.GetImportedPage(R9C10File, 1);
                var R9C10PDF = writer.GetImportedPage(R9C10File, 1);
                var R9C10 = new System.Drawing.Drawing2D.Matrix();
                R9C10.Translate(693f, 1323f);
                R9C10.Rotate(90);
                writer.DirectContent.AddTemplate(R9C10Page, R9C10);

                PdfReader R9C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[22]) + ".pdf");
                PdfImportedPage R9C11Page = writer.GetImportedPage(R9C11File, 1);
                var R9C11PDF = writer.GetImportedPage(R9C11File, 1);
                var R9C11 = new System.Drawing.Drawing2D.Matrix();
                R9C11.Translate(747f, 1323f);
                R9C11.Rotate(90);
                writer.DirectContent.AddTemplate(R9C11Page, R9C11);

                PdfReader R9C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[23]) + ".pdf");
                PdfImportedPage R9C12Page = writer.GetImportedPage(R9C12File, 1);
                var R9C12PDF = writer.GetImportedPage(R9C12File, 1);
                var R9C12 = new System.Drawing.Drawing2D.Matrix();
                R9C12.Translate(801f, 1323f);
                R9C12.Rotate(90);
                writer.DirectContent.AddTemplate(R9C12Page, R9C12);


                //Row 10
                PdfReader R10C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[12]) + ".pdf");
                PdfImportedPage R10C1Page = writer.GetImportedPage(R10C1File, 1);
                var R10C1PDF = writer.GetImportedPage(R10C1File, 1);
                var R10C1 = new System.Drawing.Drawing2D.Matrix();
                R10C1.Translate(153f, 1485f);
                R10C1.Rotate(90);
                writer.DirectContent.AddTemplate(R10C1Page, R10C1);

                PdfReader R10C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[13]) + ".pdf");
                PdfImportedPage R10C2Page = writer.GetImportedPage(R10C2File, 1);
                var R10C2PDF = writer.GetImportedPage(R10C2File, 1);
                var R10C2 = new System.Drawing.Drawing2D.Matrix();
                R10C2.Translate(207f, 1485f);
                R10C2.Rotate(90);
                writer.DirectContent.AddTemplate(R10C2Page, R10C2);

                PdfReader R10C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[14]) + ".pdf");
                PdfImportedPage R10C3Page = writer.GetImportedPage(R10C3File, 1);
                var R10C3PDF = writer.GetImportedPage(R10C3File, 1);
                var R10C3 = new System.Drawing.Drawing2D.Matrix();
                R10C3.Translate(261f, 1485f);
                R10C3.Rotate(90);
                writer.DirectContent.AddTemplate(R10C3Page, R10C3);

                PdfReader R10C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[15]) + ".pdf");
                PdfImportedPage R10C4Page = writer.GetImportedPage(R10C4File, 1);
                var R10C4PDF = writer.GetImportedPage(R10C4File, 1);
                var R10C4 = new System.Drawing.Drawing2D.Matrix();
                R10C4.Translate(333f, 1485f);
                R10C4.Rotate(90);
                writer.DirectContent.AddTemplate(R10C4Page, R10C4);

                PdfReader R10C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[16]) + ".pdf");
                PdfImportedPage R10C5Page = writer.GetImportedPage(R10C5File, 1);
                var R10C5PDF = writer.GetImportedPage(R10C5File, 1);
                var R10C5 = new System.Drawing.Drawing2D.Matrix();
                R10C5.Translate(387f, 1485f);
                R10C5.Rotate(90);
                writer.DirectContent.AddTemplate(R10C5Page, R10C5);

                PdfReader R10C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[17]) + ".pdf");
                PdfImportedPage R10C6Page = writer.GetImportedPage(R10C6File, 1);
                var R10C6PDF = writer.GetImportedPage(R10C6File, 1);
                var R10C6 = new System.Drawing.Drawing2D.Matrix();
                R10C6.Translate(441f, 1485f);
                R10C6.Rotate(90);
                writer.DirectContent.AddTemplate(R10C6Page, R10C6);

                PdfReader R10C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[18]) + ".pdf");
                PdfImportedPage R10C7Page = writer.GetImportedPage(R10C7File, 1);
                var R10C7PDF = writer.GetImportedPage(R10C7File, 1);
                var R10C7 = new System.Drawing.Drawing2D.Matrix();
                R10C7.Translate(513f, 1485f);
                R10C7.Rotate(90);
                writer.DirectContent.AddTemplate(R10C7Page, R10C7);

                PdfReader R10C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[19]) + ".pdf");
                PdfImportedPage R10C8Page = writer.GetImportedPage(R10C8File, 1);
                var R10C8PDF = writer.GetImportedPage(R10C8File, 1);
                var R10C8 = new System.Drawing.Drawing2D.Matrix();
                R10C8.Translate(567f, 1485f);
                R10C8.Rotate(90);
                writer.DirectContent.AddTemplate(R10C8Page, R10C8);

                PdfReader R10C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[20]) + ".pdf");
                PdfImportedPage R10C9Page = writer.GetImportedPage(R10C9File, 1);
                var R10C9PDF = writer.GetImportedPage(R10C9File, 1);
                var R10C9 = new System.Drawing.Drawing2D.Matrix();
                R10C9.Translate(621f, 1485f);
                R10C9.Rotate(90);
                writer.DirectContent.AddTemplate(R10C9Page, R10C9);

                PdfReader R10C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[21]) + ".pdf");
                PdfImportedPage R10C10Page = writer.GetImportedPage(R10C10File, 1);
                var R10C10PDF = writer.GetImportedPage(R10C10File, 1);
                var R10C10 = new System.Drawing.Drawing2D.Matrix();
                R10C10.Translate(693f, 1485f);
                R10C10.Rotate(90);
                writer.DirectContent.AddTemplate(R10C10Page, R10C10);

                PdfReader R10C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[22]) + ".pdf");
                PdfImportedPage R10C11Page = writer.GetImportedPage(R10C11File, 1);
                var R10C11PDF = writer.GetImportedPage(R10C11File, 1);
                var R10C11 = new System.Drawing.Drawing2D.Matrix();
                R10C11.Translate(747f, 1485f);
                R10C11.Rotate(90);
                writer.DirectContent.AddTemplate(R10C11Page, R10C11);

                PdfReader R10C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[23]) + ".pdf");
                PdfImportedPage R10C12Page = writer.GetImportedPage(R10C12File, 1);
                var R10C12PDF = writer.GetImportedPage(R10C12File, 1);
                var R10C12 = new System.Drawing.Drawing2D.Matrix();
                R10C12.Translate(801f, 1485f);
                R10C12.Rotate(90);
                writer.DirectContent.AddTemplate(R10C12Page, R10C12);

                itemTotal.RemoveRange(0, 24);

                cb.SetLineWidth(18f);

                //Cropmarks Horizontal
                cb.MoveTo(22.5f, 0);
                cb.LineTo(877.5f, 0);
                cb.Stroke();

                cb.MoveTo(22.5f, 828f);
                cb.LineTo(877.5f, 828f);
                cb.Stroke();

                cb.MoveTo(22.5f, 1656f);
                cb.LineTo(877.5f, 1656f);
                cb.Stroke();



                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(40.5f, 0);
                cb.LineTo(859.5f, 0);
                cb.LineTo(859.5f, 2673);
                cb.LineTo(40.5f, 2673);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf2x0_5_Long(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(162f, 54f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 54f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2430));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 4 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[3]);
                        itemPrint.RemoveRange(0, 4);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 45);
                        diffPerPage.Add("4 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 4);

                    }
                    else if (itemPrint.Count() % 2 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemPrint.RemoveRange(0, 2);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 90);
                        diffPerPage.Add("2 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 2);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 135);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();

                //Row 1
                PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(153f, 0);
                R1C1.Rotate(90);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(207f, 0);
                R1C2.Rotate(90);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(261f, 0);
                R1C3.Rotate(90);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                var R1C4 = new System.Drawing.Drawing2D.Matrix();
                R1C4.Translate(333f, 0);
                R1C4.Rotate(90);
                writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                PdfReader R1C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R1C5Page = writer.GetImportedPage(R1C5File, 1);
                var R1C5PDF = writer.GetImportedPage(R1C5File, 1);
                var R1C5 = new System.Drawing.Drawing2D.Matrix();
                R1C5.Translate(387f, 0);
                R1C5.Rotate(90);
                writer.DirectContent.AddTemplate(R1C5Page, R1C5);

                PdfReader R1C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R1C6Page = writer.GetImportedPage(R1C6File, 1);
                var R1C6PDF = writer.GetImportedPage(R1C6File, 1);
                var R1C6 = new System.Drawing.Drawing2D.Matrix();
                R1C6.Translate(441f, 0);
                R1C6.Rotate(90);
                writer.DirectContent.AddTemplate(R1C6Page, R1C6);

                PdfReader R1C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R1C7Page = writer.GetImportedPage(R1C7File, 1);
                var R1C7PDF = writer.GetImportedPage(R1C7File, 1);
                var R1C7 = new System.Drawing.Drawing2D.Matrix();
                R1C7.Translate(513f, 0);
                R1C7.Rotate(90);
                writer.DirectContent.AddTemplate(R1C7Page, R1C7);

                PdfReader R1C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R1C8Page = writer.GetImportedPage(R1C8File, 1);
                var R1C8PDF = writer.GetImportedPage(R1C8File, 1);
                var R1C8 = new System.Drawing.Drawing2D.Matrix();
                R1C8.Translate(567f, 0);
                R1C8.Rotate(90);
                writer.DirectContent.AddTemplate(R1C8Page, R1C8);

                PdfReader R1C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R1C9Page = writer.GetImportedPage(R1C9File, 1);
                var R1C9PDF = writer.GetImportedPage(R1C9File, 1);
                var R1C9 = new System.Drawing.Drawing2D.Matrix();
                R1C9.Translate(621f, 0);
                R1C9.Rotate(90);
                writer.DirectContent.AddTemplate(R1C9Page, R1C9);

                PdfReader R1C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R1C10Page = writer.GetImportedPage(R1C10File, 1);
                var R1C10PDF = writer.GetImportedPage(R1C10File, 1);
                var R1C10 = new System.Drawing.Drawing2D.Matrix();
                R1C10.Translate(693f, 0);
                R1C10.Rotate(90);
                writer.DirectContent.AddTemplate(R1C10Page, R1C10);

                PdfReader R1C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R1C11Page = writer.GetImportedPage(R1C11File, 1);
                var R1C11PDF = writer.GetImportedPage(R1C11File, 1);
                var R1C11 = new System.Drawing.Drawing2D.Matrix();
                R1C11.Translate(747f, 0);
                R1C11.Rotate(90);
                writer.DirectContent.AddTemplate(R1C11Page, R1C11);

                PdfReader R1C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R1C12Page = writer.GetImportedPage(R1C12File, 1);
                var R1C12PDF = writer.GetImportedPage(R1C12File, 1);
                var R1C12 = new System.Drawing.Drawing2D.Matrix();
                R1C12.Translate(801f, 0);
                R1C12.Rotate(90);
                writer.DirectContent.AddTemplate(R1C12Page, R1C12);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(153f, 162);
                R2C1.Rotate(90);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(207f, 162);
                R2C2.Rotate(90);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(261f, 162);
                R2C3.Rotate(90);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                PdfReader R2C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R2C4Page = writer.GetImportedPage(R2C4File, 1);
                var R2C4PDF = writer.GetImportedPage(R2C4File, 1);
                var R2C4 = new System.Drawing.Drawing2D.Matrix();
                R2C4.Translate(333f, 162);
                R2C4.Rotate(90);
                writer.DirectContent.AddTemplate(R2C4Page, R2C4);

                PdfReader R2C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R2C5Page = writer.GetImportedPage(R2C5File, 1);
                var R2C5PDF = writer.GetImportedPage(R2C5File, 1);
                var R2C5 = new System.Drawing.Drawing2D.Matrix();
                R2C5.Translate(387f, 162);
                R2C5.Rotate(90);
                writer.DirectContent.AddTemplate(R2C5Page, R2C5);

                PdfReader R2C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R2C6Page = writer.GetImportedPage(R2C6File, 1);
                var R2C6PDF = writer.GetImportedPage(R2C6File, 1);
                var R2C6 = new System.Drawing.Drawing2D.Matrix();
                R2C6.Translate(441f, 162);
                R2C6.Rotate(90);
                writer.DirectContent.AddTemplate(R2C6Page, R2C6);

                PdfReader R2C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R2C7Page = writer.GetImportedPage(R2C7File, 1);
                var R2C7PDF = writer.GetImportedPage(R2C7File, 1);
                var R2C7 = new System.Drawing.Drawing2D.Matrix();
                R2C7.Translate(513f, 162);
                R2C7.Rotate(90);
                writer.DirectContent.AddTemplate(R2C7Page, R2C7);

                PdfReader R2C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R2C8Page = writer.GetImportedPage(R2C8File, 1);
                var R2C8PDF = writer.GetImportedPage(R2C8File, 1);
                var R2C8 = new System.Drawing.Drawing2D.Matrix();
                R2C8.Translate(567f, 162);
                R2C8.Rotate(90);
                writer.DirectContent.AddTemplate(R2C8Page, R2C8);

                PdfReader R2C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R2C9Page = writer.GetImportedPage(R2C9File, 1);
                var R2C9PDF = writer.GetImportedPage(R2C9File, 1);
                var R2C9 = new System.Drawing.Drawing2D.Matrix();
                R2C9.Translate(621f, 162);
                R2C9.Rotate(90);
                writer.DirectContent.AddTemplate(R2C9Page, R2C9);

                PdfReader R2C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R2C10Page = writer.GetImportedPage(R2C10File, 1);
                var R2C10PDF = writer.GetImportedPage(R2C10File, 1);
                var R2C10 = new System.Drawing.Drawing2D.Matrix();
                R2C10.Translate(693f, 162);
                R2C10.Rotate(90);
                writer.DirectContent.AddTemplate(R2C10Page, R2C10);

                PdfReader R2C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R2C11Page = writer.GetImportedPage(R2C11File, 1);
                var R2C11PDF = writer.GetImportedPage(R2C11File, 1);
                var R2C11 = new System.Drawing.Drawing2D.Matrix();
                R2C11.Translate(747f, 162);
                R2C11.Rotate(90);
                writer.DirectContent.AddTemplate(R2C11Page, R2C11);

                PdfReader R2C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R2C12Page = writer.GetImportedPage(R2C12File, 1);
                var R2C12PDF = writer.GetImportedPage(R2C12File, 1);
                var R2C12 = new System.Drawing.Drawing2D.Matrix();
                R2C12.Translate(801f, 162);
                R2C12.Rotate(90);
                writer.DirectContent.AddTemplate(R2C12Page, R2C12);


                //Row 3
                PdfReader R3C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(153f, 324);
                R3C1.Rotate(90);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(207f, 324);
                R3C2.Rotate(90);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(261f, 324);
                R3C3.Rotate(90);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                PdfReader R3C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R3C4Page = writer.GetImportedPage(R3C4File, 1);
                var R3C4PDF = writer.GetImportedPage(R3C4File, 1);
                var R3C4 = new System.Drawing.Drawing2D.Matrix();
                R3C4.Translate(333f, 324);
                R3C4.Rotate(90);
                writer.DirectContent.AddTemplate(R3C4Page, R3C4);

                PdfReader R3C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R3C5Page = writer.GetImportedPage(R3C5File, 1);
                var R3C5PDF = writer.GetImportedPage(R3C5File, 1);
                var R3C5 = new System.Drawing.Drawing2D.Matrix();
                R3C5.Translate(387f, 324);
                R3C5.Rotate(90);
                writer.DirectContent.AddTemplate(R3C5Page, R3C5);

                PdfReader R3C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R3C6Page = writer.GetImportedPage(R3C6File, 1);
                var R3C6PDF = writer.GetImportedPage(R3C6File, 1);
                var R3C6 = new System.Drawing.Drawing2D.Matrix();
                R3C6.Translate(441f, 324);
                R3C6.Rotate(90);
                writer.DirectContent.AddTemplate(R3C6Page, R3C6);

                PdfReader R3C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R3C7Page = writer.GetImportedPage(R3C7File, 1);
                var R3C7PDF = writer.GetImportedPage(R3C7File, 1);
                var R3C7 = new System.Drawing.Drawing2D.Matrix();
                R3C7.Translate(513f, 324);
                R3C7.Rotate(90);
                writer.DirectContent.AddTemplate(R3C7Page, R3C7);

                PdfReader R3C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R3C8Page = writer.GetImportedPage(R3C8File, 1);
                var R3C8PDF = writer.GetImportedPage(R3C8File, 1);
                var R3C8 = new System.Drawing.Drawing2D.Matrix();
                R3C8.Translate(567f, 324);
                R3C8.Rotate(90);
                writer.DirectContent.AddTemplate(R3C8Page, R3C8);

                PdfReader R3C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R3C9Page = writer.GetImportedPage(R3C9File, 1);
                var R3C9PDF = writer.GetImportedPage(R3C9File, 1);
                var R3C9 = new System.Drawing.Drawing2D.Matrix();
                R3C9.Translate(621f, 324);
                R3C9.Rotate(90);
                writer.DirectContent.AddTemplate(R3C9Page, R3C9);

                PdfReader R3C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R3C10Page = writer.GetImportedPage(R3C10File, 1);
                var R3C10PDF = writer.GetImportedPage(R3C10File, 1);
                var R3C10 = new System.Drawing.Drawing2D.Matrix();
                R3C10.Translate(693f, 324);
                R3C10.Rotate(90);
                writer.DirectContent.AddTemplate(R3C10Page, R3C10);

                PdfReader R3C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R3C11Page = writer.GetImportedPage(R3C11File, 1);
                var R3C11PDF = writer.GetImportedPage(R3C11File, 1);
                var R3C11 = new System.Drawing.Drawing2D.Matrix();
                R3C11.Translate(747f, 324);
                R3C11.Rotate(90);
                writer.DirectContent.AddTemplate(R3C11Page, R3C11);

                PdfReader R3C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R3C12Page = writer.GetImportedPage(R3C12File, 1);
                var R3C12PDF = writer.GetImportedPage(R3C12File, 1);
                var R3C12 = new System.Drawing.Drawing2D.Matrix();
                R3C12.Translate(801f, 324);
                R3C12.Rotate(90);
                writer.DirectContent.AddTemplate(R3C12Page, R3C12);


                //Row 4
                PdfReader R4C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(153f, 486);
                R4C1.Rotate(90);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(207f, 486);
                R4C2.Rotate(90);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(261f, 486);
                R4C3.Rotate(90);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);

                PdfReader R4C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R4C4Page = writer.GetImportedPage(R4C4File, 1);
                var R4C4PDF = writer.GetImportedPage(R4C4File, 1);
                var R4C4 = new System.Drawing.Drawing2D.Matrix();
                R4C4.Translate(333f, 486);
                R4C4.Rotate(90);
                writer.DirectContent.AddTemplate(R4C4Page, R4C4);

                PdfReader R4C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R4C5Page = writer.GetImportedPage(R4C5File, 1);
                var R4C5PDF = writer.GetImportedPage(R4C5File, 1);
                var R4C5 = new System.Drawing.Drawing2D.Matrix();
                R4C5.Translate(387f, 486);
                R4C5.Rotate(90);
                writer.DirectContent.AddTemplate(R4C5Page, R4C5);

                PdfReader R4C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R4C6Page = writer.GetImportedPage(R4C6File, 1);
                var R4C6PDF = writer.GetImportedPage(R4C6File, 1);
                var R4C6 = new System.Drawing.Drawing2D.Matrix();
                R4C6.Translate(441f, 486);
                R4C6.Rotate(90);
                writer.DirectContent.AddTemplate(R4C6Page, R4C6);

                PdfReader R4C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R4C7Page = writer.GetImportedPage(R4C7File, 1);
                var R4C7PDF = writer.GetImportedPage(R4C7File, 1);
                var R4C7 = new System.Drawing.Drawing2D.Matrix();
                R4C7.Translate(513f, 486);
                R4C7.Rotate(90);
                writer.DirectContent.AddTemplate(R4C7Page, R4C7);

                PdfReader R4C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R4C8Page = writer.GetImportedPage(R4C8File, 1);
                var R4C8PDF = writer.GetImportedPage(R4C8File, 1);
                var R4C8 = new System.Drawing.Drawing2D.Matrix();
                R4C8.Translate(567f, 486);
                R4C8.Rotate(90);
                writer.DirectContent.AddTemplate(R4C8Page, R4C8);

                PdfReader R4C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R4C9Page = writer.GetImportedPage(R4C9File, 1);
                var R4C9PDF = writer.GetImportedPage(R4C9File, 1);
                var R4C9 = new System.Drawing.Drawing2D.Matrix();
                R4C9.Translate(621f, 486);
                R4C9.Rotate(90);
                writer.DirectContent.AddTemplate(R4C9Page, R4C9);

                PdfReader R4C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R4C10Page = writer.GetImportedPage(R4C10File, 1);
                var R4C10PDF = writer.GetImportedPage(R4C10File, 1);
                var R4C10 = new System.Drawing.Drawing2D.Matrix();
                R4C10.Translate(693f, 486);
                R4C10.Rotate(90);
                writer.DirectContent.AddTemplate(R4C10Page, R4C10);

                PdfReader R4C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R4C11Page = writer.GetImportedPage(R4C11File, 1);
                var R4C11PDF = writer.GetImportedPage(R4C11File, 1);
                var R4C11 = new System.Drawing.Drawing2D.Matrix();
                R4C11.Translate(747f, 486);
                R4C11.Rotate(90);
                writer.DirectContent.AddTemplate(R4C11Page, R4C11);

                PdfReader R4C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R4C12Page = writer.GetImportedPage(R4C12File, 1);
                var R4C12PDF = writer.GetImportedPage(R4C12File, 1);
                var R4C12 = new System.Drawing.Drawing2D.Matrix();
                R4C12.Translate(801f, 486);
                R4C12.Rotate(90);
                writer.DirectContent.AddTemplate(R4C12Page, R4C12);

                //Row 5
                PdfReader R5C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(153f, 648);
                R5C1.Rotate(90);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(207f, 648);
                R5C2.Rotate(90);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                PdfReader R5C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R5C3Page = writer.GetImportedPage(R5C3File, 1);
                var R5C3PDF = writer.GetImportedPage(R5C3File, 1);
                var R5C3 = new System.Drawing.Drawing2D.Matrix();
                R5C3.Translate(261f, 648);
                R5C3.Rotate(90);
                writer.DirectContent.AddTemplate(R5C3Page, R5C3);

                PdfReader R5C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R5C4Page = writer.GetImportedPage(R5C4File, 1);
                var R5C4PDF = writer.GetImportedPage(R5C4File, 1);
                var R5C4 = new System.Drawing.Drawing2D.Matrix();
                R5C4.Translate(333f, 648);
                R5C4.Rotate(90);
                writer.DirectContent.AddTemplate(R5C4Page, R5C4);

                PdfReader R5C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R5C5Page = writer.GetImportedPage(R5C5File, 1);
                var R5C5PDF = writer.GetImportedPage(R5C5File, 1);
                var R5C5 = new System.Drawing.Drawing2D.Matrix();
                R5C5.Translate(387f, 648);
                R5C5.Rotate(90);
                writer.DirectContent.AddTemplate(R5C5Page, R5C5);

                PdfReader R5C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R5C6Page = writer.GetImportedPage(R5C6File, 1);
                var R5C6PDF = writer.GetImportedPage(R5C6File, 1);
                var R5C6 = new System.Drawing.Drawing2D.Matrix();
                R5C6.Translate(441f, 648);
                R5C6.Rotate(90);
                writer.DirectContent.AddTemplate(R5C6Page, R5C6);

                PdfReader R5C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R5C7Page = writer.GetImportedPage(R5C7File, 1);
                var R5C7PDF = writer.GetImportedPage(R5C7File, 1);
                var R5C7 = new System.Drawing.Drawing2D.Matrix();
                R5C7.Translate(513f, 648);
                R5C7.Rotate(90);
                writer.DirectContent.AddTemplate(R5C7Page, R5C7);

                PdfReader R5C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R5C8Page = writer.GetImportedPage(R5C8File, 1);
                var R5C8PDF = writer.GetImportedPage(R5C8File, 1);
                var R5C8 = new System.Drawing.Drawing2D.Matrix();
                R5C8.Translate(567f, 648);
                R5C8.Rotate(90);
                writer.DirectContent.AddTemplate(R5C8Page, R5C8);

                PdfReader R5C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R5C9Page = writer.GetImportedPage(R5C9File, 1);
                var R5C9PDF = writer.GetImportedPage(R5C9File, 1);
                var R5C9 = new System.Drawing.Drawing2D.Matrix();
                R5C9.Translate(621f, 648);
                R5C9.Rotate(90);
                writer.DirectContent.AddTemplate(R5C9Page, R5C9);

                PdfReader R5C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R5C10Page = writer.GetImportedPage(R5C10File, 1);
                var R5C10PDF = writer.GetImportedPage(R5C10File, 1);
                var R5C10 = new System.Drawing.Drawing2D.Matrix();
                R5C10.Translate(693f, 648);
                R5C10.Rotate(90);
                writer.DirectContent.AddTemplate(R5C10Page, R5C10);

                PdfReader R5C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R5C11Page = writer.GetImportedPage(R5C11File, 1);
                var R5C11PDF = writer.GetImportedPage(R5C11File, 1);
                var R5C11 = new System.Drawing.Drawing2D.Matrix();
                R5C11.Translate(747f, 648);
                R5C11.Rotate(90);
                writer.DirectContent.AddTemplate(R5C11Page, R5C11);

                PdfReader R5C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R5C12Page = writer.GetImportedPage(R5C12File, 1);
                var R5C12PDF = writer.GetImportedPage(R5C12File, 1);
                var R5C12 = new System.Drawing.Drawing2D.Matrix();
                R5C12.Translate(801f, 648);
                R5C12.Rotate(90);
                writer.DirectContent.AddTemplate(R5C12Page, R5C12);

                //Row 6
                PdfReader R6C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R6C1Page = writer.GetImportedPage(R6C1File, 1);
                var R6C1PDF = writer.GetImportedPage(R6C1File, 1);
                var R6C1 = new System.Drawing.Drawing2D.Matrix();
                R6C1.Translate(153f, 810);
                R6C1.Rotate(90);
                writer.DirectContent.AddTemplate(R6C1Page, R6C1);

                PdfReader R6C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R6C2Page = writer.GetImportedPage(R6C2File, 1);
                var R6C2PDF = writer.GetImportedPage(R6C2File, 1);
                var R6C2 = new System.Drawing.Drawing2D.Matrix();
                R6C2.Translate(207f, 810);
                R6C2.Rotate(90);
                writer.DirectContent.AddTemplate(R6C2Page, R6C2);

                PdfReader R6C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R6C3Page = writer.GetImportedPage(R6C3File, 1);
                var R6C3PDF = writer.GetImportedPage(R6C3File, 1);
                var R6C3 = new System.Drawing.Drawing2D.Matrix();
                R6C3.Translate(261f, 810);
                R6C3.Rotate(90);
                writer.DirectContent.AddTemplate(R6C3Page, R6C3);

                PdfReader R6C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R6C4Page = writer.GetImportedPage(R6C4File, 1);
                var R6C4PDF = writer.GetImportedPage(R6C4File, 1);
                var R6C4 = new System.Drawing.Drawing2D.Matrix();
                R6C4.Translate(333f, 810);
                R6C4.Rotate(90);
                writer.DirectContent.AddTemplate(R6C4Page, R6C4);

                PdfReader R6C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R6C5Page = writer.GetImportedPage(R6C5File, 1);
                var R6C5PDF = writer.GetImportedPage(R6C5File, 1);
                var R6C5 = new System.Drawing.Drawing2D.Matrix();
                R6C5.Translate(387f, 810);
                R6C5.Rotate(90);
                writer.DirectContent.AddTemplate(R6C5Page, R6C5);

                PdfReader R6C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R6C6Page = writer.GetImportedPage(R6C6File, 1);
                var R6C6PDF = writer.GetImportedPage(R6C6File, 1);
                var R6C6 = new System.Drawing.Drawing2D.Matrix();
                R6C6.Translate(441f, 810);
                R6C6.Rotate(90);
                writer.DirectContent.AddTemplate(R6C6Page, R6C6);

                PdfReader R6C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R6C7Page = writer.GetImportedPage(R6C7File, 1);
                var R6C7PDF = writer.GetImportedPage(R6C7File, 1);
                var R6C7 = new System.Drawing.Drawing2D.Matrix();
                R6C7.Translate(513f, 810);
                R6C7.Rotate(90);
                writer.DirectContent.AddTemplate(R6C7Page, R6C7);

                PdfReader R6C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R6C8Page = writer.GetImportedPage(R6C8File, 1);
                var R6C8PDF = writer.GetImportedPage(R6C8File, 1);
                var R6C8 = new System.Drawing.Drawing2D.Matrix();
                R6C8.Translate(567f, 810);
                R6C8.Rotate(90);
                writer.DirectContent.AddTemplate(R6C8Page, R6C8);

                PdfReader R6C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R6C9Page = writer.GetImportedPage(R6C9File, 1);
                var R6C9PDF = writer.GetImportedPage(R6C9File, 1);
                var R6C9 = new System.Drawing.Drawing2D.Matrix();
                R6C9.Translate(621f, 810);
                R6C9.Rotate(90);
                writer.DirectContent.AddTemplate(R6C9Page, R6C9);

                PdfReader R6C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R6C10Page = writer.GetImportedPage(R6C10File, 1);
                var R6C10PDF = writer.GetImportedPage(R6C10File, 1);
                var R6C10 = new System.Drawing.Drawing2D.Matrix();
                R6C10.Translate(693f, 810);
                R6C10.Rotate(90);
                writer.DirectContent.AddTemplate(R6C10Page, R6C10);

                PdfReader R6C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R6C11Page = writer.GetImportedPage(R6C11File, 1);
                var R6C11PDF = writer.GetImportedPage(R6C11File, 1);
                var R6C11 = new System.Drawing.Drawing2D.Matrix();
                R6C11.Translate(747f, 810);
                R6C11.Rotate(90);
                writer.DirectContent.AddTemplate(R6C11Page, R6C11);

                PdfReader R6C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R6C12Page = writer.GetImportedPage(R6C12File, 1);
                var R6C12PDF = writer.GetImportedPage(R6C12File, 1);
                var R6C12 = new System.Drawing.Drawing2D.Matrix();
                R6C12.Translate(801f, 810);
                R6C12.Rotate(90);
                writer.DirectContent.AddTemplate(R6C12Page, R6C12);

                //Row 7
                PdfReader R7C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R7C1Page = writer.GetImportedPage(R7C1File, 1);
                var R7C1PDF = writer.GetImportedPage(R7C1File, 1);
                var R7C1 = new System.Drawing.Drawing2D.Matrix();
                R7C1.Translate(153f, 972);
                R7C1.Rotate(90);
                writer.DirectContent.AddTemplate(R7C1Page, R7C1);

                PdfReader R7C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R7C2Page = writer.GetImportedPage(R7C2File, 1);
                var R7C2PDF = writer.GetImportedPage(R7C2File, 1);
                var R7C2 = new System.Drawing.Drawing2D.Matrix();
                R7C2.Translate(207f, 972);
                R7C2.Rotate(90);
                writer.DirectContent.AddTemplate(R7C2Page, R7C2);

                PdfReader R7C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R7C3Page = writer.GetImportedPage(R7C3File, 1);
                var R7C3PDF = writer.GetImportedPage(R7C3File, 1);
                var R7C3 = new System.Drawing.Drawing2D.Matrix();
                R7C3.Translate(261f, 972);
                R7C3.Rotate(90);
                writer.DirectContent.AddTemplate(R7C3Page, R7C3);

                PdfReader R7C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R7C4Page = writer.GetImportedPage(R7C4File, 1);
                var R7C4PDF = writer.GetImportedPage(R7C4File, 1);
                var R7C4 = new System.Drawing.Drawing2D.Matrix();
                R7C4.Translate(333f, 972);
                R7C4.Rotate(90);
                writer.DirectContent.AddTemplate(R7C4Page, R7C4);

                PdfReader R7C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R7C5Page = writer.GetImportedPage(R7C5File, 1);
                var R7C5PDF = writer.GetImportedPage(R7C5File, 1);
                var R7C5 = new System.Drawing.Drawing2D.Matrix();
                R7C5.Translate(387f, 972);
                R7C5.Rotate(90);
                writer.DirectContent.AddTemplate(R7C5Page, R7C5);

                PdfReader R7C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R7C6Page = writer.GetImportedPage(R7C6File, 1);
                var R7C6PDF = writer.GetImportedPage(R7C6File, 1);
                var R7C6 = new System.Drawing.Drawing2D.Matrix();
                R7C6.Translate(441f, 972);
                R7C6.Rotate(90);
                writer.DirectContent.AddTemplate(R7C6Page, R7C6);

                PdfReader R7C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R7C7Page = writer.GetImportedPage(R7C7File, 1);
                var R7C7PDF = writer.GetImportedPage(R7C7File, 1);
                var R7C7 = new System.Drawing.Drawing2D.Matrix();
                R7C7.Translate(513f, 972);
                R7C7.Rotate(90);
                writer.DirectContent.AddTemplate(R7C7Page, R7C7);

                PdfReader R7C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R7C8Page = writer.GetImportedPage(R7C8File, 1);
                var R7C8PDF = writer.GetImportedPage(R7C8File, 1);
                var R7C8 = new System.Drawing.Drawing2D.Matrix();
                R7C8.Translate(567f, 972);
                R7C8.Rotate(90);
                writer.DirectContent.AddTemplate(R7C8Page, R7C8);

                PdfReader R7C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R7C9Page = writer.GetImportedPage(R7C9File, 1);
                var R7C9PDF = writer.GetImportedPage(R7C9File, 1);
                var R7C9 = new System.Drawing.Drawing2D.Matrix();
                R7C9.Translate(621f, 972);
                R7C9.Rotate(90);
                writer.DirectContent.AddTemplate(R7C9Page, R7C9);

                PdfReader R7C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R7C10Page = writer.GetImportedPage(R7C10File, 1);
                var R7C10PDF = writer.GetImportedPage(R7C10File, 1);
                var R7C10 = new System.Drawing.Drawing2D.Matrix();
                R7C10.Translate(693f, 972);
                R7C10.Rotate(90);
                writer.DirectContent.AddTemplate(R7C10Page, R7C10);

                PdfReader R7C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R7C11Page = writer.GetImportedPage(R7C11File, 1);
                var R7C11PDF = writer.GetImportedPage(R7C11File, 1);
                var R7C11 = new System.Drawing.Drawing2D.Matrix();
                R7C11.Translate(747f, 972);
                R7C11.Rotate(90);
                writer.DirectContent.AddTemplate(R7C11Page, R7C11);

                PdfReader R7C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R7C12Page = writer.GetImportedPage(R7C12File, 1);
                var R7C12PDF = writer.GetImportedPage(R7C12File, 1);
                var R7C12 = new System.Drawing.Drawing2D.Matrix();
                R7C12.Translate(801f, 972);
                R7C12.Rotate(90);
                writer.DirectContent.AddTemplate(R7C12Page, R7C12);

                //Row 8
                PdfReader R8C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R8C1Page = writer.GetImportedPage(R8C1File, 1);
                var R8C1PDF = writer.GetImportedPage(R8C1File, 1);
                var R8C1 = new System.Drawing.Drawing2D.Matrix();
                R8C1.Translate(153f, 1134);
                R8C1.Rotate(90);
                writer.DirectContent.AddTemplate(R8C1Page, R8C1);

                PdfReader R8C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R8C2Page = writer.GetImportedPage(R8C2File, 1);
                var R8C2PDF = writer.GetImportedPage(R8C2File, 1);
                var R8C2 = new System.Drawing.Drawing2D.Matrix();
                R8C2.Translate(207f, 1134);
                R8C2.Rotate(90);
                writer.DirectContent.AddTemplate(R8C2Page, R8C2);

                PdfReader R8C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R8C3Page = writer.GetImportedPage(R8C3File, 1);
                var R8C3PDF = writer.GetImportedPage(R8C3File, 1);
                var R8C3 = new System.Drawing.Drawing2D.Matrix();
                R8C3.Translate(261f, 1134);
                R8C3.Rotate(90);
                writer.DirectContent.AddTemplate(R8C3Page, R8C3);

                PdfReader R8C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R8C4Page = writer.GetImportedPage(R8C4File, 1);
                var R8C4PDF = writer.GetImportedPage(R8C4File, 1);
                var R8C4 = new System.Drawing.Drawing2D.Matrix();
                R8C4.Translate(333f, 1134);
                R8C4.Rotate(90);
                writer.DirectContent.AddTemplate(R8C4Page, R8C4);

                PdfReader R8C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R8C5Page = writer.GetImportedPage(R8C5File, 1);
                var R8C5PDF = writer.GetImportedPage(R8C5File, 1);
                var R8C5 = new System.Drawing.Drawing2D.Matrix();
                R8C5.Translate(387f, 1134);
                R8C5.Rotate(90);
                writer.DirectContent.AddTemplate(R8C5Page, R8C5);

                PdfReader R8C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R8C6Page = writer.GetImportedPage(R8C6File, 1);
                var R8C6PDF = writer.GetImportedPage(R8C6File, 1);
                var R8C6 = new System.Drawing.Drawing2D.Matrix();
                R8C6.Translate(441f, 1134);
                R8C6.Rotate(90);
                writer.DirectContent.AddTemplate(R8C6Page, R8C6);

                PdfReader R8C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R8C7Page = writer.GetImportedPage(R8C7File, 1);
                var R8C7PDF = writer.GetImportedPage(R8C7File, 1);
                var R8C7 = new System.Drawing.Drawing2D.Matrix();
                R8C7.Translate(513f, 1134);
                R8C7.Rotate(90);
                writer.DirectContent.AddTemplate(R8C7Page, R8C7);

                PdfReader R8C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R8C8Page = writer.GetImportedPage(R8C8File, 1);
                var R8C8PDF = writer.GetImportedPage(R8C8File, 1);
                var R8C8 = new System.Drawing.Drawing2D.Matrix();
                R8C8.Translate(567f, 1134);
                R8C8.Rotate(90);
                writer.DirectContent.AddTemplate(R8C8Page, R8C8);

                PdfReader R8C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R8C9Page = writer.GetImportedPage(R8C9File, 1);
                var R8C9PDF = writer.GetImportedPage(R8C9File, 1);
                var R8C9 = new System.Drawing.Drawing2D.Matrix();
                R8C9.Translate(621f, 1134);
                R8C9.Rotate(90);
                writer.DirectContent.AddTemplate(R8C9Page, R8C9);

                PdfReader R8C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R8C10Page = writer.GetImportedPage(R8C10File, 1);
                var R8C10PDF = writer.GetImportedPage(R8C10File, 1);
                var R8C10 = new System.Drawing.Drawing2D.Matrix();
                R8C10.Translate(693f, 1134);
                R8C10.Rotate(90);
                writer.DirectContent.AddTemplate(R8C10Page, R8C10);

                PdfReader R8C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R8C11Page = writer.GetImportedPage(R8C11File, 1);
                var R8C11PDF = writer.GetImportedPage(R8C11File, 1);
                var R8C11 = new System.Drawing.Drawing2D.Matrix();
                R8C11.Translate(747f, 1134);
                R8C11.Rotate(90);
                writer.DirectContent.AddTemplate(R8C11Page, R8C11);

                PdfReader R8C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R8C12Page = writer.GetImportedPage(R8C12File, 1);
                var R8C12PDF = writer.GetImportedPage(R8C12File, 1);
                var R8C12 = new System.Drawing.Drawing2D.Matrix();
                R8C12.Translate(801f, 1134);
                R8C12.Rotate(90);
                writer.DirectContent.AddTemplate(R8C12Page, R8C12);

                //Row 9
                PdfReader R9C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R9C1Page = writer.GetImportedPage(R9C1File, 1);
                var R9C1PDF = writer.GetImportedPage(R9C1File, 1);
                var R9C1 = new System.Drawing.Drawing2D.Matrix();
                R9C1.Translate(153f, 1296);
                R9C1.Rotate(90);
                writer.DirectContent.AddTemplate(R9C1Page, R9C1);

                PdfReader R9C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R9C2Page = writer.GetImportedPage(R9C2File, 1);
                var R9C2PDF = writer.GetImportedPage(R9C2File, 1);
                var R9C2 = new System.Drawing.Drawing2D.Matrix();
                R9C2.Translate(207f, 1296);
                R9C2.Rotate(90);
                writer.DirectContent.AddTemplate(R9C2Page, R9C2);

                PdfReader R9C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R9C3Page = writer.GetImportedPage(R9C3File, 1);
                var R9C3PDF = writer.GetImportedPage(R9C3File, 1);
                var R9C3 = new System.Drawing.Drawing2D.Matrix();
                R9C3.Translate(261f, 1296);
                R9C3.Rotate(90);
                writer.DirectContent.AddTemplate(R9C3Page, R9C3);

                PdfReader R9C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R9C4Page = writer.GetImportedPage(R9C4File, 1);
                var R9C4PDF = writer.GetImportedPage(R9C4File, 1);
                var R9C4 = new System.Drawing.Drawing2D.Matrix();
                R9C4.Translate(333f, 1296);
                R9C4.Rotate(90);
                writer.DirectContent.AddTemplate(R9C4Page, R9C4);

                PdfReader R9C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R9C5Page = writer.GetImportedPage(R9C5File, 1);
                var R9C5PDF = writer.GetImportedPage(R9C5File, 1);
                var R9C5 = new System.Drawing.Drawing2D.Matrix();
                R9C5.Translate(387f, 1296);
                R9C5.Rotate(90);
                writer.DirectContent.AddTemplate(R9C5Page, R9C5);

                PdfReader R9C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R9C6Page = writer.GetImportedPage(R9C6File, 1);
                var R9C6PDF = writer.GetImportedPage(R9C6File, 1);
                var R9C6 = new System.Drawing.Drawing2D.Matrix();
                R9C6.Translate(441f, 1296);
                R9C6.Rotate(90);
                writer.DirectContent.AddTemplate(R9C6Page, R9C6);

                PdfReader R9C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R9C7Page = writer.GetImportedPage(R9C7File, 1);
                var R9C7PDF = writer.GetImportedPage(R9C7File, 1);
                var R9C7 = new System.Drawing.Drawing2D.Matrix();
                R9C7.Translate(513f, 1296);
                R9C7.Rotate(90);
                writer.DirectContent.AddTemplate(R9C7Page, R9C7);

                PdfReader R9C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R9C8Page = writer.GetImportedPage(R9C8File, 1);
                var R9C8PDF = writer.GetImportedPage(R9C8File, 1);
                var R9C8 = new System.Drawing.Drawing2D.Matrix();
                R9C8.Translate(567f, 1296);
                R9C8.Rotate(90);
                writer.DirectContent.AddTemplate(R9C8Page, R9C8);

                PdfReader R9C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R9C9Page = writer.GetImportedPage(R9C9File, 1);
                var R9C9PDF = writer.GetImportedPage(R9C9File, 1);
                var R9C9 = new System.Drawing.Drawing2D.Matrix();
                R9C9.Translate(621f, 1296);
                R9C9.Rotate(90);
                writer.DirectContent.AddTemplate(R9C9Page, R9C9);

                PdfReader R9C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R9C10Page = writer.GetImportedPage(R9C10File, 1);
                var R9C10PDF = writer.GetImportedPage(R9C10File, 1);
                var R9C10 = new System.Drawing.Drawing2D.Matrix();
                R9C10.Translate(693f, 1296);
                R9C10.Rotate(90);
                writer.DirectContent.AddTemplate(R9C10Page, R9C10);

                PdfReader R9C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R9C11Page = writer.GetImportedPage(R9C11File, 1);
                var R9C11PDF = writer.GetImportedPage(R9C11File, 1);
                var R9C11 = new System.Drawing.Drawing2D.Matrix();
                R9C11.Translate(747f, 1296);
                R9C11.Rotate(90);
                writer.DirectContent.AddTemplate(R9C11Page, R9C11);

                PdfReader R9C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R9C12Page = writer.GetImportedPage(R9C12File, 1);
                var R9C12PDF = writer.GetImportedPage(R9C12File, 1);
                var R9C12 = new System.Drawing.Drawing2D.Matrix();
                R9C12.Translate(801f, 1296);
                R9C12.Rotate(90);
                writer.DirectContent.AddTemplate(R9C12Page, R9C12);

                //Row 10
                PdfReader R10C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R10C1Page = writer.GetImportedPage(R10C1File, 1);
                var R10C1PDF = writer.GetImportedPage(R10C1File, 1);
                var R10C1 = new System.Drawing.Drawing2D.Matrix();
                R10C1.Translate(153f, 1458);
                R10C1.Rotate(90);
                writer.DirectContent.AddTemplate(R10C1Page, R10C1);

                PdfReader R10C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R10C2Page = writer.GetImportedPage(R10C2File, 1);
                var R10C2PDF = writer.GetImportedPage(R10C2File, 1);
                var R10C2 = new System.Drawing.Drawing2D.Matrix();
                R10C2.Translate(207f, 1458);
                R10C2.Rotate(90);
                writer.DirectContent.AddTemplate(R10C2Page, R10C2);

                PdfReader R10C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R10C3Page = writer.GetImportedPage(R10C3File, 1);
                var R10C3PDF = writer.GetImportedPage(R10C3File, 1);
                var R10C3 = new System.Drawing.Drawing2D.Matrix();
                R10C3.Translate(261f, 1458);
                R10C3.Rotate(90);
                writer.DirectContent.AddTemplate(R10C3Page, R10C3);

                PdfReader R10C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R10C4Page = writer.GetImportedPage(R10C4File, 1);
                var R10C4PDF = writer.GetImportedPage(R10C4File, 1);
                var R10C4 = new System.Drawing.Drawing2D.Matrix();
                R10C4.Translate(333f, 1458);
                R10C4.Rotate(90);
                writer.DirectContent.AddTemplate(R10C4Page, R10C4);

                PdfReader R10C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R10C5Page = writer.GetImportedPage(R10C5File, 1);
                var R10C5PDF = writer.GetImportedPage(R10C5File, 1);
                var R10C5 = new System.Drawing.Drawing2D.Matrix();
                R10C5.Translate(387f, 1458);
                R10C5.Rotate(90);
                writer.DirectContent.AddTemplate(R10C5Page, R10C5);

                PdfReader R10C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R10C6Page = writer.GetImportedPage(R10C6File, 1);
                var R10C6PDF = writer.GetImportedPage(R10C6File, 1);
                var R10C6 = new System.Drawing.Drawing2D.Matrix();
                R10C6.Translate(441f, 1458);
                R10C6.Rotate(90);
                writer.DirectContent.AddTemplate(R10C6Page, R10C6);

                PdfReader R10C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R10C7Page = writer.GetImportedPage(R10C7File, 1);
                var R10C7PDF = writer.GetImportedPage(R10C7File, 1);
                var R10C7 = new System.Drawing.Drawing2D.Matrix();
                R10C7.Translate(513f, 1458);
                R10C7.Rotate(90);
                writer.DirectContent.AddTemplate(R10C7Page, R10C7);

                PdfReader R10C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R10C8Page = writer.GetImportedPage(R10C8File, 1);
                var R10C8PDF = writer.GetImportedPage(R10C8File, 1);
                var R10C8 = new System.Drawing.Drawing2D.Matrix();
                R10C8.Translate(567f, 1458);
                R10C8.Rotate(90);
                writer.DirectContent.AddTemplate(R10C8Page, R10C8);

                PdfReader R10C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R10C9Page = writer.GetImportedPage(R10C9File, 1);
                var R10C9PDF = writer.GetImportedPage(R10C9File, 1);
                var R10C9 = new System.Drawing.Drawing2D.Matrix();
                R10C9.Translate(621f, 1458);
                R10C9.Rotate(90);
                writer.DirectContent.AddTemplate(R10C9Page, R10C9);

                PdfReader R10C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R10C10Page = writer.GetImportedPage(R10C10File, 1);
                var R10C10PDF = writer.GetImportedPage(R10C10File, 1);
                var R10C10 = new System.Drawing.Drawing2D.Matrix();
                R10C10.Translate(693f, 1458);
                R10C10.Rotate(90);
                writer.DirectContent.AddTemplate(R10C10Page, R10C10);

                PdfReader R10C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R10C11Page = writer.GetImportedPage(R10C11File, 1);
                var R10C11PDF = writer.GetImportedPage(R10C11File, 1);
                var R10C11 = new System.Drawing.Drawing2D.Matrix();
                R10C11.Translate(747f, 1458);
                R10C11.Rotate(90);
                writer.DirectContent.AddTemplate(R10C11Page, R10C11);

                PdfReader R10C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R10C12Page = writer.GetImportedPage(R10C12File, 1);
                var R10C12PDF = writer.GetImportedPage(R10C12File, 1);
                var R10C12 = new System.Drawing.Drawing2D.Matrix();
                R10C12.Translate(801f, 1458);
                R10C12.Rotate(90);
                writer.DirectContent.AddTemplate(R10C12Page, R10C12);

                //Row 11
                PdfReader R11C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R11C1Page = writer.GetImportedPage(R11C1File, 1);
                var R11C1PDF = writer.GetImportedPage(R11C1File, 1);
                var R11C1 = new System.Drawing.Drawing2D.Matrix();
                R11C1.Translate(153f, 1620);
                R11C1.Rotate(90);
                writer.DirectContent.AddTemplate(R11C1Page, R11C1);

                PdfReader R11C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R11C2Page = writer.GetImportedPage(R11C2File, 1);
                var R11C2PDF = writer.GetImportedPage(R11C2File, 1);
                var R11C2 = new System.Drawing.Drawing2D.Matrix();
                R11C2.Translate(207f, 1620);
                R11C2.Rotate(90);
                writer.DirectContent.AddTemplate(R11C2Page, R11C2);

                PdfReader R11C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R11C3Page = writer.GetImportedPage(R11C3File, 1);
                var R11C3PDF = writer.GetImportedPage(R11C3File, 1);
                var R11C3 = new System.Drawing.Drawing2D.Matrix();
                R11C3.Translate(261f, 1620);
                R11C3.Rotate(90);
                writer.DirectContent.AddTemplate(R11C3Page, R11C3);

                PdfReader R11C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R11C4Page = writer.GetImportedPage(R11C4File, 1);
                var R11C4PDF = writer.GetImportedPage(R11C4File, 1);
                var R11C4 = new System.Drawing.Drawing2D.Matrix();
                R11C4.Translate(333f, 1620);
                R11C4.Rotate(90);
                writer.DirectContent.AddTemplate(R11C4Page, R11C4);

                PdfReader R11C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R11C5Page = writer.GetImportedPage(R11C5File, 1);
                var R11C5PDF = writer.GetImportedPage(R11C5File, 1);
                var R11C5 = new System.Drawing.Drawing2D.Matrix();
                R11C5.Translate(387f, 1620);
                R11C5.Rotate(90);
                writer.DirectContent.AddTemplate(R11C5Page, R11C5);

                PdfReader R11C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R11C6Page = writer.GetImportedPage(R11C6File, 1);
                var R11C6PDF = writer.GetImportedPage(R11C6File, 1);
                var R11C6 = new System.Drawing.Drawing2D.Matrix();
                R11C6.Translate(441f, 1620);
                R11C6.Rotate(90);
                writer.DirectContent.AddTemplate(R11C6Page, R11C6);

                PdfReader R11C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R11C7Page = writer.GetImportedPage(R11C7File, 1);
                var R11C7PDF = writer.GetImportedPage(R11C7File, 1);
                var R11C7 = new System.Drawing.Drawing2D.Matrix();
                R11C7.Translate(513f, 1620);
                R11C7.Rotate(90);
                writer.DirectContent.AddTemplate(R11C7Page, R11C7);

                PdfReader R11C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R11C8Page = writer.GetImportedPage(R11C8File, 1);
                var R11C8PDF = writer.GetImportedPage(R11C8File, 1);
                var R11C8 = new System.Drawing.Drawing2D.Matrix();
                R11C8.Translate(567f, 1620);
                R11C8.Rotate(90);
                writer.DirectContent.AddTemplate(R11C8Page, R11C8);

                PdfReader R11C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R11C9Page = writer.GetImportedPage(R11C9File, 1);
                var R11C9PDF = writer.GetImportedPage(R11C9File, 1);
                var R11C9 = new System.Drawing.Drawing2D.Matrix();
                R11C9.Translate(621f, 1620);
                R11C9.Rotate(90);
                writer.DirectContent.AddTemplate(R11C9Page, R11C9);

                PdfReader R11C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R11C10Page = writer.GetImportedPage(R11C10File, 1);
                var R11C10PDF = writer.GetImportedPage(R11C10File, 1);
                var R11C10 = new System.Drawing.Drawing2D.Matrix();
                R11C10.Translate(693f, 1620);
                R11C10.Rotate(90);
                writer.DirectContent.AddTemplate(R11C10Page, R11C10);

                PdfReader R11C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R11C11Page = writer.GetImportedPage(R11C11File, 1);
                var R11C11PDF = writer.GetImportedPage(R11C11File, 1);
                var R11C11 = new System.Drawing.Drawing2D.Matrix();
                R11C11.Translate(747f, 1620);
                R11C11.Rotate(90);
                writer.DirectContent.AddTemplate(R11C11Page, R11C11);

                PdfReader R11C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R11C12Page = writer.GetImportedPage(R11C12File, 1);
                var R11C12PDF = writer.GetImportedPage(R11C12File, 1);
                var R11C12 = new System.Drawing.Drawing2D.Matrix();
                R11C12.Translate(801f, 1620);
                R11C12.Rotate(90);
                writer.DirectContent.AddTemplate(R11C12Page, R11C12);

                //Row 12
                PdfReader R12C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R12C1Page = writer.GetImportedPage(R12C1File, 1);
                var R12C1PDF = writer.GetImportedPage(R12C1File, 1);
                var R12C1 = new System.Drawing.Drawing2D.Matrix();
                R12C1.Translate(153f, 1782);
                R12C1.Rotate(90);
                writer.DirectContent.AddTemplate(R12C1Page, R12C1);

                PdfReader R12C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R12C2Page = writer.GetImportedPage(R12C2File, 1);
                var R12C2PDF = writer.GetImportedPage(R12C2File, 1);
                var R12C2 = new System.Drawing.Drawing2D.Matrix();
                R12C2.Translate(207f, 1782);
                R12C2.Rotate(90);
                writer.DirectContent.AddTemplate(R12C2Page, R12C2);

                PdfReader R12C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R12C3Page = writer.GetImportedPage(R12C3File, 1);
                var R12C3PDF = writer.GetImportedPage(R12C3File, 1);
                var R12C3 = new System.Drawing.Drawing2D.Matrix();
                R12C3.Translate(261f, 1782);
                R12C3.Rotate(90);
                writer.DirectContent.AddTemplate(R12C3Page, R12C3);

                PdfReader R12C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R12C4Page = writer.GetImportedPage(R12C4File, 1);
                var R12C4PDF = writer.GetImportedPage(R12C4File, 1);
                var R12C4 = new System.Drawing.Drawing2D.Matrix();
                R12C4.Translate(333f, 1782);
                R12C4.Rotate(90);
                writer.DirectContent.AddTemplate(R12C4Page, R12C4);

                PdfReader R12C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R12C5Page = writer.GetImportedPage(R12C5File, 1);
                var R12C5PDF = writer.GetImportedPage(R12C5File, 1);
                var R12C5 = new System.Drawing.Drawing2D.Matrix();
                R12C5.Translate(387f, 1782);
                R12C5.Rotate(90);
                writer.DirectContent.AddTemplate(R12C5Page, R12C5);

                PdfReader R12C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R12C6Page = writer.GetImportedPage(R12C6File, 1);
                var R12C6PDF = writer.GetImportedPage(R12C6File, 1);
                var R12C6 = new System.Drawing.Drawing2D.Matrix();
                R12C6.Translate(441f, 1782);
                R12C6.Rotate(90);
                writer.DirectContent.AddTemplate(R12C6Page, R12C6);

                PdfReader R12C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R12C7Page = writer.GetImportedPage(R12C7File, 1);
                var R12C7PDF = writer.GetImportedPage(R12C7File, 1);
                var R12C7 = new System.Drawing.Drawing2D.Matrix();
                R12C7.Translate(513f, 1782);
                R12C7.Rotate(90);
                writer.DirectContent.AddTemplate(R12C7Page, R12C7);

                PdfReader R12C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R12C8Page = writer.GetImportedPage(R12C8File, 1);
                var R12C8PDF = writer.GetImportedPage(R12C8File, 1);
                var R12C8 = new System.Drawing.Drawing2D.Matrix();
                R12C8.Translate(567f, 1782);
                R12C8.Rotate(90);
                writer.DirectContent.AddTemplate(R12C8Page, R12C8);

                PdfReader R12C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R12C9Page = writer.GetImportedPage(R12C9File, 1);
                var R12C9PDF = writer.GetImportedPage(R12C9File, 1);
                var R12C9 = new System.Drawing.Drawing2D.Matrix();
                R12C9.Translate(621f, 1782);
                R12C9.Rotate(90);
                writer.DirectContent.AddTemplate(R12C9Page, R12C9);

                PdfReader R12C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R12C10Page = writer.GetImportedPage(R12C10File, 1);
                var R12C10PDF = writer.GetImportedPage(R12C10File, 1);
                var R12C10 = new System.Drawing.Drawing2D.Matrix();
                R12C10.Translate(693f, 1782);
                R12C10.Rotate(90);
                writer.DirectContent.AddTemplate(R12C10Page, R12C10);

                PdfReader R12C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R12C11Page = writer.GetImportedPage(R12C11File, 1);
                var R12C11PDF = writer.GetImportedPage(R12C11File, 1);
                var R12C11 = new System.Drawing.Drawing2D.Matrix();
                R12C11.Translate(747f, 1782);
                R12C11.Rotate(90);
                writer.DirectContent.AddTemplate(R12C11Page, R12C11);

                PdfReader R12C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R12C12Page = writer.GetImportedPage(R12C12File, 1);
                var R12C12PDF = writer.GetImportedPage(R12C12File, 1);
                var R12C12 = new System.Drawing.Drawing2D.Matrix();
                R12C12.Translate(801f, 1782);
                R12C12.Rotate(90);
                writer.DirectContent.AddTemplate(R12C12Page, R12C12);

                //Row 13
                PdfReader R13C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R13C1Page = writer.GetImportedPage(R13C1File, 1);
                var R13C1PDF = writer.GetImportedPage(R13C1File, 1);
                var R13C1 = new System.Drawing.Drawing2D.Matrix();
                R13C1.Translate(153f, 1944);
                R13C1.Rotate(90);
                writer.DirectContent.AddTemplate(R13C1Page, R13C1);

                PdfReader R13C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R13C2Page = writer.GetImportedPage(R13C2File, 1);
                var R13C2PDF = writer.GetImportedPage(R13C2File, 1);
                var R13C2 = new System.Drawing.Drawing2D.Matrix();
                R13C2.Translate(207f, 1944);
                R13C2.Rotate(90);
                writer.DirectContent.AddTemplate(R13C2Page, R13C2);

                PdfReader R13C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R13C3Page = writer.GetImportedPage(R13C3File, 1);
                var R13C3PDF = writer.GetImportedPage(R13C3File, 1);
                var R13C3 = new System.Drawing.Drawing2D.Matrix();
                R13C3.Translate(261f, 1944);
                R13C3.Rotate(90);
                writer.DirectContent.AddTemplate(R13C3Page, R13C3);

                PdfReader R13C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R13C4Page = writer.GetImportedPage(R13C4File, 1);
                var R13C4PDF = writer.GetImportedPage(R13C4File, 1);
                var R13C4 = new System.Drawing.Drawing2D.Matrix();
                R13C4.Translate(333f, 1944);
                R13C4.Rotate(90);
                writer.DirectContent.AddTemplate(R13C4Page, R13C4);

                PdfReader R13C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R13C5Page = writer.GetImportedPage(R13C5File, 1);
                var R13C5PDF = writer.GetImportedPage(R13C5File, 1);
                var R13C5 = new System.Drawing.Drawing2D.Matrix();
                R13C5.Translate(387f, 1944);
                R13C5.Rotate(90);
                writer.DirectContent.AddTemplate(R13C5Page, R13C5);

                PdfReader R13C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R13C6Page = writer.GetImportedPage(R13C6File, 1);
                var R13C6PDF = writer.GetImportedPage(R13C6File, 1);
                var R13C6 = new System.Drawing.Drawing2D.Matrix();
                R13C6.Translate(441f, 1944);
                R13C6.Rotate(90);
                writer.DirectContent.AddTemplate(R13C6Page, R13C6);

                PdfReader R13C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R13C7Page = writer.GetImportedPage(R13C7File, 1);
                var R13C7PDF = writer.GetImportedPage(R13C7File, 1);
                var R13C7 = new System.Drawing.Drawing2D.Matrix();
                R13C7.Translate(513f, 1944);
                R13C7.Rotate(90);
                writer.DirectContent.AddTemplate(R13C7Page, R13C7);

                PdfReader R13C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R13C8Page = writer.GetImportedPage(R13C8File, 1);
                var R13C8PDF = writer.GetImportedPage(R13C8File, 1);
                var R13C8 = new System.Drawing.Drawing2D.Matrix();
                R13C8.Translate(567f, 1944);
                R13C8.Rotate(90);
                writer.DirectContent.AddTemplate(R13C8Page, R13C8);

                PdfReader R13C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R13C9Page = writer.GetImportedPage(R13C9File, 1);
                var R13C9PDF = writer.GetImportedPage(R13C9File, 1);
                var R13C9 = new System.Drawing.Drawing2D.Matrix();
                R13C9.Translate(621f, 1944);
                R13C9.Rotate(90);
                writer.DirectContent.AddTemplate(R13C9Page, R13C9);

                PdfReader R13C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R13C10Page = writer.GetImportedPage(R13C10File, 1);
                var R13C10PDF = writer.GetImportedPage(R13C10File, 1);
                var R13C10 = new System.Drawing.Drawing2D.Matrix();
                R13C10.Translate(693f, 1944);
                R13C10.Rotate(90);
                writer.DirectContent.AddTemplate(R13C10Page, R13C10);

                PdfReader R13C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R13C11Page = writer.GetImportedPage(R13C11File, 1);
                var R13C11PDF = writer.GetImportedPage(R13C11File, 1);
                var R13C11 = new System.Drawing.Drawing2D.Matrix();
                R13C11.Translate(747f, 1944);
                R13C11.Rotate(90);
                writer.DirectContent.AddTemplate(R13C11Page, R13C11);

                PdfReader R13C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R13C12Page = writer.GetImportedPage(R13C12File, 1);
                var R13C12PDF = writer.GetImportedPage(R13C12File, 1);
                var R13C12 = new System.Drawing.Drawing2D.Matrix();
                R13C12.Translate(801f, 1944);
                R13C12.Rotate(90);
                writer.DirectContent.AddTemplate(R13C12Page, R13C12);

                //Row 14
                PdfReader R14C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R14C1Page = writer.GetImportedPage(R14C1File, 1);
                var R14C1PDF = writer.GetImportedPage(R14C1File, 1);
                var R14C1 = new System.Drawing.Drawing2D.Matrix();
                R14C1.Translate(153f, 2106);
                R14C1.Rotate(90);
                writer.DirectContent.AddTemplate(R14C1Page, R14C1);

                PdfReader R14C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R14C2Page = writer.GetImportedPage(R14C2File, 1);
                var R14C2PDF = writer.GetImportedPage(R14C2File, 1);
                var R14C2 = new System.Drawing.Drawing2D.Matrix();
                R14C2.Translate(207f, 2106);
                R14C2.Rotate(90);
                writer.DirectContent.AddTemplate(R14C2Page, R14C2);

                PdfReader R14C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R14C3Page = writer.GetImportedPage(R14C3File, 1);
                var R14C3PDF = writer.GetImportedPage(R14C3File, 1);
                var R14C3 = new System.Drawing.Drawing2D.Matrix();
                R14C3.Translate(261f, 2106);
                R14C3.Rotate(90);
                writer.DirectContent.AddTemplate(R14C3Page, R14C3);

                PdfReader R14C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R14C4Page = writer.GetImportedPage(R14C4File, 1);
                var R14C4PDF = writer.GetImportedPage(R14C4File, 1);
                var R14C4 = new System.Drawing.Drawing2D.Matrix();
                R14C4.Translate(333f, 2106);
                R14C4.Rotate(90);
                writer.DirectContent.AddTemplate(R14C4Page, R14C4);

                PdfReader R14C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R14C5Page = writer.GetImportedPage(R14C5File, 1);
                var R14C5PDF = writer.GetImportedPage(R14C5File, 1);
                var R14C5 = new System.Drawing.Drawing2D.Matrix();
                R14C5.Translate(387f, 2106);
                R14C5.Rotate(90);
                writer.DirectContent.AddTemplate(R14C5Page, R14C5);

                PdfReader R14C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R14C6Page = writer.GetImportedPage(R14C6File, 1);
                var R14C6PDF = writer.GetImportedPage(R14C6File, 1);
                var R14C6 = new System.Drawing.Drawing2D.Matrix();
                R14C6.Translate(441f, 2106);
                R14C6.Rotate(90);
                writer.DirectContent.AddTemplate(R14C6Page, R14C6);

                PdfReader R14C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R14C7Page = writer.GetImportedPage(R14C7File, 1);
                var R14C7PDF = writer.GetImportedPage(R14C7File, 1);
                var R14C7 = new System.Drawing.Drawing2D.Matrix();
                R14C7.Translate(513f, 2106);
                R14C7.Rotate(90);
                writer.DirectContent.AddTemplate(R14C7Page, R14C7);

                PdfReader R14C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R14C8Page = writer.GetImportedPage(R14C8File, 1);
                var R14C8PDF = writer.GetImportedPage(R14C8File, 1);
                var R14C8 = new System.Drawing.Drawing2D.Matrix();
                R14C8.Translate(567f, 2106);
                R14C8.Rotate(90);
                writer.DirectContent.AddTemplate(R14C8Page, R14C8);

                PdfReader R14C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R14C9Page = writer.GetImportedPage(R14C9File, 1);
                var R14C9PDF = writer.GetImportedPage(R14C9File, 1);
                var R14C9 = new System.Drawing.Drawing2D.Matrix();
                R14C9.Translate(621f, 2106);
                R14C9.Rotate(90);
                writer.DirectContent.AddTemplate(R14C9Page, R14C9);

                PdfReader R14C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R14C10Page = writer.GetImportedPage(R14C10File, 1);
                var R14C10PDF = writer.GetImportedPage(R14C10File, 1);
                var R14C10 = new System.Drawing.Drawing2D.Matrix();
                R14C10.Translate(693f, 2106);
                R14C10.Rotate(90);
                writer.DirectContent.AddTemplate(R14C10Page, R14C10);

                PdfReader R14C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R14C11Page = writer.GetImportedPage(R14C11File, 1);
                var R14C11PDF = writer.GetImportedPage(R14C11File, 1);
                var R14C11 = new System.Drawing.Drawing2D.Matrix();
                R14C11.Translate(747f, 2106);
                R14C11.Rotate(90);
                writer.DirectContent.AddTemplate(R14C11Page, R14C11);

                PdfReader R14C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R14C12Page = writer.GetImportedPage(R14C12File, 1);
                var R14C12PDF = writer.GetImportedPage(R14C12File, 1);
                var R14C12 = new System.Drawing.Drawing2D.Matrix();
                R14C12.Translate(801f, 2106);
                R14C12.Rotate(90);
                writer.DirectContent.AddTemplate(R14C12Page, R14C12);

                //Row 15
                PdfReader R15C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R15C1Page = writer.GetImportedPage(R15C1File, 1);
                var R15C1PDF = writer.GetImportedPage(R15C1File, 1);
                var R15C1 = new System.Drawing.Drawing2D.Matrix();
                R15C1.Translate(153f, 2268);
                R15C1.Rotate(90);
                writer.DirectContent.AddTemplate(R15C1Page, R15C1);

                PdfReader R15C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R15C2Page = writer.GetImportedPage(R15C2File, 1);
                var R15C2PDF = writer.GetImportedPage(R15C2File, 1);
                var R15C2 = new System.Drawing.Drawing2D.Matrix();
                R15C2.Translate(207f, 2268);
                R15C2.Rotate(90);
                writer.DirectContent.AddTemplate(R15C2Page, R15C2);

                PdfReader R15C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R15C3Page = writer.GetImportedPage(R15C3File, 1);
                var R15C3PDF = writer.GetImportedPage(R15C3File, 1);
                var R15C3 = new System.Drawing.Drawing2D.Matrix();
                R15C3.Translate(261f, 2268);
                R15C3.Rotate(90);
                writer.DirectContent.AddTemplate(R15C3Page, R15C3);

                PdfReader R15C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R15C4Page = writer.GetImportedPage(R15C4File, 1);
                var R15C4PDF = writer.GetImportedPage(R15C4File, 1);
                var R15C4 = new System.Drawing.Drawing2D.Matrix();
                R15C4.Translate(333f, 2268);
                R15C4.Rotate(90);
                writer.DirectContent.AddTemplate(R15C4Page, R15C4);

                PdfReader R15C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R15C5Page = writer.GetImportedPage(R15C5File, 1);
                var R15C5PDF = writer.GetImportedPage(R15C5File, 1);
                var R15C5 = new System.Drawing.Drawing2D.Matrix();
                R15C5.Translate(387f, 2268);
                R15C5.Rotate(90);
                writer.DirectContent.AddTemplate(R15C5Page, R15C5);

                PdfReader R15C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R15C6Page = writer.GetImportedPage(R15C6File, 1);
                var R15C6PDF = writer.GetImportedPage(R15C6File, 1);
                var R15C6 = new System.Drawing.Drawing2D.Matrix();
                R15C6.Translate(441f, 2268);
                R15C6.Rotate(90);
                writer.DirectContent.AddTemplate(R15C6Page, R15C6);

                PdfReader R15C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R15C7Page = writer.GetImportedPage(R15C7File, 1);
                var R15C7PDF = writer.GetImportedPage(R15C7File, 1);
                var R15C7 = new System.Drawing.Drawing2D.Matrix();
                R15C7.Translate(513f, 2268);
                R15C7.Rotate(90);
                writer.DirectContent.AddTemplate(R15C7Page, R15C7);

                PdfReader R15C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R15C8Page = writer.GetImportedPage(R15C8File, 1);
                var R15C8PDF = writer.GetImportedPage(R15C8File, 1);
                var R15C8 = new System.Drawing.Drawing2D.Matrix();
                R15C8.Translate(567f, 2268);
                R15C8.Rotate(90);
                writer.DirectContent.AddTemplate(R15C8Page, R15C8);

                PdfReader R15C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R15C9Page = writer.GetImportedPage(R15C9File, 1);
                var R15C9PDF = writer.GetImportedPage(R15C9File, 1);
                var R15C9 = new System.Drawing.Drawing2D.Matrix();
                R15C9.Translate(621f, 2268);
                R15C9.Rotate(90);
                writer.DirectContent.AddTemplate(R15C9Page, R15C9);

                PdfReader R15C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                PdfImportedPage R15C10Page = writer.GetImportedPage(R15C10File, 1);
                var R15C10PDF = writer.GetImportedPage(R15C10File, 1);
                var R15C10 = new System.Drawing.Drawing2D.Matrix();
                R15C10.Translate(693f, 2268);
                R15C10.Rotate(90);
                writer.DirectContent.AddTemplate(R15C10Page, R15C10);

                PdfReader R15C11File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[10]) + ".pdf");
                PdfImportedPage R15C11Page = writer.GetImportedPage(R15C11File, 1);
                var R15C11PDF = writer.GetImportedPage(R15C11File, 1);
                var R15C11 = new System.Drawing.Drawing2D.Matrix();
                R15C11.Translate(747f, 2268);
                R15C11.Rotate(90);
                writer.DirectContent.AddTemplate(R15C11Page, R15C11);

                PdfReader R15C12File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[11]) + ".pdf");
                PdfImportedPage R15C12Page = writer.GetImportedPage(R15C12File, 1);
                var R15C12PDF = writer.GetImportedPage(R15C12File, 1);
                var R15C12 = new System.Drawing.Drawing2D.Matrix();
                R15C12.Translate(801f, 2268);
                R15C12.Rotate(90);
                writer.DirectContent.AddTemplate(R15C12Page, R15C12);

                itemTotal.RemoveRange(0, 12);

                cb.SetLineWidth(18f);

                //Cropmarks Horizontal
                cb.MoveTo(0f, 0);
                cb.LineTo(900f, 0);
                cb.Stroke();

                cb.MoveTo(0f, 810);
                cb.LineTo(900f, 810);
                cb.Stroke();

                cb.MoveTo(0f, 1620);
                cb.LineTo(900f, 1620);
                cb.Stroke();

                cb.MoveTo(0f, 2430);
                cb.LineTo(900f, 2430);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27f, 0);
                cb.LineTo(873f, 0);
                cb.LineTo(873f, 2430);
                cb.LineTo(27f, 2430);
                cb.Fill();
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf2x1_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(162f, 90f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 90f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 1944));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 6 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[4]);
                        itemTotal.Add(itemPrint[5]);
                        itemPrint.RemoveRange(0, 6);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 12);
                        diffPerPage.Add("6 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 6);

                    }
                    else if (itemPrint.Count() % 3 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[2]);
                        itemPrint.RemoveRange(0, 3);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 24);
                        diffPerPage.Add("3 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 3);
                    }
                    else if (itemPrint.Count() % 2 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemPrint.RemoveRange(0, 2);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 36);
                        diffPerPage.Add("2 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 2);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 72);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }

            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();

                //Row 1
                PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(270f, 0f);
                R1C1.Rotate(90);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(360f, 0f);
                R1C2.Rotate(90);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(450f, 0f);
                R1C3.Rotate(90);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                var R1C4 = new System.Drawing.Drawing2D.Matrix();
                R1C4.Translate(540f, 0f);
                R1C4.Rotate(90);
                writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                PdfReader R1C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R1C5Page = writer.GetImportedPage(R1C5File, 1);
                var R1C5PDF = writer.GetImportedPage(R1C5File, 1);
                var R1C5 = new System.Drawing.Drawing2D.Matrix();
                R1C5.Translate(630f, 0f);
                R1C5.Rotate(90);
                writer.DirectContent.AddTemplate(R1C5Page, R1C5);

                PdfReader R1C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R1C6Page = writer.GetImportedPage(R1C6File, 1);
                var R1C6PDF = writer.GetImportedPage(R1C6File, 1);
                var R1C6 = new System.Drawing.Drawing2D.Matrix();
                R1C6.Translate(720f, 0f);
                R1C6.Rotate(90);
                writer.DirectContent.AddTemplate(R1C6Page, R1C6);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(270f, 162f);
                R2C1.Rotate(90);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(360f, 162f);
                R2C2.Rotate(90);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(450f, 162f);
                R2C3.Rotate(90);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                PdfReader R2C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R2C4Page = writer.GetImportedPage(R2C4File, 1);
                var R2C4PDF = writer.GetImportedPage(R2C4File, 1);
                var R2C4 = new System.Drawing.Drawing2D.Matrix();
                R2C4.Translate(540f, 162f);
                R2C4.Rotate(90);
                writer.DirectContent.AddTemplate(R2C4Page, R2C4);

                PdfReader R2C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R2C5Page = writer.GetImportedPage(R2C5File, 1);
                var R2C5PDF = writer.GetImportedPage(R2C5File, 1);
                var R2C5 = new System.Drawing.Drawing2D.Matrix();
                R2C5.Translate(630f, 162f);
                R2C5.Rotate(90);
                writer.DirectContent.AddTemplate(R2C5Page, R2C5);

                PdfReader R2C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R2C6Page = writer.GetImportedPage(R2C6File, 1);
                var R2C6PDF = writer.GetImportedPage(R2C6File, 1);
                var R2C6 = new System.Drawing.Drawing2D.Matrix();
                R2C6.Translate(720f, 162f);
                R2C6.Rotate(90);
                writer.DirectContent.AddTemplate(R2C6Page, R2C6);


                //Row 3
                PdfReader R3C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(270f, 324f);
                R3C1.Rotate(90);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(360f, 324f);
                R3C2.Rotate(90);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(450f, 325f);
                R3C3.Rotate(90);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                PdfReader R3C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R3C4Page = writer.GetImportedPage(R3C4File, 1);
                var R3C4PDF = writer.GetImportedPage(R3C4File, 1);
                var R3C4 = new System.Drawing.Drawing2D.Matrix();
                R3C4.Translate(540f, 324f);
                R3C4.Rotate(90);
                writer.DirectContent.AddTemplate(R3C4Page, R3C4);

                PdfReader R3C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R3C5Page = writer.GetImportedPage(R3C5File, 1);
                var R3C5PDF = writer.GetImportedPage(R3C5File, 1);
                var R3C5 = new System.Drawing.Drawing2D.Matrix();
                R3C5.Translate(630f, 324f);
                R3C5.Rotate(90);
                writer.DirectContent.AddTemplate(R3C5Page, R3C5);

                PdfReader R3C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R3C6Page = writer.GetImportedPage(R3C6File, 1);
                var R3C6PDF = writer.GetImportedPage(R3C6File, 1);
                var R3C6 = new System.Drawing.Drawing2D.Matrix();
                R3C6.Translate(720f, 324f);
                R3C6.Rotate(90);
                writer.DirectContent.AddTemplate(R3C6Page, R3C6);

                //Row 4
                PdfReader R4C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(270f, 486f);
                R4C1.Rotate(90);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(360f, 486f);
                R4C2.Rotate(90);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(450f, 486f);
                R4C3.Rotate(90);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);

                PdfReader R4C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R4C4Page = writer.GetImportedPage(R4C4File, 1);
                var R4C4PDF = writer.GetImportedPage(R4C4File, 1);
                var R4C4 = new System.Drawing.Drawing2D.Matrix();
                R4C4.Translate(540f, 486f);
                R4C4.Rotate(90);
                writer.DirectContent.AddTemplate(R4C4Page, R4C4);

                PdfReader R4C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R4C5Page = writer.GetImportedPage(R4C5File, 1);
                var R4C5PDF = writer.GetImportedPage(R4C5File, 1);
                var R4C5 = new System.Drawing.Drawing2D.Matrix();
                R4C5.Translate(630f, 486f);
                R4C5.Rotate(90);
                writer.DirectContent.AddTemplate(R4C5Page, R4C5);

                PdfReader R4C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R4C6Page = writer.GetImportedPage(R4C6File, 1);
                var R4C6PDF = writer.GetImportedPage(R4C6File, 1);
                var R4C6 = new System.Drawing.Drawing2D.Matrix();
                R4C6.Translate(720f, 486f);
                R4C6.Rotate(90);
                writer.DirectContent.AddTemplate(R4C6Page, R4C6);

                //Row 5
                PdfReader R5C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(270f, 648f);
                R5C1.Rotate(90);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(360f, 648f);
                R5C2.Rotate(90);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                PdfReader R5C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R5C3Page = writer.GetImportedPage(R5C3File, 1);
                var R5C3PDF = writer.GetImportedPage(R5C3File, 1);
                var R5C3 = new System.Drawing.Drawing2D.Matrix();
                R5C3.Translate(450f, 648f);
                R5C3.Rotate(90);
                writer.DirectContent.AddTemplate(R5C3Page, R5C3);

                PdfReader R5C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R5C4Page = writer.GetImportedPage(R5C4File, 1);
                var R5C4PDF = writer.GetImportedPage(R5C4File, 1);
                var R5C4 = new System.Drawing.Drawing2D.Matrix();
                R5C4.Translate(540f, 648f);
                R5C4.Rotate(90);
                writer.DirectContent.AddTemplate(R5C4Page, R5C4);

                PdfReader R5C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R5C5Page = writer.GetImportedPage(R5C5File, 1);
                var R5C5PDF = writer.GetImportedPage(R5C5File, 1);
                var R5C5 = new System.Drawing.Drawing2D.Matrix();
                R5C5.Translate(630f, 648f);
                R5C5.Rotate(90);
                writer.DirectContent.AddTemplate(R5C5Page, R5C5);

                PdfReader R5C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R5C6Page = writer.GetImportedPage(R5C6File, 1);
                var R5C6PDF = writer.GetImportedPage(R5C6File, 1);
                var R5C6 = new System.Drawing.Drawing2D.Matrix();
                R5C6.Translate(720f, 648f);
                R5C6.Rotate(90);
                writer.DirectContent.AddTemplate(R5C6Page, R5C6);

                //Row 6
                PdfReader R6C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R6C1Page = writer.GetImportedPage(R6C1File, 1);
                var R6C1PDF = writer.GetImportedPage(R6C1File, 1);
                var R6C1 = new System.Drawing.Drawing2D.Matrix();
                R6C1.Translate(270f, 810f);
                R6C1.Rotate(90);
                writer.DirectContent.AddTemplate(R6C1Page, R6C1);

                PdfReader R6C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R6C2Page = writer.GetImportedPage(R6C2File, 1);
                var R6C2PDF = writer.GetImportedPage(R6C2File, 1);
                var R6C2 = new System.Drawing.Drawing2D.Matrix();
                R6C2.Translate(360f, 810f);
                R6C2.Rotate(90);
                writer.DirectContent.AddTemplate(R6C2Page, R6C2);

                PdfReader R6C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R6C3Page = writer.GetImportedPage(R6C3File, 1);
                var R6C3PDF = writer.GetImportedPage(R6C3File, 1);
                var R6C3 = new System.Drawing.Drawing2D.Matrix();
                R6C3.Translate(450f, 810f);
                R6C3.Rotate(90);
                writer.DirectContent.AddTemplate(R6C3Page, R6C3);

                PdfReader R6C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R6C4Page = writer.GetImportedPage(R6C4File, 1);
                var R6C4PDF = writer.GetImportedPage(R6C4File, 1);
                var R6C4 = new System.Drawing.Drawing2D.Matrix();
                R6C4.Translate(540f, 810f);
                R6C4.Rotate(90);
                writer.DirectContent.AddTemplate(R6C4Page, R6C4);

                PdfReader R6C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R6C5Page = writer.GetImportedPage(R6C5File, 1);
                var R6C5PDF = writer.GetImportedPage(R6C5File, 1);
                var R6C5 = new System.Drawing.Drawing2D.Matrix();
                R6C5.Translate(630f, 810f);
                R6C5.Rotate(90);
                writer.DirectContent.AddTemplate(R6C5Page, R6C5);

                PdfReader R6C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R6C6Page = writer.GetImportedPage(R6C6File, 1);
                var R6C6PDF = writer.GetImportedPage(R6C6File, 1);
                var R6C6 = new System.Drawing.Drawing2D.Matrix();
                R6C6.Translate(720f, 810f);
                R6C6.Rotate(90);
                writer.DirectContent.AddTemplate(R6C6Page, R6C6);


                //Row 7
                PdfReader R7C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R7C1Page = writer.GetImportedPage(R7C1File, 1);
                var R7C1PDF = writer.GetImportedPage(R7C1File, 1);
                var R7C1 = new System.Drawing.Drawing2D.Matrix();
                R7C1.Translate(270f, 972f);
                R7C1.Rotate(90);
                writer.DirectContent.AddTemplate(R7C1Page, R7C1);

                PdfReader R7C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R7C2Page = writer.GetImportedPage(R7C2File, 1);
                var R7C2PDF = writer.GetImportedPage(R7C2File, 1);
                var R7C2 = new System.Drawing.Drawing2D.Matrix();
                R7C2.Translate(360f, 972f);
                R7C2.Rotate(90);
                writer.DirectContent.AddTemplate(R7C2Page, R7C2);

                PdfReader R7C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R7C3Page = writer.GetImportedPage(R7C3File, 1);
                var R7C3PDF = writer.GetImportedPage(R7C3File, 1);
                var R7C3 = new System.Drawing.Drawing2D.Matrix();
                R7C3.Translate(450f, 972f);
                R7C3.Rotate(90);
                writer.DirectContent.AddTemplate(R7C3Page, R7C3);

                PdfReader R7C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R7C4Page = writer.GetImportedPage(R7C4File, 1);
                var R7C4PDF = writer.GetImportedPage(R7C4File, 1);
                var R7C4 = new System.Drawing.Drawing2D.Matrix();
                R7C4.Translate(540f, 972f);
                R7C4.Rotate(90);
                writer.DirectContent.AddTemplate(R7C4Page, R7C4);

                PdfReader R7C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R7C5Page = writer.GetImportedPage(R7C5File, 1);
                var R7C5PDF = writer.GetImportedPage(R7C5File, 1);
                var R7C5 = new System.Drawing.Drawing2D.Matrix();
                R7C5.Translate(630f, 972f);
                R7C5.Rotate(90);
                writer.DirectContent.AddTemplate(R7C5Page, R7C5);

                PdfReader R7C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R7C6Page = writer.GetImportedPage(R7C6File, 1);
                var R7C6PDF = writer.GetImportedPage(R7C6File, 1);
                var R7C6 = new System.Drawing.Drawing2D.Matrix();
                R7C6.Translate(720f, 972f);
                R7C6.Rotate(90);
                writer.DirectContent.AddTemplate(R7C6Page, R7C6);

                //Row 8
                PdfReader R8C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R8C1Page = writer.GetImportedPage(R8C1File, 1);
                var R8C1PDF = writer.GetImportedPage(R8C1File, 1);
                var R8C1 = new System.Drawing.Drawing2D.Matrix();
                R8C1.Translate(270f, 1134f);
                R8C1.Rotate(90);
                writer.DirectContent.AddTemplate(R8C1Page, R8C1);

                PdfReader R8C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R8C2Page = writer.GetImportedPage(R8C2File, 1);
                var R8C2PDF = writer.GetImportedPage(R8C2File, 1);
                var R8C2 = new System.Drawing.Drawing2D.Matrix();
                R8C2.Translate(360f, 1134f);
                R8C2.Rotate(90);
                writer.DirectContent.AddTemplate(R8C2Page, R8C2);

                PdfReader R8C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R8C3Page = writer.GetImportedPage(R8C3File, 1);
                var R8C3PDF = writer.GetImportedPage(R8C3File, 1);
                var R8C3 = new System.Drawing.Drawing2D.Matrix();
                R8C3.Translate(450f, 1134f);
                R8C3.Rotate(90);
                writer.DirectContent.AddTemplate(R8C3Page, R8C3);

                PdfReader R8C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R8C4Page = writer.GetImportedPage(R8C4File, 1);
                var R8C4PDF = writer.GetImportedPage(R8C4File, 1);
                var R8C4 = new System.Drawing.Drawing2D.Matrix();
                R8C4.Translate(540f, 1134f);
                R8C4.Rotate(90);
                writer.DirectContent.AddTemplate(R8C4Page, R8C4);

                PdfReader R8C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R8C5Page = writer.GetImportedPage(R8C5File, 1);
                var R8C5PDF = writer.GetImportedPage(R8C5File, 1);
                var R8C5 = new System.Drawing.Drawing2D.Matrix();
                R8C5.Translate(630f, 1134f);
                R8C5.Rotate(90);
                writer.DirectContent.AddTemplate(R8C5Page, R8C5);

                PdfReader R8C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R8C6Page = writer.GetImportedPage(R8C6File, 1);
                var R8C6PDF = writer.GetImportedPage(R8C6File, 1);
                var R8C6 = new System.Drawing.Drawing2D.Matrix();
                R8C6.Translate(720f, 1134f);
                R8C6.Rotate(90);
                writer.DirectContent.AddTemplate(R8C6Page, R8C6);

                //Row 9
                PdfReader R9C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R9C1Page = writer.GetImportedPage(R9C1File, 1);
                var R9C1PDF = writer.GetImportedPage(R9C1File, 1);
                var R9C1 = new System.Drawing.Drawing2D.Matrix();
                R9C1.Translate(270f, 1296f);
                R9C1.Rotate(90);
                writer.DirectContent.AddTemplate(R9C1Page, R9C1);

                PdfReader R9C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R9C2Page = writer.GetImportedPage(R9C2File, 1);
                var R9C2PDF = writer.GetImportedPage(R9C2File, 1);
                var R9C2 = new System.Drawing.Drawing2D.Matrix();
                R9C2.Translate(360f, 1296f);
                R9C2.Rotate(90);
                writer.DirectContent.AddTemplate(R9C2Page, R9C2);

                PdfReader R9C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R9C3Page = writer.GetImportedPage(R9C3File, 1);
                var R9C3PDF = writer.GetImportedPage(R9C3File, 1);
                var R9C3 = new System.Drawing.Drawing2D.Matrix();
                R9C3.Translate(450f, 1296f);
                R9C3.Rotate(90);
                writer.DirectContent.AddTemplate(R9C3Page, R9C3);

                PdfReader R9C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R9C4Page = writer.GetImportedPage(R9C4File, 1);
                var R9C4PDF = writer.GetImportedPage(R9C4File, 1);
                var R9C4 = new System.Drawing.Drawing2D.Matrix();
                R9C4.Translate(540f, 1296f);
                R9C4.Rotate(90);
                writer.DirectContent.AddTemplate(R9C4Page, R9C4);

                PdfReader R9C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R9C5Page = writer.GetImportedPage(R9C5File, 1);
                var R9C5PDF = writer.GetImportedPage(R9C5File, 1);
                var R9C5 = new System.Drawing.Drawing2D.Matrix();
                R9C5.Translate(630f, 1296f);
                R9C5.Rotate(90);
                writer.DirectContent.AddTemplate(R9C5Page, R9C5);

                PdfReader R9C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R9C6Page = writer.GetImportedPage(R9C6File, 1);
                var R9C6PDF = writer.GetImportedPage(R9C6File, 1);
                var R9C6 = new System.Drawing.Drawing2D.Matrix();
                R9C6.Translate(720f, 1296f);
                R9C6.Rotate(90);
                writer.DirectContent.AddTemplate(R9C6Page, R9C6);

                //Row 10
                PdfReader R10C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R10C1Page = writer.GetImportedPage(R10C1File, 1);
                var R10C1PDF = writer.GetImportedPage(R10C1File, 1);
                var R10C1 = new System.Drawing.Drawing2D.Matrix();
                R10C1.Translate(270f, 1458f);
                R10C1.Rotate(90);
                writer.DirectContent.AddTemplate(R10C1Page, R10C1);

                PdfReader R10C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R10C2Page = writer.GetImportedPage(R10C2File, 1);
                var R10C2PDF = writer.GetImportedPage(R10C2File, 1);
                var R10C2 = new System.Drawing.Drawing2D.Matrix();
                R10C2.Translate(360f, 1458f);
                R10C2.Rotate(90);
                writer.DirectContent.AddTemplate(R10C2Page, R10C2);

                PdfReader R10C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R10C3Page = writer.GetImportedPage(R10C3File, 1);
                var R10C3PDF = writer.GetImportedPage(R10C3File, 1);
                var R10C3 = new System.Drawing.Drawing2D.Matrix();
                R10C3.Translate(450f, 1458f);
                R10C3.Rotate(90);
                writer.DirectContent.AddTemplate(R10C3Page, R10C3);

                PdfReader R10C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R10C4Page = writer.GetImportedPage(R10C4File, 1);
                var R10C4PDF = writer.GetImportedPage(R10C4File, 1);
                var R10C4 = new System.Drawing.Drawing2D.Matrix();
                R10C4.Translate(540f, 1458f);
                R10C4.Rotate(90);
                writer.DirectContent.AddTemplate(R10C4Page, R10C4);

                PdfReader R10C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R10C5Page = writer.GetImportedPage(R10C5File, 1);
                var R10C5PDF = writer.GetImportedPage(R10C5File, 1);
                var R10C5 = new System.Drawing.Drawing2D.Matrix();
                R10C5.Translate(630f, 1458f);
                R10C5.Rotate(90);
                writer.DirectContent.AddTemplate(R10C5Page, R10C5);

                PdfReader R10C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R10C6Page = writer.GetImportedPage(R10C6File, 1);
                var R10C6PDF = writer.GetImportedPage(R10C6File, 1);
                var R10C6 = new System.Drawing.Drawing2D.Matrix();
                R10C6.Translate(720f, 1458f);
                R10C6.Rotate(90);
                writer.DirectContent.AddTemplate(R10C6Page, R10C6);

                //Row 11
                PdfReader R11C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R11C1Page = writer.GetImportedPage(R11C1File, 1);
                var R11C1PDF = writer.GetImportedPage(R11C1File, 1);
                var R11C1 = new System.Drawing.Drawing2D.Matrix();
                R11C1.Translate(270f, 1620f);
                R11C1.Rotate(90);
                writer.DirectContent.AddTemplate(R11C1Page, R11C1);

                PdfReader R11C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R11C2Page = writer.GetImportedPage(R11C2File, 1);
                var R11C2PDF = writer.GetImportedPage(R11C2File, 1);
                var R11C2 = new System.Drawing.Drawing2D.Matrix();
                R11C2.Translate(360f, 1620f);
                R11C2.Rotate(90);
                writer.DirectContent.AddTemplate(R11C2Page, R11C2);

                PdfReader R11C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R11C3Page = writer.GetImportedPage(R11C3File, 1);
                var R11C3PDF = writer.GetImportedPage(R11C3File, 1);
                var R11C3 = new System.Drawing.Drawing2D.Matrix();
                R11C3.Translate(450f, 1620f);
                R11C3.Rotate(90);
                writer.DirectContent.AddTemplate(R11C3Page, R11C3);

                PdfReader R11C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R11C4Page = writer.GetImportedPage(R11C4File, 1);
                var R11C4PDF = writer.GetImportedPage(R11C4File, 1);
                var R11C4 = new System.Drawing.Drawing2D.Matrix();
                R11C4.Translate(540f, 1620f);
                R11C4.Rotate(90);
                writer.DirectContent.AddTemplate(R11C4Page, R11C4);

                PdfReader R11C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R11C5Page = writer.GetImportedPage(R11C5File, 1);
                var R11C5PDF = writer.GetImportedPage(R11C5File, 1);
                var R11C5 = new System.Drawing.Drawing2D.Matrix();
                R11C5.Translate(630f, 1620f);
                R11C5.Rotate(90);
                writer.DirectContent.AddTemplate(R11C5Page, R11C5);

                PdfReader R11C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R11C6Page = writer.GetImportedPage(R11C6File, 1);
                var R11C6PDF = writer.GetImportedPage(R11C6File, 1);
                var R11C6 = new System.Drawing.Drawing2D.Matrix();
                R11C6.Translate(720f, 1620f);
                R11C6.Rotate(90);
                writer.DirectContent.AddTemplate(R11C6Page, R11C6);

                //Row 12
                PdfReader R12C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R12C1Page = writer.GetImportedPage(R12C1File, 1);
                var R12C1PDF = writer.GetImportedPage(R12C1File, 1);
                var R12C1 = new System.Drawing.Drawing2D.Matrix();
                R12C1.Translate(270f, 1782f);
                R12C1.Rotate(90);
                writer.DirectContent.AddTemplate(R12C1Page, R12C1);

                PdfReader R12C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R12C2Page = writer.GetImportedPage(R12C2File, 1);
                var R12C2PDF = writer.GetImportedPage(R12C2File, 1);
                var R12C2 = new System.Drawing.Drawing2D.Matrix();
                R12C2.Translate(360f, 1782f);
                R12C2.Rotate(90);
                writer.DirectContent.AddTemplate(R12C2Page, R12C2);

                PdfReader R12C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R12C3Page = writer.GetImportedPage(R12C3File, 1);
                var R12C3PDF = writer.GetImportedPage(R12C3File, 1);
                var R12C3 = new System.Drawing.Drawing2D.Matrix();
                R12C3.Translate(450f, 1782f);
                R12C3.Rotate(90);
                writer.DirectContent.AddTemplate(R12C3Page, R12C3);

                PdfReader R12C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R12C4Page = writer.GetImportedPage(R12C4File, 1);
                var R12C4PDF = writer.GetImportedPage(R12C4File, 1);
                var R12C4 = new System.Drawing.Drawing2D.Matrix();
                R12C4.Translate(540f, 1782f);
                R12C4.Rotate(90);
                writer.DirectContent.AddTemplate(R12C4Page, R12C4);

                PdfReader R12C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R12C5Page = writer.GetImportedPage(R12C5File, 1);
                var R12C5PDF = writer.GetImportedPage(R12C5File, 1);
                var R12C5 = new System.Drawing.Drawing2D.Matrix();
                R12C5.Translate(630f, 1782f);
                R12C5.Rotate(90);
                writer.DirectContent.AddTemplate(R12C5Page, R12C5);

                PdfReader R12C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R12C6Page = writer.GetImportedPage(R12C6File, 1);
                var R12C6PDF = writer.GetImportedPage(R12C6File, 1);
                var R12C6 = new System.Drawing.Drawing2D.Matrix();
                R12C6.Translate(720f, 1782f);
                R12C6.Rotate(90);
                writer.DirectContent.AddTemplate(R12C6Page, R12C6);


                itemTotal.RemoveRange(0, 6);

                cb.SetLineWidth(18f);

                //Cropmarks Horizontal
                cb.MoveTo(0f, 0);
                cb.LineTo(900f, 0);
                cb.Stroke();

                cb.MoveTo(0f, 162);
                cb.LineTo(900f, 162);
                cb.Stroke();

                cb.MoveTo(0f, 324);
                cb.LineTo(900f, 324);
                cb.Stroke();

                cb.MoveTo(0f, 486);
                cb.LineTo(900f, 486);
                cb.Stroke();

                cb.MoveTo(0f, 648);
                cb.LineTo(900f, 648);
                cb.Stroke();

                cb.MoveTo(0f, 810);
                cb.LineTo(900f, 810);
                cb.Stroke();

                cb.MoveTo(0f, 972);
                cb.LineTo(900f, 972);
                cb.Stroke();

                cb.MoveTo(0f, 1134);
                cb.LineTo(900f, 1134);
                cb.Stroke();

                cb.MoveTo(0f, 1296);
                cb.LineTo(900f, 1296);
                cb.Stroke();

                cb.MoveTo(0f, 1458);
                cb.LineTo(900f, 1458);
                cb.Stroke();

                cb.MoveTo(0f, 1620);
                cb.LineTo(900f, 1620);
                cb.Stroke();

                cb.MoveTo(0f, 1782);
                cb.LineTo(900f, 1782);
                cb.Stroke();

                cb.MoveTo(0f, 1944);
                cb.LineTo(900f, 1944);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27f, 0);
                cb.LineTo(873f, 0);
                cb.LineTo(873f, 1944);
                cb.LineTo(27f, 1944);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf2_625x1_0625_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(207f, 94.5f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 94.5f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900f, 2740.5f));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 4 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemPrint.RemoveRange(0, 4);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 29);
                        diffPerPage.Add("4 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 4);
                    }
                    else if (itemPrint.Count() % 2 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemPrint.RemoveRange(0, 2);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 58);
                        diffPerPage.Add("2 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 2);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 116);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            /*while (itemTotal.Count() % 2 != 0)
            {
                itemTotal.Insert(0, itemTotal[0]);
            }*/

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                float stepDistance = 0f;
                for (int i = 1; i <= 29; i++)
                {
                    //Row 1
                    PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                    PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                    var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                    var R1C1 = new System.Drawing.Drawing2D.Matrix();
                    R1C1.Translate(36f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                    PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                    PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                    var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                    var R1C2 = new System.Drawing.Drawing2D.Matrix();
                    R1C2.Translate(243f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                    PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                    PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                    var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                    var R1C3 = new System.Drawing.Drawing2D.Matrix();
                    R1C3.Translate(450f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                    PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                    PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                    var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                    var R1C4 = new System.Drawing.Drawing2D.Matrix();
                    R1C4.Translate(657f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                    stepDistance = stepDistance + 94.5f;
                }
                stepDistance = 0;

                itemTotal.RemoveRange(0, 4);

                cb.SetLineWidth(18f);

                //Cropmarks Horizontal
                for (int i = 1; i <= 30; i++)
                {
                    cb.MoveTo(18f, stepDistance);
                    cb.LineTo(882f, stepDistance);
                    cb.Stroke();
                    stepDistance = stepDistance + 94.5f;
                }
                stepDistance = 0;

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(36f, 0);
                cb.LineTo(864f, 0);
                cb.LineTo(864f, 2740.5f);
                cb.LineTo(36f, 2740.5f);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf2_625x1_125_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(207f, 99f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 99f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900f, 2673f));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 4 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemPrint.RemoveRange(0, 4);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 27);
                        diffPerPage.Add("4 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 4);
                    }
                    else if (itemPrint.Count() % 2 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemPrint.RemoveRange(0, 2);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 54);
                        diffPerPage.Add("2 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 2);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 108);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            /*while (itemTotal.Count() % 2 != 0)
            {
                itemTotal.Insert(0, itemTotal[0]);
            }*/

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                float stepDistance = 0f;
                for (int i = 1; i <= 27; i++)
                {
                    //Row 1
                    PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                    PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                    var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                    var R1C1 = new System.Drawing.Drawing2D.Matrix();
                    R1C1.Translate(36f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                    PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                    PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                    var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                    var R1C2 = new System.Drawing.Drawing2D.Matrix();
                    R1C2.Translate(243f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                    PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                    PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                    var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                    var R1C3 = new System.Drawing.Drawing2D.Matrix();
                    R1C3.Translate(450f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                    PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                    PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                    var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                    var R1C4 = new System.Drawing.Drawing2D.Matrix();
                    R1C4.Translate(657f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                    stepDistance = stepDistance + 99f;
                }
                stepDistance = 0;

                itemTotal.RemoveRange(0, 4);

                cb.SetLineWidth(18f);

                //Cropmarks Horizontal
                for (int i = 1; i <= 10; i++)
                {
                    cb.MoveTo(18f, stepDistance);
                    cb.LineTo(882f, stepDistance);
                    cb.Stroke();
                    stepDistance = stepDistance + 297f;
                }
                stepDistance = 0;

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(36f, 0);
                cb.LineTo(864f, 0);
                cb.LineTo(864f, 2673f);
                cb.LineTo(36f, 2673f);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf2_75x0_312_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(207f, 31.5f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 31.5f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2677.5f));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 3 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemPrint.RemoveRange(0, 3);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 85);
                        diffPerPage.Add("3 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 3);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 255);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                float stepDistance = 0;
                for (int i = 1; i <= 85; i++)
                {
                    PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                    PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                    var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                    var R1C1 = new System.Drawing.Drawing2D.Matrix();
                    R1C1.Translate(139.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                    PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                    PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                    var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                    var R1C2 = new System.Drawing.Drawing2D.Matrix();
                    R1C2.Translate(346.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                    PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                    PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                    var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                    var R1C3 = new System.Drawing.Drawing2D.Matrix();
                    R1C3.Translate(553.5f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                    stepDistance = stepDistance + 31.5f;
                }
                stepDistance = 0;

                itemTotal.RemoveRange(0, 3);


                cb.SetLineWidth(18f);

                for (int i = 1; i <= 18; i++)
                {
                    //Cropmarks Horizontal
                    cb.MoveTo(117f, stepDistance);
                    cb.LineTo(783f, stepDistance);
                    cb.Stroke();
                    stepDistance = stepDistance + (31.5f * 5);
                }

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(135f, 0);
                cb.LineTo(765f, 0);
                cb.LineTo(765f, 2677.5f);
                cb.LineTo(135f, 2677.5f);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf3x1_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(234f, 90f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 90f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2340));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 9 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[4]);
                        itemTotal.Add(itemPrint[5]);
                        itemTotal.Add(itemPrint[6]);
                        itemTotal.Add(itemPrint[7]);
                        itemTotal.Add(itemPrint[8]);
                        itemPrint.RemoveRange(0, 9);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 10);
                        diffPerPage.Add("9 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 9);

                    }
                    else if (itemPrint.Count() % 3 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[2]);
                        itemPrint.RemoveRange(0, 3);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 30);
                        diffPerPage.Add("3 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 3);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 90);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }

            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();

                //Row 1
                PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(135f, 0f);
                R1C1.Rotate(90);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(225f, 0f);
                R1C2.Rotate(90);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(315f, 0f);
                R1C3.Rotate(90);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                var R1C4 = new System.Drawing.Drawing2D.Matrix();
                R1C4.Translate(405f, 0f);
                R1C4.Rotate(90);
                writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                PdfReader R1C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R1C5Page = writer.GetImportedPage(R1C5File, 1);
                var R1C5PDF = writer.GetImportedPage(R1C5File, 1);
                var R1C5 = new System.Drawing.Drawing2D.Matrix();
                R1C5.Translate(495f, 0f);
                R1C5.Rotate(90);
                writer.DirectContent.AddTemplate(R1C5Page, R1C5);

                PdfReader R1C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R1C6Page = writer.GetImportedPage(R1C6File, 1);
                var R1C6PDF = writer.GetImportedPage(R1C6File, 1);
                var R1C6 = new System.Drawing.Drawing2D.Matrix();
                R1C6.Translate(585f, 0f);
                R1C6.Rotate(90);
                writer.DirectContent.AddTemplate(R1C6Page, R1C6);

                PdfReader R1C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R1C7Page = writer.GetImportedPage(R1C7File, 1);
                var R1C7PDF = writer.GetImportedPage(R1C7File, 1);
                var R1C7 = new System.Drawing.Drawing2D.Matrix();
                R1C7.Translate(675f, 0f);
                R1C7.Rotate(90);
                writer.DirectContent.AddTemplate(R1C7Page, R1C7);

                PdfReader R1C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R1C8Page = writer.GetImportedPage(R1C8File, 1);
                var R1C8PDF = writer.GetImportedPage(R1C8File, 1);
                var R1C8 = new System.Drawing.Drawing2D.Matrix();
                R1C8.Translate(765f, 0f);
                R1C8.Rotate(90);
                writer.DirectContent.AddTemplate(R1C8Page, R1C8);

                PdfReader R1C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R1C9Page = writer.GetImportedPage(R1C9File, 1);
                var R1C9PDF = writer.GetImportedPage(R1C9File, 1);
                var R1C9 = new System.Drawing.Drawing2D.Matrix();
                R1C9.Translate(855f, 0f);
                R1C9.Rotate(90);
                writer.DirectContent.AddTemplate(R1C9Page, R1C9);


                //Row 2
                PdfReader R2C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(135f, 234f);
                R2C1.Rotate(90);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(225f, 234f);
                R2C2.Rotate(90);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(315f, 234f);
                R2C3.Rotate(90);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                PdfReader R2C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R2C4Page = writer.GetImportedPage(R2C4File, 1);
                var R2C4PDF = writer.GetImportedPage(R2C4File, 1);
                var R2C4 = new System.Drawing.Drawing2D.Matrix();
                R2C4.Translate(405f, 234f);
                R2C4.Rotate(90);
                writer.DirectContent.AddTemplate(R2C4Page, R2C4);

                PdfReader R2C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R2C5Page = writer.GetImportedPage(R2C5File, 1);
                var R2C5PDF = writer.GetImportedPage(R2C5File, 1);
                var R2C5 = new System.Drawing.Drawing2D.Matrix();
                R2C5.Translate(495f, 234f);
                R2C5.Rotate(90);
                writer.DirectContent.AddTemplate(R2C5Page, R2C5);

                PdfReader R2C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R2C6Page = writer.GetImportedPage(R2C6File, 1);
                var R2C6PDF = writer.GetImportedPage(R2C6File, 1);
                var R2C6 = new System.Drawing.Drawing2D.Matrix();
                R2C6.Translate(585f, 234f);
                R2C6.Rotate(90);
                writer.DirectContent.AddTemplate(R2C6Page, R2C6);

                PdfReader R2C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R2C7Page = writer.GetImportedPage(R2C7File, 1);
                var R2C7PDF = writer.GetImportedPage(R2C7File, 1);
                var R2C7 = new System.Drawing.Drawing2D.Matrix();
                R2C7.Translate(675f, 234f);
                R2C7.Rotate(90);
                writer.DirectContent.AddTemplate(R2C7Page, R2C7);

                PdfReader R2C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R2C8Page = writer.GetImportedPage(R2C8File, 1);
                var R2C8PDF = writer.GetImportedPage(R2C8File, 1);
                var R2C8 = new System.Drawing.Drawing2D.Matrix();
                R2C8.Translate(765f, 234f);
                R2C8.Rotate(90);
                writer.DirectContent.AddTemplate(R2C8Page, R2C8);

                PdfReader R2C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R2C9Page = writer.GetImportedPage(R2C9File, 1);
                var R2C9PDF = writer.GetImportedPage(R2C9File, 1);
                var R2C9 = new System.Drawing.Drawing2D.Matrix();
                R2C9.Translate(855f, 234f);
                R2C9.Rotate(90);
                writer.DirectContent.AddTemplate(R2C9Page, R2C9);


                //Row 3
                PdfReader R3C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(135f, 468f);
                R3C1.Rotate(90);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(225f, 468f);
                R3C2.Rotate(90);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(315f, 468f);
                R3C3.Rotate(90);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                PdfReader R3C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R3C4Page = writer.GetImportedPage(R3C4File, 1);
                var R3C4PDF = writer.GetImportedPage(R3C4File, 1);
                var R3C4 = new System.Drawing.Drawing2D.Matrix();
                R3C4.Translate(405f, 468f);
                R3C4.Rotate(90);
                writer.DirectContent.AddTemplate(R3C4Page, R3C4);

                PdfReader R3C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R3C5Page = writer.GetImportedPage(R3C5File, 1);
                var R3C5PDF = writer.GetImportedPage(R3C5File, 1);
                var R3C5 = new System.Drawing.Drawing2D.Matrix();
                R3C5.Translate(495f, 468f);
                R3C5.Rotate(90);
                writer.DirectContent.AddTemplate(R3C5Page, R3C5);

                PdfReader R3C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R3C6Page = writer.GetImportedPage(R3C6File, 1);
                var R3C6PDF = writer.GetImportedPage(R3C6File, 1);
                var R3C6 = new System.Drawing.Drawing2D.Matrix();
                R3C6.Translate(585f, 468f);
                R3C6.Rotate(90);
                writer.DirectContent.AddTemplate(R3C6Page, R3C6);

                PdfReader R3C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R3C7Page = writer.GetImportedPage(R3C7File, 1);
                var R3C7PDF = writer.GetImportedPage(R3C7File, 1);
                var R3C7 = new System.Drawing.Drawing2D.Matrix();
                R3C7.Translate(675f, 468f);
                R3C7.Rotate(90);
                writer.DirectContent.AddTemplate(R3C7Page, R3C7);

                PdfReader R3C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R3C8Page = writer.GetImportedPage(R3C8File, 1);
                var R3C8PDF = writer.GetImportedPage(R3C8File, 1);
                var R3C8 = new System.Drawing.Drawing2D.Matrix();
                R3C8.Translate(765f, 468f);
                R3C8.Rotate(90);
                writer.DirectContent.AddTemplate(R3C8Page, R3C8);

                PdfReader R3C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R3C9Page = writer.GetImportedPage(R3C9File, 1);
                var R3C9PDF = writer.GetImportedPage(R3C9File, 1);
                var R3C9 = new System.Drawing.Drawing2D.Matrix();
                R3C9.Translate(855f, 468f);
                R3C9.Rotate(90);
                writer.DirectContent.AddTemplate(R3C9Page, R3C9);


                //Row 4
                PdfReader R4C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(135f, 702f);
                R4C1.Rotate(90);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(225f, 702f);
                R4C2.Rotate(90);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(315f, 702f);
                R4C3.Rotate(90);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);

                PdfReader R4C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R4C4Page = writer.GetImportedPage(R4C4File, 1);
                var R4C4PDF = writer.GetImportedPage(R4C4File, 1);
                var R4C4 = new System.Drawing.Drawing2D.Matrix();
                R4C4.Translate(405f, 702f);
                R4C4.Rotate(90);
                writer.DirectContent.AddTemplate(R4C4Page, R4C4);

                PdfReader R4C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R4C5Page = writer.GetImportedPage(R4C5File, 1);
                var R4C5PDF = writer.GetImportedPage(R4C5File, 1);
                var R4C5 = new System.Drawing.Drawing2D.Matrix();
                R4C5.Translate(495f, 702f);
                R4C5.Rotate(90);
                writer.DirectContent.AddTemplate(R4C5Page, R4C5);

                PdfReader R4C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R4C6Page = writer.GetImportedPage(R4C6File, 1);
                var R4C6PDF = writer.GetImportedPage(R4C6File, 1);
                var R4C6 = new System.Drawing.Drawing2D.Matrix();
                R4C6.Translate(585f, 702f);
                R4C6.Rotate(90);
                writer.DirectContent.AddTemplate(R4C6Page, R4C6);

                PdfReader R4C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R4C7Page = writer.GetImportedPage(R4C7File, 1);
                var R4C7PDF = writer.GetImportedPage(R4C7File, 1);
                var R4C7 = new System.Drawing.Drawing2D.Matrix();
                R4C7.Translate(675f, 702f);
                R4C7.Rotate(90);
                writer.DirectContent.AddTemplate(R4C7Page, R4C7);

                PdfReader R4C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R4C8Page = writer.GetImportedPage(R4C8File, 1);
                var R4C8PDF = writer.GetImportedPage(R4C8File, 1);
                var R4C8 = new System.Drawing.Drawing2D.Matrix();
                R4C8.Translate(765f, 702f);
                R4C8.Rotate(90);
                writer.DirectContent.AddTemplate(R4C8Page, R4C8);

                PdfReader R4C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R4C9Page = writer.GetImportedPage(R4C9File, 1);
                var R4C9PDF = writer.GetImportedPage(R4C9File, 1);
                var R4C9 = new System.Drawing.Drawing2D.Matrix();
                R4C9.Translate(855f, 702f);
                R4C9.Rotate(90);
                writer.DirectContent.AddTemplate(R4C9Page, R4C9);


                //Row 5
                PdfReader R5C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(135f, 936f);
                R5C1.Rotate(90);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(225f, 936f);
                R5C2.Rotate(90);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                PdfReader R5C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R5C3Page = writer.GetImportedPage(R5C3File, 1);
                var R5C3PDF = writer.GetImportedPage(R5C3File, 1);
                var R5C3 = new System.Drawing.Drawing2D.Matrix();
                R5C3.Translate(315f, 936f);
                R5C3.Rotate(90);
                writer.DirectContent.AddTemplate(R5C3Page, R5C3);

                PdfReader R5C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R5C4Page = writer.GetImportedPage(R5C4File, 1);
                var R5C4PDF = writer.GetImportedPage(R5C4File, 1);
                var R5C4 = new System.Drawing.Drawing2D.Matrix();
                R5C4.Translate(405f, 936f);
                R5C4.Rotate(90);
                writer.DirectContent.AddTemplate(R5C4Page, R5C4);

                PdfReader R5C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R5C5Page = writer.GetImportedPage(R5C5File, 1);
                var R5C5PDF = writer.GetImportedPage(R5C5File, 1);
                var R5C5 = new System.Drawing.Drawing2D.Matrix();
                R5C5.Translate(495f, 936f);
                R5C5.Rotate(90);
                writer.DirectContent.AddTemplate(R5C5Page, R5C5);

                PdfReader R5C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R5C6Page = writer.GetImportedPage(R5C6File, 1);
                var R5C6PDF = writer.GetImportedPage(R5C6File, 1);
                var R5C6 = new System.Drawing.Drawing2D.Matrix();
                R5C6.Translate(585f, 936f);
                R5C6.Rotate(90);
                writer.DirectContent.AddTemplate(R5C6Page, R5C6);

                PdfReader R5C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R5C7Page = writer.GetImportedPage(R5C7File, 1);
                var R5C7PDF = writer.GetImportedPage(R5C7File, 1);
                var R5C7 = new System.Drawing.Drawing2D.Matrix();
                R5C7.Translate(675f, 936f);
                R5C7.Rotate(90);
                writer.DirectContent.AddTemplate(R5C7Page, R5C7);

                PdfReader R5C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R5C8Page = writer.GetImportedPage(R5C8File, 1);
                var R5C8PDF = writer.GetImportedPage(R5C8File, 1);
                var R5C8 = new System.Drawing.Drawing2D.Matrix();
                R5C8.Translate(765f, 936f);
                R5C8.Rotate(90);
                writer.DirectContent.AddTemplate(R5C8Page, R5C8);

                PdfReader R5C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R5C9Page = writer.GetImportedPage(R5C9File, 1);
                var R5C9PDF = writer.GetImportedPage(R5C9File, 1);
                var R5C9 = new System.Drawing.Drawing2D.Matrix();
                R5C9.Translate(855f, 936f);
                R5C9.Rotate(90);
                writer.DirectContent.AddTemplate(R5C9Page, R5C9);


                //Row 6
                PdfReader R6C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R6C1Page = writer.GetImportedPage(R6C1File, 1);
                var R6C1PDF = writer.GetImportedPage(R6C1File, 1);
                var R6C1 = new System.Drawing.Drawing2D.Matrix();
                R6C1.Translate(135f, 1170f);
                R6C1.Rotate(90);
                writer.DirectContent.AddTemplate(R6C1Page, R6C1);

                PdfReader R6C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R6C2Page = writer.GetImportedPage(R6C2File, 1);
                var R6C2PDF = writer.GetImportedPage(R6C2File, 1);
                var R6C2 = new System.Drawing.Drawing2D.Matrix();
                R6C2.Translate(225f, 1170f);
                R6C2.Rotate(90);
                writer.DirectContent.AddTemplate(R6C2Page, R6C2);

                PdfReader R6C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R6C3Page = writer.GetImportedPage(R6C3File, 1);
                var R6C3PDF = writer.GetImportedPage(R6C3File, 1);
                var R6C3 = new System.Drawing.Drawing2D.Matrix();
                R6C3.Translate(315f, 1170f);
                R6C3.Rotate(90);
                writer.DirectContent.AddTemplate(R6C3Page, R6C3);

                PdfReader R6C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R6C4Page = writer.GetImportedPage(R6C4File, 1);
                var R6C4PDF = writer.GetImportedPage(R6C4File, 1);
                var R6C4 = new System.Drawing.Drawing2D.Matrix();
                R6C4.Translate(405f, 1170f);
                R6C4.Rotate(90);
                writer.DirectContent.AddTemplate(R6C4Page, R6C4);

                PdfReader R6C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R6C5Page = writer.GetImportedPage(R6C5File, 1);
                var R6C5PDF = writer.GetImportedPage(R6C5File, 1);
                var R6C5 = new System.Drawing.Drawing2D.Matrix();
                R6C5.Translate(495f, 1170f);
                R6C5.Rotate(90);
                writer.DirectContent.AddTemplate(R6C5Page, R6C5);

                PdfReader R6C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R6C6Page = writer.GetImportedPage(R6C6File, 1);
                var R6C6PDF = writer.GetImportedPage(R6C6File, 1);
                var R6C6 = new System.Drawing.Drawing2D.Matrix();
                R6C6.Translate(585f, 1170f);
                R6C6.Rotate(90);
                writer.DirectContent.AddTemplate(R6C6Page, R6C6);

                PdfReader R6C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R6C7Page = writer.GetImportedPage(R6C7File, 1);
                var R6C7PDF = writer.GetImportedPage(R6C7File, 1);
                var R6C7 = new System.Drawing.Drawing2D.Matrix();
                R6C7.Translate(675f, 1170f);
                R6C7.Rotate(90);
                writer.DirectContent.AddTemplate(R6C7Page, R6C7);

                PdfReader R6C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R6C8Page = writer.GetImportedPage(R6C8File, 1);
                var R6C8PDF = writer.GetImportedPage(R6C8File, 1);
                var R6C8 = new System.Drawing.Drawing2D.Matrix();
                R6C8.Translate(765f, 1170f);
                R6C8.Rotate(90);
                writer.DirectContent.AddTemplate(R6C8Page, R6C8);

                PdfReader R6C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R6C9Page = writer.GetImportedPage(R6C9File, 1);
                var R6C9PDF = writer.GetImportedPage(R6C9File, 1);
                var R6C9 = new System.Drawing.Drawing2D.Matrix();
                R6C9.Translate(855f, 1170f);
                R6C9.Rotate(90);
                writer.DirectContent.AddTemplate(R6C9Page, R6C9);


                //Row 7
                PdfReader R7C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R7C1Page = writer.GetImportedPage(R7C1File, 1);
                var R7C1PDF = writer.GetImportedPage(R7C1File, 1);
                var R7C1 = new System.Drawing.Drawing2D.Matrix();
                R7C1.Translate(135f, 1404f);
                R7C1.Rotate(90);
                writer.DirectContent.AddTemplate(R7C1Page, R7C1);

                PdfReader R7C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R7C2Page = writer.GetImportedPage(R7C2File, 1);
                var R7C2PDF = writer.GetImportedPage(R7C2File, 1);
                var R7C2 = new System.Drawing.Drawing2D.Matrix();
                R7C2.Translate(225f, 1404f);
                R7C2.Rotate(90);
                writer.DirectContent.AddTemplate(R7C2Page, R7C2);

                PdfReader R7C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R7C3Page = writer.GetImportedPage(R7C3File, 1);
                var R7C3PDF = writer.GetImportedPage(R7C3File, 1);
                var R7C3 = new System.Drawing.Drawing2D.Matrix();
                R7C3.Translate(315f, 1404f);
                R7C3.Rotate(90);
                writer.DirectContent.AddTemplate(R7C3Page, R7C3);

                PdfReader R7C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R7C4Page = writer.GetImportedPage(R7C4File, 1);
                var R7C4PDF = writer.GetImportedPage(R7C4File, 1);
                var R7C4 = new System.Drawing.Drawing2D.Matrix();
                R7C4.Translate(405f, 1404f);
                R7C4.Rotate(90);
                writer.DirectContent.AddTemplate(R7C4Page, R7C4);

                PdfReader R7C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R7C5Page = writer.GetImportedPage(R7C5File, 1);
                var R7C5PDF = writer.GetImportedPage(R7C5File, 1);
                var R7C5 = new System.Drawing.Drawing2D.Matrix();
                R7C5.Translate(495f, 1404f);
                R7C5.Rotate(90);
                writer.DirectContent.AddTemplate(R7C5Page, R7C5);

                PdfReader R7C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R7C6Page = writer.GetImportedPage(R7C6File, 1);
                var R7C6PDF = writer.GetImportedPage(R7C6File, 1);
                var R7C6 = new System.Drawing.Drawing2D.Matrix();
                R7C6.Translate(585f, 1404f);
                R7C6.Rotate(90);
                writer.DirectContent.AddTemplate(R7C6Page, R7C6);

                PdfReader R7C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R7C7Page = writer.GetImportedPage(R7C7File, 1);
                var R7C7PDF = writer.GetImportedPage(R7C7File, 1);
                var R7C7 = new System.Drawing.Drawing2D.Matrix();
                R7C7.Translate(675f, 1404f);
                R7C7.Rotate(90);
                writer.DirectContent.AddTemplate(R7C7Page, R7C7);

                PdfReader R7C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R7C8Page = writer.GetImportedPage(R7C8File, 1);
                var R7C8PDF = writer.GetImportedPage(R7C8File, 1);
                var R7C8 = new System.Drawing.Drawing2D.Matrix();
                R7C8.Translate(765f, 1404f);
                R7C8.Rotate(90);
                writer.DirectContent.AddTemplate(R7C8Page, R7C8);

                PdfReader R7C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R7C9Page = writer.GetImportedPage(R7C9File, 1);
                var R7C9PDF = writer.GetImportedPage(R7C9File, 1);
                var R7C9 = new System.Drawing.Drawing2D.Matrix();
                R7C9.Translate(855f, 1404f);
                R7C9.Rotate(90);
                writer.DirectContent.AddTemplate(R7C9Page, R7C9);

                //Row 8
                PdfReader R8C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R8C1Page = writer.GetImportedPage(R8C1File, 1);
                var R8C1PDF = writer.GetImportedPage(R8C1File, 1);
                var R8C1 = new System.Drawing.Drawing2D.Matrix();
                R8C1.Translate(135f, 1638f);
                R8C1.Rotate(90);
                writer.DirectContent.AddTemplate(R8C1Page, R8C1);

                PdfReader R8C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R8C2Page = writer.GetImportedPage(R8C2File, 1);
                var R8C2PDF = writer.GetImportedPage(R8C2File, 1);
                var R8C2 = new System.Drawing.Drawing2D.Matrix();
                R8C2.Translate(225f, 1638f);
                R8C2.Rotate(90);
                writer.DirectContent.AddTemplate(R8C2Page, R8C2);

                PdfReader R8C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R8C3Page = writer.GetImportedPage(R8C3File, 1);
                var R8C3PDF = writer.GetImportedPage(R8C3File, 1);
                var R8C3 = new System.Drawing.Drawing2D.Matrix();
                R8C3.Translate(315f, 1638f);
                R8C3.Rotate(90);
                writer.DirectContent.AddTemplate(R8C3Page, R8C3);

                PdfReader R8C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R8C4Page = writer.GetImportedPage(R8C4File, 1);
                var R8C4PDF = writer.GetImportedPage(R8C4File, 1);
                var R8C4 = new System.Drawing.Drawing2D.Matrix();
                R8C4.Translate(405f, 1638f);
                R8C4.Rotate(90);
                writer.DirectContent.AddTemplate(R8C4Page, R8C4);

                PdfReader R8C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R8C5Page = writer.GetImportedPage(R8C5File, 1);
                var R8C5PDF = writer.GetImportedPage(R8C5File, 1);
                var R8C5 = new System.Drawing.Drawing2D.Matrix();
                R8C5.Translate(495f, 1638f);
                R8C5.Rotate(90);
                writer.DirectContent.AddTemplate(R8C5Page, R8C5);

                PdfReader R8C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R8C6Page = writer.GetImportedPage(R8C6File, 1);
                var R8C6PDF = writer.GetImportedPage(R8C6File, 1);
                var R8C6 = new System.Drawing.Drawing2D.Matrix();
                R8C6.Translate(585f, 1638f);
                R8C6.Rotate(90);
                writer.DirectContent.AddTemplate(R8C6Page, R8C6);

                PdfReader R8C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R8C7Page = writer.GetImportedPage(R8C7File, 1);
                var R8C7PDF = writer.GetImportedPage(R8C7File, 1);
                var R8C7 = new System.Drawing.Drawing2D.Matrix();
                R8C7.Translate(675f, 1638f);
                R8C7.Rotate(90);
                writer.DirectContent.AddTemplate(R8C7Page, R8C7);

                PdfReader R8C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R8C8Page = writer.GetImportedPage(R8C8File, 1);
                var R8C8PDF = writer.GetImportedPage(R8C8File, 1);
                var R8C8 = new System.Drawing.Drawing2D.Matrix();
                R8C8.Translate(765f, 1638f);
                R8C8.Rotate(90);
                writer.DirectContent.AddTemplate(R8C8Page, R8C8);

                PdfReader R8C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R8C9Page = writer.GetImportedPage(R8C9File, 1);
                var R8C9PDF = writer.GetImportedPage(R8C9File, 1);
                var R8C9 = new System.Drawing.Drawing2D.Matrix();
                R8C9.Translate(855f, 1638f);
                R8C9.Rotate(90);
                writer.DirectContent.AddTemplate(R8C9Page, R8C9);

                //Row 9
                PdfReader R9C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R9C1Page = writer.GetImportedPage(R9C1File, 1);
                var R9C1PDF = writer.GetImportedPage(R9C1File, 1);
                var R9C1 = new System.Drawing.Drawing2D.Matrix();
                R9C1.Translate(135f, 1872f);
                R9C1.Rotate(90);
                writer.DirectContent.AddTemplate(R9C1Page, R9C1);

                PdfReader R9C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R9C2Page = writer.GetImportedPage(R9C2File, 1);
                var R9C2PDF = writer.GetImportedPage(R9C2File, 1);
                var R9C2 = new System.Drawing.Drawing2D.Matrix();
                R9C2.Translate(225f, 1872f);
                R9C2.Rotate(90);
                writer.DirectContent.AddTemplate(R9C2Page, R9C2);

                PdfReader R9C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R9C3Page = writer.GetImportedPage(R9C3File, 1);
                var R9C3PDF = writer.GetImportedPage(R9C3File, 1);
                var R9C3 = new System.Drawing.Drawing2D.Matrix();
                R9C3.Translate(315f, 1872f);
                R9C3.Rotate(90);
                writer.DirectContent.AddTemplate(R9C3Page, R9C3);

                PdfReader R9C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R9C4Page = writer.GetImportedPage(R9C4File, 1);
                var R9C4PDF = writer.GetImportedPage(R9C4File, 1);
                var R9C4 = new System.Drawing.Drawing2D.Matrix();
                R9C4.Translate(405f, 1872f);
                R9C4.Rotate(90);
                writer.DirectContent.AddTemplate(R9C4Page, R9C4);

                PdfReader R9C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R9C5Page = writer.GetImportedPage(R9C5File, 1);
                var R9C5PDF = writer.GetImportedPage(R9C5File, 1);
                var R9C5 = new System.Drawing.Drawing2D.Matrix();
                R9C5.Translate(495f, 1872f);
                R9C5.Rotate(90);
                writer.DirectContent.AddTemplate(R9C5Page, R9C5);

                PdfReader R9C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R9C6Page = writer.GetImportedPage(R9C6File, 1);
                var R9C6PDF = writer.GetImportedPage(R9C6File, 1);
                var R9C6 = new System.Drawing.Drawing2D.Matrix();
                R9C6.Translate(585f, 1872f);
                R9C6.Rotate(90);
                writer.DirectContent.AddTemplate(R9C6Page, R9C6);

                PdfReader R9C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R9C7Page = writer.GetImportedPage(R9C7File, 1);
                var R9C7PDF = writer.GetImportedPage(R9C7File, 1);
                var R9C7 = new System.Drawing.Drawing2D.Matrix();
                R9C7.Translate(675f, 1872f);
                R9C7.Rotate(90);
                writer.DirectContent.AddTemplate(R9C7Page, R9C7);

                PdfReader R9C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R9C8Page = writer.GetImportedPage(R9C8File, 1);
                var R9C8PDF = writer.GetImportedPage(R9C8File, 1);
                var R9C8 = new System.Drawing.Drawing2D.Matrix();
                R9C8.Translate(765f, 1872f);
                R9C8.Rotate(90);
                writer.DirectContent.AddTemplate(R9C8Page, R9C8);

                PdfReader R9C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R9C9Page = writer.GetImportedPage(R9C9File, 1);
                var R9C9PDF = writer.GetImportedPage(R9C9File, 1);
                var R9C9 = new System.Drawing.Drawing2D.Matrix();
                R9C9.Translate(855f, 1872f);
                R9C9.Rotate(90);
                writer.DirectContent.AddTemplate(R9C9Page, R9C9);

                //Row 10
                PdfReader R10C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R10C1Page = writer.GetImportedPage(R10C1File, 1);
                var R10C1PDF = writer.GetImportedPage(R10C1File, 1);
                var R10C1 = new System.Drawing.Drawing2D.Matrix();
                R10C1.Translate(135f, 2106f);
                R10C1.Rotate(90);
                writer.DirectContent.AddTemplate(R10C1Page, R10C1);

                PdfReader R10C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R10C2Page = writer.GetImportedPage(R10C2File, 1);
                var R10C2PDF = writer.GetImportedPage(R10C2File, 1);
                var R10C2 = new System.Drawing.Drawing2D.Matrix();
                R10C2.Translate(225f, 2106f);
                R10C2.Rotate(90);
                writer.DirectContent.AddTemplate(R10C2Page, R10C2);

                PdfReader R10C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R10C3Page = writer.GetImportedPage(R10C3File, 1);
                var R10C3PDF = writer.GetImportedPage(R10C3File, 1);
                var R10C3 = new System.Drawing.Drawing2D.Matrix();
                R10C3.Translate(315f, 2106f);
                R10C3.Rotate(90);
                writer.DirectContent.AddTemplate(R10C3Page, R10C3);

                PdfReader R10C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R10C4Page = writer.GetImportedPage(R10C4File, 1);
                var R10C4PDF = writer.GetImportedPage(R10C4File, 1);
                var R10C4 = new System.Drawing.Drawing2D.Matrix();
                R10C4.Translate(405f, 2106f);
                R10C4.Rotate(90);
                writer.DirectContent.AddTemplate(R10C4Page, R10C4);

                PdfReader R10C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R10C5Page = writer.GetImportedPage(R10C5File, 1);
                var R10C5PDF = writer.GetImportedPage(R10C5File, 1);
                var R10C5 = new System.Drawing.Drawing2D.Matrix();
                R10C5.Translate(495f, 2106f);
                R10C5.Rotate(90);
                writer.DirectContent.AddTemplate(R10C5Page, R10C5);

                PdfReader R10C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R10C6Page = writer.GetImportedPage(R10C6File, 1);
                var R10C6PDF = writer.GetImportedPage(R10C6File, 1);
                var R10C6 = new System.Drawing.Drawing2D.Matrix();
                R10C6.Translate(585f, 2106f);
                R10C6.Rotate(90);
                writer.DirectContent.AddTemplate(R10C6Page, R10C6);

                PdfReader R10C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R10C7Page = writer.GetImportedPage(R10C7File, 1);
                var R10C7PDF = writer.GetImportedPage(R10C7File, 1);
                var R10C7 = new System.Drawing.Drawing2D.Matrix();
                R10C7.Translate(675f, 2106f);
                R10C7.Rotate(90);
                writer.DirectContent.AddTemplate(R10C7Page, R10C7);

                PdfReader R10C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                PdfImportedPage R10C8Page = writer.GetImportedPage(R10C8File, 1);
                var R10C8PDF = writer.GetImportedPage(R10C8File, 1);
                var R10C8 = new System.Drawing.Drawing2D.Matrix();
                R10C8.Translate(765f, 2106f);
                R10C8.Rotate(90);
                writer.DirectContent.AddTemplate(R10C8Page, R10C8);

                PdfReader R10C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                PdfImportedPage R10C9Page = writer.GetImportedPage(R10C9File, 1);
                var R10C9PDF = writer.GetImportedPage(R10C9File, 1);
                var R10C9 = new System.Drawing.Drawing2D.Matrix();
                R10C9.Translate(855f, 2106f);
                R10C9.Rotate(90);
                writer.DirectContent.AddTemplate(R10C9Page, R10C9);


                itemTotal.RemoveRange(0, 9);

                cb.SetLineWidth(18f);

                //Cropmarks Horizontal
                cb.MoveTo(27f, 0);
                cb.LineTo(873f, 0);
                cb.Stroke();

                cb.MoveTo(27f, 234);
                cb.LineTo(873f, 234);
                cb.Stroke();

                cb.MoveTo(27f, 468);
                cb.LineTo(873f, 468);
                cb.Stroke();

                cb.MoveTo(27f, 702);
                cb.LineTo(873f, 702);
                cb.Stroke();

                cb.MoveTo(27f, 936);
                cb.LineTo(873f, 936);
                cb.Stroke();

                cb.MoveTo(27f, 1170);
                cb.LineTo(873f, 1170);
                cb.Stroke();

                cb.MoveTo(27f, 1404);
                cb.LineTo(873f, 1404);
                cb.Stroke();

                cb.MoveTo(27f, 1638);
                cb.LineTo(873f, 1638);
                cb.Stroke();

                cb.MoveTo(27f, 1872);
                cb.LineTo(873f, 1872);
                cb.Stroke();

                cb.MoveTo(27f, 2106);
                cb.LineTo(873f, 2106);
                cb.Stroke();

                cb.MoveTo(27f, 2340);
                cb.LineTo(873f, 2340);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(45f, 0);
                cb.LineTo(855f, 0);
                cb.LineTo(855f, 2340);
                cb.LineTo(45f, 2340);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf2x2_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(162f, 162f));
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
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2754));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[4]);
                        itemPrint.RemoveRange(0, 5);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 17);
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
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 85);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(45f, 0f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(207f, 0f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(369f, 0f);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                var R1C4 = new System.Drawing.Drawing2D.Matrix();
                R1C4.Translate(531f, 0f);
                writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                PdfReader R1C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R1C5Page = writer.GetImportedPage(R1C5File, 1);
                var R1C5PDF = writer.GetImportedPage(R1C5File, 1);
                var R1C5 = new System.Drawing.Drawing2D.Matrix();
                R1C5.Translate(693f, 0f);
                writer.DirectContent.AddTemplate(R1C5Page, R1C5);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(45f, 162f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(207f, 162f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(369f, 162f);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                PdfReader R2C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R2C4Page = writer.GetImportedPage(R2C4File, 1);
                var R2C4PDF = writer.GetImportedPage(R2C4File, 1);
                var R2C4 = new System.Drawing.Drawing2D.Matrix();
                R2C4.Translate(531f, 162f);
                writer.DirectContent.AddTemplate(R2C4Page, R2C4);

                PdfReader R2C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R2C5Page = writer.GetImportedPage(R2C5File, 1);
                var R2C5PDF = writer.GetImportedPage(R2C5File, 1);
                var R2C5 = new System.Drawing.Drawing2D.Matrix();
                R2C5.Translate(693f, 162f);
                writer.DirectContent.AddTemplate(R2C5Page, R2C5);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(45f, 324f);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(207f, 324f);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(369f, 324f);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                PdfReader R3C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R3C4Page = writer.GetImportedPage(R3C4File, 1);
                var R3C4PDF = writer.GetImportedPage(R3C4File, 1);
                var R3C4 = new System.Drawing.Drawing2D.Matrix();
                R3C4.Translate(531f, 324f);
                writer.DirectContent.AddTemplate(R3C4Page, R3C4);

                PdfReader R3C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R3C5Page = writer.GetImportedPage(R3C5File, 1);
                var R3C5PDF = writer.GetImportedPage(R3C5File, 1);
                var R3C5 = new System.Drawing.Drawing2D.Matrix();
                R3C5.Translate(693f, 324f);
                writer.DirectContent.AddTemplate(R3C5Page, R3C5);

                //Row 4
                PdfReader R4C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(45f, 486f);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(207f, 486f);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(369f, 486f);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);

                PdfReader R4C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R4C4Page = writer.GetImportedPage(R4C4File, 1);
                var R4C4PDF = writer.GetImportedPage(R4C4File, 1);
                var R4C4 = new System.Drawing.Drawing2D.Matrix();
                R4C4.Translate(531f, 486f);
                writer.DirectContent.AddTemplate(R4C4Page, R4C4);

                PdfReader R4C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R4C5Page = writer.GetImportedPage(R4C5File, 1);
                var R4C5PDF = writer.GetImportedPage(R4C5File, 1);
                var R4C5 = new System.Drawing.Drawing2D.Matrix();
                R4C5.Translate(693f, 486f);
                writer.DirectContent.AddTemplate(R4C5Page, R4C5);

                //Row 5
                PdfReader R5C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(45f, 648f);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(207f, 648f);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                PdfReader R5C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R5C3Page = writer.GetImportedPage(R5C3File, 1);
                var R5C3PDF = writer.GetImportedPage(R5C3File, 1);
                var R5C3 = new System.Drawing.Drawing2D.Matrix();
                R5C3.Translate(369f, 648f);
                writer.DirectContent.AddTemplate(R5C3Page, R5C3);

                PdfReader R5C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R5C4Page = writer.GetImportedPage(R5C4File, 1);
                var R5C4PDF = writer.GetImportedPage(R5C4File, 1);
                var R5C4 = new System.Drawing.Drawing2D.Matrix();
                R5C4.Translate(531f, 648f);
                writer.DirectContent.AddTemplate(R5C4Page, R5C4);

                PdfReader R5C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R5C5Page = writer.GetImportedPage(R5C5File, 1);
                var R5C5PDF = writer.GetImportedPage(R5C5File, 1);
                var R5C5 = new System.Drawing.Drawing2D.Matrix();
                R5C5.Translate(693f, 648f);
                writer.DirectContent.AddTemplate(R5C5Page, R5C5);

                //Row 6
                PdfReader R6C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R6C1Page = writer.GetImportedPage(R6C1File, 1);
                var R6C1PDF = writer.GetImportedPage(R6C1File, 1);
                var R6C1 = new System.Drawing.Drawing2D.Matrix();
                R6C1.Translate(45f, 810f);
                writer.DirectContent.AddTemplate(R6C1Page, R6C1);

                PdfReader R6C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R6C2Page = writer.GetImportedPage(R6C2File, 1);
                var R6C2PDF = writer.GetImportedPage(R6C2File, 1);
                var R6C2 = new System.Drawing.Drawing2D.Matrix();
                R6C2.Translate(207f, 810f);
                writer.DirectContent.AddTemplate(R6C2Page, R6C2);

                PdfReader R6C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R6C3Page = writer.GetImportedPage(R6C3File, 1);
                var R6C3PDF = writer.GetImportedPage(R6C3File, 1);
                var R6C3 = new System.Drawing.Drawing2D.Matrix();
                R6C3.Translate(369f, 810f);
                writer.DirectContent.AddTemplate(R6C3Page, R6C3);

                PdfReader R6C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R6C4Page = writer.GetImportedPage(R6C4File, 1);
                var R6C4PDF = writer.GetImportedPage(R6C4File, 1);
                var R6C4 = new System.Drawing.Drawing2D.Matrix();
                R6C4.Translate(531f, 810f);
                writer.DirectContent.AddTemplate(R6C4Page, R6C4);

                PdfReader R6C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R6C5Page = writer.GetImportedPage(R6C5File, 1);
                var R6C5PDF = writer.GetImportedPage(R6C5File, 1);
                var R6C5 = new System.Drawing.Drawing2D.Matrix();
                R6C5.Translate(693f, 810f);
                writer.DirectContent.AddTemplate(R6C5Page, R6C5);

                //Row 7
                PdfReader R7C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R7C1Page = writer.GetImportedPage(R7C1File, 1);
                var R7C1PDF = writer.GetImportedPage(R7C1File, 1);
                var R7C1 = new System.Drawing.Drawing2D.Matrix();
                R7C1.Translate(45f, 972f);
                writer.DirectContent.AddTemplate(R7C1Page, R7C1);

                PdfReader R7C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R7C2Page = writer.GetImportedPage(R7C2File, 1);
                var R7C2PDF = writer.GetImportedPage(R7C2File, 1);
                var R7C2 = new System.Drawing.Drawing2D.Matrix();
                R7C2.Translate(207f, 972f);
                writer.DirectContent.AddTemplate(R7C2Page, R7C2);

                PdfReader R7C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R7C3Page = writer.GetImportedPage(R7C3File, 1);
                var R7C3PDF = writer.GetImportedPage(R7C3File, 1);
                var R7C3 = new System.Drawing.Drawing2D.Matrix();
                R7C3.Translate(369f, 972f);
                writer.DirectContent.AddTemplate(R7C3Page, R7C3);

                PdfReader R7C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R7C4Page = writer.GetImportedPage(R7C4File, 1);
                var R7C4PDF = writer.GetImportedPage(R7C4File, 1);
                var R7C4 = new System.Drawing.Drawing2D.Matrix();
                R7C4.Translate(531f, 972f);
                writer.DirectContent.AddTemplate(R7C4Page, R7C4);

                PdfReader R7C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R7C5Page = writer.GetImportedPage(R7C5File, 1);
                var R7C5PDF = writer.GetImportedPage(R7C5File, 1);
                var R7C5 = new System.Drawing.Drawing2D.Matrix();
                R7C5.Translate(693f, 972f);
                writer.DirectContent.AddTemplate(R7C5Page, R7C5);

                //Row 8
                PdfReader R8C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R8C1Page = writer.GetImportedPage(R8C1File, 1);
                var R8C1PDF = writer.GetImportedPage(R8C1File, 1);
                var R8C1 = new System.Drawing.Drawing2D.Matrix();
                R8C1.Translate(45f, 1134f);
                writer.DirectContent.AddTemplate(R8C1Page, R8C1);

                PdfReader R8C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R8C2Page = writer.GetImportedPage(R8C2File, 1);
                var R8C2PDF = writer.GetImportedPage(R8C2File, 1);
                var R8C2 = new System.Drawing.Drawing2D.Matrix();
                R8C2.Translate(207f, 1134f);
                writer.DirectContent.AddTemplate(R8C2Page, R8C2);

                PdfReader R8C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R8C3Page = writer.GetImportedPage(R8C3File, 1);
                var R8C3PDF = writer.GetImportedPage(R8C3File, 1);
                var R8C3 = new System.Drawing.Drawing2D.Matrix();
                R8C3.Translate(369f, 1134f);
                writer.DirectContent.AddTemplate(R8C3Page, R8C3);

                PdfReader R8C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R8C4Page = writer.GetImportedPage(R8C4File, 1);
                var R8C4PDF = writer.GetImportedPage(R8C4File, 1);
                var R8C4 = new System.Drawing.Drawing2D.Matrix();
                R8C4.Translate(531f, 1134f);
                writer.DirectContent.AddTemplate(R8C4Page, R8C4);

                PdfReader R8C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R8C5Page = writer.GetImportedPage(R8C5File, 1);
                var R8C5PDF = writer.GetImportedPage(R8C5File, 1);
                var R8C5 = new System.Drawing.Drawing2D.Matrix();
                R8C5.Translate(693f, 1134f);
                writer.DirectContent.AddTemplate(R8C5Page, R8C5);

                //Row 9
                PdfReader R9C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R9C1Page = writer.GetImportedPage(R9C1File, 1);
                var R9C1PDF = writer.GetImportedPage(R9C1File, 1);
                var R9C1 = new System.Drawing.Drawing2D.Matrix();
                R9C1.Translate(45f, 1296f);
                writer.DirectContent.AddTemplate(R9C1Page, R9C1);

                PdfReader R9C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R9C2Page = writer.GetImportedPage(R9C2File, 1);
                var R9C2PDF = writer.GetImportedPage(R9C2File, 1);
                var R9C2 = new System.Drawing.Drawing2D.Matrix();
                R9C2.Translate(207f, 1296f);
                writer.DirectContent.AddTemplate(R9C2Page, R9C2);

                PdfReader R9C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R9C3Page = writer.GetImportedPage(R9C3File, 1);
                var R9C3PDF = writer.GetImportedPage(R9C3File, 1);
                var R9C3 = new System.Drawing.Drawing2D.Matrix();
                R9C3.Translate(369f, 1296f);
                writer.DirectContent.AddTemplate(R9C3Page, R9C3);

                PdfReader R9C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R9C4Page = writer.GetImportedPage(R9C4File, 1);
                var R9C4PDF = writer.GetImportedPage(R9C4File, 1);
                var R9C4 = new System.Drawing.Drawing2D.Matrix();
                R9C4.Translate(531f, 1296f);
                writer.DirectContent.AddTemplate(R9C4Page, R9C4);

                PdfReader R9C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R9C5Page = writer.GetImportedPage(R9C5File, 1);
                var R9C5PDF = writer.GetImportedPage(R9C5File, 1);
                var R9C5 = new System.Drawing.Drawing2D.Matrix();
                R9C5.Translate(693f, 1296f);
                writer.DirectContent.AddTemplate(R9C5Page, R9C5);

                //Row 10
                PdfReader R10C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R10C1Page = writer.GetImportedPage(R10C1File, 1);
                var R10C1PDF = writer.GetImportedPage(R10C1File, 1);
                var R10C1 = new System.Drawing.Drawing2D.Matrix();
                R10C1.Translate(45f, 1458f);
                writer.DirectContent.AddTemplate(R10C1Page, R10C1);

                PdfReader R10C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R10C2Page = writer.GetImportedPage(R10C2File, 1);
                var R10C2PDF = writer.GetImportedPage(R10C2File, 1);
                var R10C2 = new System.Drawing.Drawing2D.Matrix();
                R10C2.Translate(207f, 1458f);
                writer.DirectContent.AddTemplate(R10C2Page, R10C2);

                PdfReader R10C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R10C3Page = writer.GetImportedPage(R10C3File, 1);
                var R10C3PDF = writer.GetImportedPage(R10C3File, 1);
                var R10C3 = new System.Drawing.Drawing2D.Matrix();
                R10C3.Translate(369f, 1458f);
                writer.DirectContent.AddTemplate(R10C3Page, R10C3);

                PdfReader R10C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R10C4Page = writer.GetImportedPage(R10C4File, 1);
                var R10C4PDF = writer.GetImportedPage(R10C4File, 1);
                var R10C4 = new System.Drawing.Drawing2D.Matrix();
                R10C4.Translate(531f, 1458f);
                writer.DirectContent.AddTemplate(R10C4Page, R10C4);

                PdfReader R10C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R10C5Page = writer.GetImportedPage(R10C5File, 1);
                var R10C5PDF = writer.GetImportedPage(R10C5File, 1);
                var R10C5 = new System.Drawing.Drawing2D.Matrix();
                R10C5.Translate(693f, 1458f);
                writer.DirectContent.AddTemplate(R10C5Page, R10C5);

                //Row 11
                PdfReader R11C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R11C1Page = writer.GetImportedPage(R11C1File, 1);
                var R11C1PDF = writer.GetImportedPage(R11C1File, 1);
                var R11C1 = new System.Drawing.Drawing2D.Matrix();
                R11C1.Translate(45f, 1620f);
                writer.DirectContent.AddTemplate(R11C1Page, R11C1);

                PdfReader R11C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R11C2Page = writer.GetImportedPage(R11C2File, 1);
                var R11C2PDF = writer.GetImportedPage(R11C2File, 1);
                var R11C2 = new System.Drawing.Drawing2D.Matrix();
                R11C2.Translate(207f, 1620f);
                writer.DirectContent.AddTemplate(R11C2Page, R11C2);

                PdfReader R11C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R11C3Page = writer.GetImportedPage(R11C3File, 1);
                var R11C3PDF = writer.GetImportedPage(R11C3File, 1);
                var R11C3 = new System.Drawing.Drawing2D.Matrix();
                R11C3.Translate(369f, 1620f);
                writer.DirectContent.AddTemplate(R11C3Page, R11C3);

                PdfReader R11C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R11C4Page = writer.GetImportedPage(R11C4File, 1);
                var R11C4PDF = writer.GetImportedPage(R11C4File, 1);
                var R11C4 = new System.Drawing.Drawing2D.Matrix();
                R11C4.Translate(531f, 1620f);
                writer.DirectContent.AddTemplate(R11C4Page, R11C4);

                PdfReader R11C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R11C5Page = writer.GetImportedPage(R11C5File, 1);
                var R11C5PDF = writer.GetImportedPage(R11C5File, 1);
                var R11C5 = new System.Drawing.Drawing2D.Matrix();
                R11C5.Translate(693f, 1620f);
                writer.DirectContent.AddTemplate(R11C5Page, R11C5);

                //Row 12
                PdfReader R12C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R12C1Page = writer.GetImportedPage(R12C1File, 1);
                var R12C1PDF = writer.GetImportedPage(R12C1File, 1);
                var R12C1 = new System.Drawing.Drawing2D.Matrix();
                R12C1.Translate(45f, 1782f);
                writer.DirectContent.AddTemplate(R12C1Page, R12C1);

                PdfReader R12C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R12C2Page = writer.GetImportedPage(R12C2File, 1);
                var R12C2PDF = writer.GetImportedPage(R12C2File, 1);
                var R12C2 = new System.Drawing.Drawing2D.Matrix();
                R12C2.Translate(207f, 1782f);
                writer.DirectContent.AddTemplate(R12C2Page, R12C2);

                PdfReader R12C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R12C3Page = writer.GetImportedPage(R12C3File, 1);
                var R12C3PDF = writer.GetImportedPage(R12C3File, 1);
                var R12C3 = new System.Drawing.Drawing2D.Matrix();
                R12C3.Translate(369f, 1782f);
                writer.DirectContent.AddTemplate(R12C3Page, R12C3);

                PdfReader R12C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R12C4Page = writer.GetImportedPage(R12C4File, 1);
                var R12C4PDF = writer.GetImportedPage(R12C4File, 1);
                var R12C4 = new System.Drawing.Drawing2D.Matrix();
                R12C4.Translate(531f, 1782f);
                writer.DirectContent.AddTemplate(R12C4Page, R12C4);

                PdfReader R12C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R12C5Page = writer.GetImportedPage(R12C5File, 1);
                var R12C5PDF = writer.GetImportedPage(R12C5File, 1);
                var R12C5 = new System.Drawing.Drawing2D.Matrix();
                R12C5.Translate(693f, 1782f);
                writer.DirectContent.AddTemplate(R12C5Page, R12C5);

                //Row 13
                PdfReader R13C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R13C1Page = writer.GetImportedPage(R13C1File, 1);
                var R13C1PDF = writer.GetImportedPage(R13C1File, 1);
                var R13C1 = new System.Drawing.Drawing2D.Matrix();
                R13C1.Translate(45f, 1944f);
                writer.DirectContent.AddTemplate(R13C1Page, R13C1);

                PdfReader R13C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R13C2Page = writer.GetImportedPage(R13C2File, 1);
                var R13C2PDF = writer.GetImportedPage(R13C2File, 1);
                var R13C2 = new System.Drawing.Drawing2D.Matrix();
                R13C2.Translate(207f, 1944f);
                writer.DirectContent.AddTemplate(R13C2Page, R13C2);

                PdfReader R13C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R13C3Page = writer.GetImportedPage(R13C3File, 1);
                var R13C3PDF = writer.GetImportedPage(R13C3File, 1);
                var R13C3 = new System.Drawing.Drawing2D.Matrix();
                R13C3.Translate(369f, 1944f);
                writer.DirectContent.AddTemplate(R13C3Page, R13C3);

                PdfReader R13C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R13C4Page = writer.GetImportedPage(R13C4File, 1);
                var R13C4PDF = writer.GetImportedPage(R13C4File, 1);
                var R13C4 = new System.Drawing.Drawing2D.Matrix();
                R13C4.Translate(531f, 1944f);
                writer.DirectContent.AddTemplate(R13C4Page, R13C4);

                PdfReader R13C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R13C5Page = writer.GetImportedPage(R13C5File, 1);
                var R13C5PDF = writer.GetImportedPage(R13C5File, 1);
                var R13C5 = new System.Drawing.Drawing2D.Matrix();
                R13C5.Translate(693f, 1944f);
                writer.DirectContent.AddTemplate(R13C5Page, R13C5);

                //Row 14
                PdfReader R14C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R14C1Page = writer.GetImportedPage(R14C1File, 1);
                var R14C1PDF = writer.GetImportedPage(R14C1File, 1);
                var R14C1 = new System.Drawing.Drawing2D.Matrix();
                R14C1.Translate(45f, 2106f);
                writer.DirectContent.AddTemplate(R14C1Page, R14C1);

                PdfReader R14C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R14C2Page = writer.GetImportedPage(R14C2File, 1);
                var R14C2PDF = writer.GetImportedPage(R14C2File, 1);
                var R14C2 = new System.Drawing.Drawing2D.Matrix();
                R14C2.Translate(207f, 2106f);
                writer.DirectContent.AddTemplate(R14C2Page, R14C2);

                PdfReader R14C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R14C3Page = writer.GetImportedPage(R14C3File, 1);
                var R14C3PDF = writer.GetImportedPage(R14C3File, 1);
                var R14C3 = new System.Drawing.Drawing2D.Matrix();
                R14C3.Translate(369f, 2106f);
                writer.DirectContent.AddTemplate(R14C3Page, R14C3);

                PdfReader R14C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R14C4Page = writer.GetImportedPage(R14C4File, 1);
                var R14C4PDF = writer.GetImportedPage(R14C4File, 1);
                var R14C4 = new System.Drawing.Drawing2D.Matrix();
                R14C4.Translate(531f, 2106f);
                writer.DirectContent.AddTemplate(R14C4Page, R14C4);

                PdfReader R14C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R14C5Page = writer.GetImportedPage(R14C5File, 1);
                var R14C5PDF = writer.GetImportedPage(R14C5File, 1);
                var R14C5 = new System.Drawing.Drawing2D.Matrix();
                R14C5.Translate(693f, 2106f);
                writer.DirectContent.AddTemplate(R14C5Page, R14C5);

                //Row 15
                PdfReader R15C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R15C1Page = writer.GetImportedPage(R15C1File, 1);
                var R15C1PDF = writer.GetImportedPage(R15C1File, 1);
                var R15C1 = new System.Drawing.Drawing2D.Matrix();
                R15C1.Translate(45f, 2268f);
                writer.DirectContent.AddTemplate(R15C1Page, R15C1);

                PdfReader R15C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R15C2Page = writer.GetImportedPage(R15C2File, 1);
                var R15C2PDF = writer.GetImportedPage(R15C2File, 1);
                var R15C2 = new System.Drawing.Drawing2D.Matrix();
                R15C2.Translate(207f, 2268f);
                writer.DirectContent.AddTemplate(R15C2Page, R15C2);

                PdfReader R15C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R15C3Page = writer.GetImportedPage(R15C3File, 1);
                var R15C3PDF = writer.GetImportedPage(R15C3File, 1);
                var R15C3 = new System.Drawing.Drawing2D.Matrix();
                R15C3.Translate(369f, 2268f);
                writer.DirectContent.AddTemplate(R15C3Page, R15C3);

                PdfReader R15C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R15C4Page = writer.GetImportedPage(R15C4File, 1);
                var R15C4PDF = writer.GetImportedPage(R15C4File, 1);
                var R15C4 = new System.Drawing.Drawing2D.Matrix();
                R15C4.Translate(531f, 2268f);
                writer.DirectContent.AddTemplate(R15C4Page, R15C4);

                PdfReader R15C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R15C5Page = writer.GetImportedPage(R15C5File, 1);
                var R15C5PDF = writer.GetImportedPage(R15C5File, 1);
                var R15C5 = new System.Drawing.Drawing2D.Matrix();
                R15C5.Translate(693f, 2268f);
                writer.DirectContent.AddTemplate(R15C5Page, R15C5);

                //Row 16
                PdfReader R16C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R16C1Page = writer.GetImportedPage(R16C1File, 1);
                var R16C1PDF = writer.GetImportedPage(R16C1File, 1);
                var R16C1 = new System.Drawing.Drawing2D.Matrix();
                R16C1.Translate(45f, 2430f);
                writer.DirectContent.AddTemplate(R16C1Page, R16C1);

                PdfReader R16C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R16C2Page = writer.GetImportedPage(R16C2File, 1);
                var R16C2PDF = writer.GetImportedPage(R16C2File, 1);
                var R16C2 = new System.Drawing.Drawing2D.Matrix();
                R16C2.Translate(207f, 2430f);
                writer.DirectContent.AddTemplate(R16C2Page, R16C2);

                PdfReader R16C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R16C3Page = writer.GetImportedPage(R16C3File, 1);
                var R16C3PDF = writer.GetImportedPage(R16C3File, 1);
                var R16C3 = new System.Drawing.Drawing2D.Matrix();
                R16C3.Translate(369f, 2430f);
                writer.DirectContent.AddTemplate(R16C3Page, R16C3);

                PdfReader R16C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R16C4Page = writer.GetImportedPage(R16C4File, 1);
                var R16C4PDF = writer.GetImportedPage(R16C4File, 1);
                var R16C4 = new System.Drawing.Drawing2D.Matrix();
                R16C4.Translate(531f, 2430f);
                writer.DirectContent.AddTemplate(R16C4Page, R16C4);

                PdfReader R16C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R16C5Page = writer.GetImportedPage(R16C5File, 1);
                var R16C5PDF = writer.GetImportedPage(R16C5File, 1);
                var R16C5 = new System.Drawing.Drawing2D.Matrix();
                R16C5.Translate(693f, 2430f);
                writer.DirectContent.AddTemplate(R16C5Page, R16C5);

                //Row 17
                PdfReader R17C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R17C1Page = writer.GetImportedPage(R17C1File, 1);
                var R17C1PDF = writer.GetImportedPage(R17C1File, 1);
                var R17C1 = new System.Drawing.Drawing2D.Matrix();
                R17C1.Translate(45f, 2592f);
                writer.DirectContent.AddTemplate(R17C1Page, R17C1);

                PdfReader R17C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R17C2Page = writer.GetImportedPage(R17C2File, 1);
                var R17C2PDF = writer.GetImportedPage(R17C2File, 1);
                var R17C2 = new System.Drawing.Drawing2D.Matrix();
                R17C2.Translate(207f, 2592f);
                writer.DirectContent.AddTemplate(R17C2Page, R17C2);

                PdfReader R17C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R17C3Page = writer.GetImportedPage(R17C3File, 1);
                var R17C3PDF = writer.GetImportedPage(R17C3File, 1);
                var R17C3 = new System.Drawing.Drawing2D.Matrix();
                R17C3.Translate(369f, 2592f);
                writer.DirectContent.AddTemplate(R17C3Page, R17C3);

                PdfReader R17C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R17C4Page = writer.GetImportedPage(R17C4File, 1);
                var R17C4PDF = writer.GetImportedPage(R17C4File, 1);
                var R17C4 = new System.Drawing.Drawing2D.Matrix();
                R17C4.Translate(531f, 2592f);
                writer.DirectContent.AddTemplate(R17C4Page, R17C4);

                PdfReader R17C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R17C5Page = writer.GetImportedPage(R17C5File, 1);
                var R17C5PDF = writer.GetImportedPage(R17C5File, 1);
                var R17C5 = new System.Drawing.Drawing2D.Matrix();
                R17C5.Translate(693f, 2592f);
                writer.DirectContent.AddTemplate(R17C5Page, R17C5);


                itemTotal.RemoveRange(0, 5);


                cb.SetLineWidth(18f);

                //Cropmarks Horizontal
                cb.MoveTo(0f, 0);
                cb.LineTo(900f, 0);
                cb.Stroke();

                cb.MoveTo(0f, 162);
                cb.LineTo(900f, 162);
                cb.Stroke();

                cb.MoveTo(0f, 324);
                cb.LineTo(900f, 324);
                cb.Stroke();

                cb.MoveTo(0f, 486);
                cb.LineTo(900f, 486);
                cb.Stroke();

                cb.MoveTo(0f, 648);
                cb.LineTo(900f, 648);
                cb.Stroke();

                cb.MoveTo(0f, 810);
                cb.LineTo(900f, 810);
                cb.Stroke();

                cb.MoveTo(0f, 972);
                cb.LineTo(900f, 972);
                cb.Stroke();

                cb.MoveTo(0f, 1134);
                cb.LineTo(900f, 1134);
                cb.Stroke();

                cb.MoveTo(0f, 1296);
                cb.LineTo(900f, 1296);
                cb.Stroke();

                cb.MoveTo(0f, 1458);
                cb.LineTo(900f, 1458);
                cb.Stroke();

                cb.MoveTo(0f, 1620);
                cb.LineTo(900f, 1620);
                cb.Stroke();

                cb.MoveTo(0f, 1782);
                cb.LineTo(900f, 1782);
                cb.Stroke();

                cb.MoveTo(0f, 1944);
                cb.LineTo(900f, 1944);
                cb.Stroke();

                cb.MoveTo(0f, 2106);
                cb.LineTo(900f, 2106);
                cb.Stroke();

                cb.MoveTo(0f, 2268);
                cb.LineTo(900f, 2268);
                cb.Stroke();

                cb.MoveTo(0f, 2430);
                cb.LineTo(900f, 2430);
                cb.Stroke();

                cb.MoveTo(0f, 2592);
                cb.LineTo(900f, 2592);
                cb.Stroke();

                cb.MoveTo(0f, 2754);
                cb.LineTo(900f, 2754);
                cb.Stroke();


                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27f, 0);
                cb.LineTo(873f, 0);
                cb.LineTo(873f, 2754);
                cb.LineTo(27f, 2754);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf2x2Circle_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(162f, 162f));
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
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2754));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    itemTotal.Add(itemPrint[0]);
                    itemTotal.Add(itemPrint[0]);
                    itemTotal.Add(itemPrint[0]);
                    itemTotal.Add(itemPrint[0]);
                    itemPrint.RemoveAt(0);
                    printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 68);
                    diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                    itemQtyPrint.RemoveAt(0);
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                float stepDistance = 0;
                for (int i = 1; i <= 17; i++)
                {

                    PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                    PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                    var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                    var R1C1 = new System.Drawing.Drawing2D.Matrix();
                    R1C1.Translate(126f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                    PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                    PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                    var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                    var R1C2 = new System.Drawing.Drawing2D.Matrix();
                    R1C2.Translate(288f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                    PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                    PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                    var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                    var R1C3 = new System.Drawing.Drawing2D.Matrix();
                    R1C3.Translate(450f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                    PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                    PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                    var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                    var R1C4 = new System.Drawing.Drawing2D.Matrix();
                    R1C4.Translate(612f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                    stepDistance = stepDistance + 162;

                }
                stepDistance = 0;

                itemTotal.RemoveRange(0, 4);

                cb.SetLineWidth(18f);

                for (int i = 1; i <= 18; i++)
                {
                    //Cropmarks Horizontal
                    cb.MoveTo(103.5f, stepDistance);
                    cb.LineTo(796.5f, stepDistance);
                    cb.Stroke();
                    stepDistance = stepDistance + 162;
                }

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(121.5f, 0);
                cb.LineTo(778.5f, 0);
                cb.LineTo(778.5f, 2754);
                cb.LineTo(121.5f, 2754);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf3x_5_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(234f, 54f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 54f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2106f));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[4]);
                        itemTotal.Add(itemPrint[4]);
                        itemPrint.RemoveRange(0, 5);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 18);
                        diffPerPage.Add("5 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 5);
                    }
                    else if (itemPrint.Count() % 2 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemPrint.RemoveRange(0, 2);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 45);
                        diffPerPage.Add("2 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 2);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 90);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                float stepDistance = 0;
                for (int i = 1; i <= 9; i++)
                {
                    PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                    PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                    var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                    var R1C1 = new System.Drawing.Drawing2D.Matrix();
                    R1C1.Translate(234f, stepDistance);
                    R1C1.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                    PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                    PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                    var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                    var R1C2 = new System.Drawing.Drawing2D.Matrix();
                    R1C2.Translate(288f, stepDistance);
                    R1C2.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                    PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                    PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                    var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                    var R1C3 = new System.Drawing.Drawing2D.Matrix();
                    R1C3.Translate(342f, stepDistance);
                    R1C3.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                    PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                    PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                    var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                    var R1C4 = new System.Drawing.Drawing2D.Matrix();
                    R1C4.Translate(396f, stepDistance);
                    R1C4.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                    PdfReader R1C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                    PdfImportedPage R1C5Page = writer.GetImportedPage(R1C5File, 1);
                    var R1C5PDF = writer.GetImportedPage(R1C5File, 1);
                    var R1C5 = new System.Drawing.Drawing2D.Matrix();
                    R1C5.Translate(450f, stepDistance);
                    R1C5.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C5Page, R1C5);

                    PdfReader R1C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                    PdfImportedPage R1C6Page = writer.GetImportedPage(R1C6File, 1);
                    var R1C6PDF = writer.GetImportedPage(R1C6File, 1);
                    var R1C6 = new System.Drawing.Drawing2D.Matrix();
                    R1C6.Translate(504f, stepDistance);
                    R1C6.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C6Page, R1C6);

                    PdfReader R1C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                    PdfImportedPage R1C7Page = writer.GetImportedPage(R1C7File, 1);
                    var R1C7PDF = writer.GetImportedPage(R1C7File, 1);
                    var R1C7 = new System.Drawing.Drawing2D.Matrix();
                    R1C7.Translate(558f, stepDistance);
                    R1C7.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C7Page, R1C7);

                    PdfReader R1C8File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[7]) + ".pdf");
                    PdfImportedPage R1C8Page = writer.GetImportedPage(R1C8File, 1);
                    var R1C8PDF = writer.GetImportedPage(R1C8File, 1);
                    var R1C8 = new System.Drawing.Drawing2D.Matrix();
                    R1C8.Translate(612f, stepDistance);
                    R1C8.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C8Page, R1C8);

                    PdfReader R1C9File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[8]) + ".pdf");
                    PdfImportedPage R1C9Page = writer.GetImportedPage(R1C9File, 1);
                    var R1C9PDF = writer.GetImportedPage(R1C9File, 1);
                    var R1C9 = new System.Drawing.Drawing2D.Matrix();
                    R1C9.Translate(666f, stepDistance);
                    R1C9.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C9Page, R1C9);

                    PdfReader R1C10File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[9]) + ".pdf");
                    PdfImportedPage R1C10Page = writer.GetImportedPage(R1C10File, 1);
                    var R1C10PDF = writer.GetImportedPage(R1C10File, 1);
                    var R1C10 = new System.Drawing.Drawing2D.Matrix();
                    R1C10.Translate(720f, stepDistance);
                    R1C10.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C10Page, R1C10);

                    stepDistance = stepDistance + 234;
                }
                stepDistance = 0;

                itemTotal.RemoveRange(0, 10);


                cb.SetLineWidth(18f);

                for (int i = 1; i <= 4; i++)
                {
                    //Cropmarks Horizontal
                    cb.MoveTo(153f, stepDistance);
                    cb.LineTo(747f, stepDistance);
                    cb.Stroke();
                    stepDistance = stepDistance + (234 * 3);
                }

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(171f, 0);
                cb.LineTo(729f, 0);
                cb.LineTo(729f, 2106f);
                cb.LineTo(171f, 2106f);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf3x2_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(225f, 162f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 162f)
                {
                    tm.Translate(-19.62f, -15.12f);
                }
                else
                {
                    tm.Translate(0f, 0f);
                }
                writer1.DirectContent.AddTemplate(imp, tm);
                doc1.Close();
            }

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2700));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 4 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemPrint.RemoveRange(0, 4);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 12);
                        diffPerPage.Add("4 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 4);

                    }
                    else if (itemPrint.Count() % 2 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[1]);
                        itemPrint.RemoveRange(0, 2);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 24);
                        diffPerPage.Add("2 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 2);

                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 48);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }

            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                float stepDistance = 0;
                for (int i = 1; i <= 12; i++)
                {
                    PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                    PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                    var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                    var R1C1 = new System.Drawing.Drawing2D.Matrix();
                    R1C1.Translate(288f, 0f + stepDistance);
                    R1C1.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                    PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                    PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                    var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                    var R1C2 = new System.Drawing.Drawing2D.Matrix();
                    R1C2.Translate(450f, 0f + stepDistance);
                    R1C2.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                    PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                    PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                    var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                    var R1C3 = new System.Drawing.Drawing2D.Matrix();
                    R1C3.Translate(612f, 0f + stepDistance);
                    R1C3.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                    PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                    PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                    var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                    var R1C4 = new System.Drawing.Drawing2D.Matrix();
                    R1C4.Translate(774f, 0f + stepDistance);
                    R1C4.Rotate(90);
                    writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                    stepDistance = stepDistance + 225;
                }
                stepDistance = 0;

                itemTotal.RemoveRange(0, 4);

                cb.SetLineWidth(18f);

                for (int i = 1; i <= 13; i++)
                {
                    cb.MoveTo(108f, 0 + stepDistance);
                    cb.LineTo(792f, 0 + stepDistance);
                    cb.Stroke();
                    stepDistance = stepDistance + 225;
                }

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(126f, 0);
                cb.LineTo(774f, 0);
                cb.LineTo(774f, 2700);
                cb.LineTo(126f, 2700);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf3_5x3_5_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(270f, 270f));
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2700));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 3 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemPrint.RemoveRange(0, 3);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 10);
                        diffPerPage.Add("3 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 3);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 30);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(45f, 0f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(315f, 0f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(585f, 0f);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(45f, 270f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(315f, 270f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(585f, 270f);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);


                //Row 3
                PdfReader R3C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(45f, 540f);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(315f, 540f);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(585f, 540f);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);


                //Row 4
                PdfReader R4C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(45f, 810f);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(315f, 810f);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(585f, 810f);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);


                //Row 5
                PdfReader R5C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(45f, 1080f);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(315f, 1080f);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                PdfReader R5C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R5C3Page = writer.GetImportedPage(R5C3File, 1);
                var R5C3PDF = writer.GetImportedPage(R5C3File, 1);
                var R5C3 = new System.Drawing.Drawing2D.Matrix();
                R5C3.Translate(585f, 1080f);
                writer.DirectContent.AddTemplate(R5C3Page, R5C3);


                //Row 6
                PdfReader R6C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R6C1Page = writer.GetImportedPage(R6C1File, 1);
                var R6C1PDF = writer.GetImportedPage(R6C1File, 1);
                var R6C1 = new System.Drawing.Drawing2D.Matrix();
                R6C1.Translate(45f, 1350f);
                writer.DirectContent.AddTemplate(R6C1Page, R6C1);

                PdfReader R6C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R6C2Page = writer.GetImportedPage(R6C2File, 1);
                var R6C2PDF = writer.GetImportedPage(R6C2File, 1);
                var R6C2 = new System.Drawing.Drawing2D.Matrix();
                R6C2.Translate(315f, 1350f);
                writer.DirectContent.AddTemplate(R6C2Page, R6C2);

                PdfReader R6C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R6C3Page = writer.GetImportedPage(R6C3File, 1);
                var R6C3PDF = writer.GetImportedPage(R6C3File, 1);
                var R6C3 = new System.Drawing.Drawing2D.Matrix();
                R6C3.Translate(585f, 1350f);
                writer.DirectContent.AddTemplate(R6C3Page, R6C3);


                //Row 7
                PdfReader R7C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R7C1Page = writer.GetImportedPage(R7C1File, 1);
                var R7C1PDF = writer.GetImportedPage(R7C1File, 1);
                var R7C1 = new System.Drawing.Drawing2D.Matrix();
                R7C1.Translate(45f, 1620f);
                writer.DirectContent.AddTemplate(R7C1Page, R7C1);

                PdfReader R7C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R7C2Page = writer.GetImportedPage(R7C2File, 1);
                var R7C2PDF = writer.GetImportedPage(R7C2File, 1);
                var R7C2 = new System.Drawing.Drawing2D.Matrix();
                R7C2.Translate(315f, 1620f);
                writer.DirectContent.AddTemplate(R7C2Page, R7C2);

                PdfReader R7C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R7C3Page = writer.GetImportedPage(R7C3File, 1);
                var R7C3PDF = writer.GetImportedPage(R7C3File, 1);
                var R7C3 = new System.Drawing.Drawing2D.Matrix();
                R7C3.Translate(585f, 1620f);
                writer.DirectContent.AddTemplate(R7C3Page, R7C3);


                //Row 8
                PdfReader R8C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R8C1Page = writer.GetImportedPage(R8C1File, 1);
                var R8C1PDF = writer.GetImportedPage(R8C1File, 1);
                var R8C1 = new System.Drawing.Drawing2D.Matrix();
                R8C1.Translate(45f, 1890f);
                writer.DirectContent.AddTemplate(R8C1Page, R8C1);

                PdfReader R8C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R8C2Page = writer.GetImportedPage(R8C2File, 1);
                var R8C2PDF = writer.GetImportedPage(R8C2File, 1);
                var R8C2 = new System.Drawing.Drawing2D.Matrix();
                R8C2.Translate(315f, 1890f);
                writer.DirectContent.AddTemplate(R8C2Page, R8C2);

                PdfReader R8C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R8C3Page = writer.GetImportedPage(R8C3File, 1);
                var R8C3PDF = writer.GetImportedPage(R8C3File, 1);
                var R8C3 = new System.Drawing.Drawing2D.Matrix();
                R8C3.Translate(585f, 1890f);
                writer.DirectContent.AddTemplate(R8C3Page, R8C3);


                //Row 9
                PdfReader R9C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R9C1Page = writer.GetImportedPage(R9C1File, 1);
                var R9C1PDF = writer.GetImportedPage(R9C1File, 1);
                var R9C1 = new System.Drawing.Drawing2D.Matrix();
                R9C1.Translate(45f, 2160f);
                writer.DirectContent.AddTemplate(R9C1Page, R9C1);

                PdfReader R9C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R9C2Page = writer.GetImportedPage(R9C2File, 1);
                var R9C2PDF = writer.GetImportedPage(R9C2File, 1);
                var R9C2 = new System.Drawing.Drawing2D.Matrix();
                R9C2.Translate(315f, 2160f);
                writer.DirectContent.AddTemplate(R9C2Page, R9C2);

                PdfReader R9C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R9C3Page = writer.GetImportedPage(R9C3File, 1);
                var R9C3PDF = writer.GetImportedPage(R9C3File, 1);
                var R9C3 = new System.Drawing.Drawing2D.Matrix();
                R9C3.Translate(585f, 2160f);
                writer.DirectContent.AddTemplate(R9C3Page, R9C3);


                //Row 10
                PdfReader R10C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R10C1Page = writer.GetImportedPage(R10C1File, 1);
                var R10C1PDF = writer.GetImportedPage(R10C1File, 1);
                var R10C1 = new System.Drawing.Drawing2D.Matrix();
                R10C1.Translate(45f, 2430f);
                writer.DirectContent.AddTemplate(R10C1Page, R10C1);

                PdfReader R10C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R10C2Page = writer.GetImportedPage(R10C2File, 1);
                var R10C2PDF = writer.GetImportedPage(R10C2File, 1);
                var R10C2 = new System.Drawing.Drawing2D.Matrix();
                R10C2.Translate(315f, 2430f);
                writer.DirectContent.AddTemplate(R10C2Page, R10C2);

                PdfReader R10C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R10C3Page = writer.GetImportedPage(R10C3File, 1);
                var R10C3PDF = writer.GetImportedPage(R10C3File, 1);
                var R10C3 = new System.Drawing.Drawing2D.Matrix();
                R10C3.Translate(585f, 2430f);
                writer.DirectContent.AddTemplate(R10C3Page, R10C3);

                itemTotal.RemoveRange(0, 3);


                cb.SetLineWidth(18f);

                //Cropmarks Horizontal
                cb.MoveTo(0f, 0);
                cb.LineTo(900f, 0);
                cb.Stroke();

                cb.MoveTo(0f, 270);
                cb.LineTo(900f, 270);
                cb.Stroke();

                cb.MoveTo(0f, 540);
                cb.LineTo(900f, 540);
                cb.Stroke();

                cb.MoveTo(0f, 810);
                cb.LineTo(900f, 810);
                cb.Stroke();

                cb.MoveTo(0f, 1080);
                cb.LineTo(900f, 1080);
                cb.Stroke();

                cb.MoveTo(0f, 1350);
                cb.LineTo(900f, 1350);
                cb.Stroke();

                cb.MoveTo(0f, 1620);
                cb.LineTo(900f, 1620);
                cb.Stroke();

                cb.MoveTo(0f, 1890);
                cb.LineTo(900f, 1890);
                cb.Stroke();

                cb.MoveTo(0f, 2160);
                cb.LineTo(900f, 2160);
                cb.Stroke();

                cb.MoveTo(0f, 2430);
                cb.LineTo(900f, 2430);
                cb.Stroke();

                cb.MoveTo(0f, 2700);
                cb.LineTo(900f, 2700);
                cb.Stroke();


                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27f, 0);
                cb.LineTo(873f, 0);
                cb.LineTo(873f, 2700);
                cb.LineTo(27f, 2700);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf3_5_Triangle_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(283.5f, 283.5f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 283.5f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2551.5f));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 3 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemPrint.RemoveRange(0, 3);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 10);
                        diffPerPage.Add("3 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 3);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 30);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(24.75f, 0f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(308.25f, 0f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(591.75f, 0f);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(24.75f, 283.5f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(308.25f, 283.5f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(591.75f, 283.5f);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);


                //Row 3
                PdfReader R3C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(24.75f, 567f);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(308.25f, 567f);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(591.75f, 567f);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);


                //Row 4
                PdfReader R4C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(24.75f, 850.5f);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(308.25f, 850.5f);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(591.75f, 850.5f);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);


                //Row 5
                PdfReader R5C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(24.75f, 1134f);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(308.25f, 1134f);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                PdfReader R5C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R5C3Page = writer.GetImportedPage(R5C3File, 1);
                var R5C3PDF = writer.GetImportedPage(R5C3File, 1);
                var R5C3 = new System.Drawing.Drawing2D.Matrix();
                R5C3.Translate(591.75f, 1134f);
                writer.DirectContent.AddTemplate(R5C3Page, R5C3);


                //Row 6
                PdfReader R6C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R6C1Page = writer.GetImportedPage(R6C1File, 1);
                var R6C1PDF = writer.GetImportedPage(R6C1File, 1);
                var R6C1 = new System.Drawing.Drawing2D.Matrix();
                R6C1.Translate(24.75f, 1417.5f);
                writer.DirectContent.AddTemplate(R6C1Page, R6C1);

                PdfReader R6C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R6C2Page = writer.GetImportedPage(R6C2File, 1);
                var R6C2PDF = writer.GetImportedPage(R6C2File, 1);
                var R6C2 = new System.Drawing.Drawing2D.Matrix();
                R6C2.Translate(308.25f, 1417.5f);
                writer.DirectContent.AddTemplate(R6C2Page, R6C2);

                PdfReader R6C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R6C3Page = writer.GetImportedPage(R6C3File, 1);
                var R6C3PDF = writer.GetImportedPage(R6C3File, 1);
                var R6C3 = new System.Drawing.Drawing2D.Matrix();
                R6C3.Translate(591.75f, 1417.5f);
                writer.DirectContent.AddTemplate(R6C3Page, R6C3);


                //Row 7
                PdfReader R7C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R7C1Page = writer.GetImportedPage(R7C1File, 1);
                var R7C1PDF = writer.GetImportedPage(R7C1File, 1);
                var R7C1 = new System.Drawing.Drawing2D.Matrix();
                R7C1.Translate(24.75f, 1701f);
                writer.DirectContent.AddTemplate(R7C1Page, R7C1);

                PdfReader R7C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R7C2Page = writer.GetImportedPage(R7C2File, 1);
                var R7C2PDF = writer.GetImportedPage(R7C2File, 1);
                var R7C2 = new System.Drawing.Drawing2D.Matrix();
                R7C2.Translate(308.25f, 1701f);
                writer.DirectContent.AddTemplate(R7C2Page, R7C2);

                PdfReader R7C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R7C3Page = writer.GetImportedPage(R7C3File, 1);
                var R7C3PDF = writer.GetImportedPage(R7C3File, 1);
                var R7C3 = new System.Drawing.Drawing2D.Matrix();
                R7C3.Translate(591.75f, 1701f);
                writer.DirectContent.AddTemplate(R7C3Page, R7C3);


                //Row 8
                PdfReader R8C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R8C1Page = writer.GetImportedPage(R8C1File, 1);
                var R8C1PDF = writer.GetImportedPage(R8C1File, 1);
                var R8C1 = new System.Drawing.Drawing2D.Matrix();
                R8C1.Translate(24.75f, 1984.5f);
                writer.DirectContent.AddTemplate(R8C1Page, R8C1);

                PdfReader R8C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R8C2Page = writer.GetImportedPage(R8C2File, 1);
                var R8C2PDF = writer.GetImportedPage(R8C2File, 1);
                var R8C2 = new System.Drawing.Drawing2D.Matrix();
                R8C2.Translate(308.25f, 1984.5f);
                writer.DirectContent.AddTemplate(R8C2Page, R8C2);

                PdfReader R8C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R8C3Page = writer.GetImportedPage(R8C3File, 1);
                var R8C3PDF = writer.GetImportedPage(R8C3File, 1);
                var R8C3 = new System.Drawing.Drawing2D.Matrix();
                R8C3.Translate(591.75f, 1984.5f);
                writer.DirectContent.AddTemplate(R8C3Page, R8C3);


                //Row 9
                PdfReader R9C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R9C1Page = writer.GetImportedPage(R9C1File, 1);
                var R9C1PDF = writer.GetImportedPage(R9C1File, 1);
                var R9C1 = new System.Drawing.Drawing2D.Matrix();
                R9C1.Translate(24.75f, 2268f);
                writer.DirectContent.AddTemplate(R9C1Page, R9C1);

                PdfReader R9C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R9C2Page = writer.GetImportedPage(R9C2File, 1);
                var R9C2PDF = writer.GetImportedPage(R9C2File, 1);
                var R9C2 = new System.Drawing.Drawing2D.Matrix();
                R9C2.Translate(308.25f, 2268f);
                writer.DirectContent.AddTemplate(R9C2Page, R9C2);

                PdfReader R9C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R9C3Page = writer.GetImportedPage(R9C3File, 1);
                var R9C3PDF = writer.GetImportedPage(R9C3File, 1);
                var R9C3 = new System.Drawing.Drawing2D.Matrix();
                R9C3.Translate(591.75f, 2268f);
                writer.DirectContent.AddTemplate(R9C3Page, R9C3);


                itemTotal.RemoveRange(0, 3);


                cb.SetLineWidth(18f);

                //Cropmarks Horizontal
                cb.MoveTo(0f, 0);
                cb.LineTo(900f, 0);
                cb.Stroke();

                cb.MoveTo(0f, 283.5f);
                cb.LineTo(900f, 283.5f);
                cb.Stroke();

                cb.MoveTo(0f, 567);
                cb.LineTo(900f, 567);
                cb.Stroke();

                cb.MoveTo(0f, 850.5f);
                cb.LineTo(900f, 850.5f);
                cb.Stroke();

                cb.MoveTo(0f, 1134);
                cb.LineTo(900f, 1134);
                cb.Stroke();

                cb.MoveTo(0f, 1417.5f);
                cb.LineTo(900f, 1417.5f);
                cb.Stroke();

                cb.MoveTo(0f, 1701);
                cb.LineTo(900f, 1701);
                cb.Stroke();

                cb.MoveTo(0f, 1984.5f);
                cb.LineTo(900f, 1984.5f);
                cb.Stroke();

                cb.MoveTo(0f, 2268);
                cb.LineTo(900f, 2268);
                cb.Stroke();

                cb.MoveTo(0f, 2551.5f);
                cb.LineTo(900f, 2551.5f);
                cb.Stroke();


                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27f, 0);
                cb.LineTo(873f, 0);
                cb.LineTo(873f, 2551.5f);
                cb.LineTo(27f, 2551.5f);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf3_5x1_25_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(270f, 108f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 90f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2700));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 7 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemTotal.Add(itemPrint[2]);
                        itemTotal.Add(itemPrint[3]);
                        itemTotal.Add(itemPrint[4]);
                        itemTotal.Add(itemPrint[5]);
                        itemTotal.Add(itemPrint[6]);
                        itemPrint.RemoveRange(0, 7);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 10);
                        diffPerPage.Add("7 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 7);

                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 70);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }

            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();

                //Row 1
                PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(180f, 0f);
                R1C1.Rotate(90);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(288f, 0f);
                R1C2.Rotate(90);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                PdfReader R1C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R1C3Page = writer.GetImportedPage(R1C3File, 1);
                var R1C3PDF = writer.GetImportedPage(R1C3File, 1);
                var R1C3 = new System.Drawing.Drawing2D.Matrix();
                R1C3.Translate(396f, 0f);
                R1C3.Rotate(90);
                writer.DirectContent.AddTemplate(R1C3Page, R1C3);

                PdfReader R1C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R1C4Page = writer.GetImportedPage(R1C4File, 1);
                var R1C4PDF = writer.GetImportedPage(R1C4File, 1);
                var R1C4 = new System.Drawing.Drawing2D.Matrix();
                R1C4.Translate(504f, 0f);
                R1C4.Rotate(90);
                writer.DirectContent.AddTemplate(R1C4Page, R1C4);

                PdfReader R1C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R1C5Page = writer.GetImportedPage(R1C5File, 1);
                var R1C5PDF = writer.GetImportedPage(R1C5File, 1);
                var R1C5 = new System.Drawing.Drawing2D.Matrix();
                R1C5.Translate(612f, 0f);
                R1C5.Rotate(90);
                writer.DirectContent.AddTemplate(R1C5Page, R1C5);

                PdfReader R1C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R1C6Page = writer.GetImportedPage(R1C6File, 1);
                var R1C6PDF = writer.GetImportedPage(R1C6File, 1);
                var R1C6 = new System.Drawing.Drawing2D.Matrix();
                R1C6.Translate(720f, 0f);
                R1C6.Rotate(90);
                writer.DirectContent.AddTemplate(R1C6Page, R1C6);

                PdfReader R1C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R1C7Page = writer.GetImportedPage(R1C7File, 1);
                var R1C7PDF = writer.GetImportedPage(R1C7File, 1);
                var R1C7 = new System.Drawing.Drawing2D.Matrix();
                R1C7.Translate(828f, 0f);
                R1C7.Rotate(90);
                writer.DirectContent.AddTemplate(R1C7Page, R1C7);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(180f, 270f);
                R2C1.Rotate(90);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(288f, 270f);
                R2C2.Rotate(90);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                PdfReader R2C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R2C3Page = writer.GetImportedPage(R2C3File, 1);
                var R2C3PDF = writer.GetImportedPage(R2C3File, 1);
                var R2C3 = new System.Drawing.Drawing2D.Matrix();
                R2C3.Translate(396f, 270f);
                R2C3.Rotate(90);
                writer.DirectContent.AddTemplate(R2C3Page, R2C3);

                PdfReader R2C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R2C4Page = writer.GetImportedPage(R2C4File, 1);
                var R2C4PDF = writer.GetImportedPage(R2C4File, 1);
                var R2C4 = new System.Drawing.Drawing2D.Matrix();
                R2C4.Translate(504f, 270f);
                R2C4.Rotate(90);
                writer.DirectContent.AddTemplate(R2C4Page, R2C4);

                PdfReader R2C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R2C5Page = writer.GetImportedPage(R2C5File, 1);
                var R2C5PDF = writer.GetImportedPage(R2C5File, 1);
                var R2C5 = new System.Drawing.Drawing2D.Matrix();
                R2C5.Translate(612f, 270f);
                R2C5.Rotate(90);
                writer.DirectContent.AddTemplate(R2C5Page, R2C5);

                PdfReader R2C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R2C6Page = writer.GetImportedPage(R2C6File, 1);
                var R2C6PDF = writer.GetImportedPage(R2C6File, 1);
                var R2C6 = new System.Drawing.Drawing2D.Matrix();
                R2C6.Translate(720f, 270f);
                R2C6.Rotate(90);
                writer.DirectContent.AddTemplate(R2C6Page, R2C6);

                PdfReader R2C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R2C7Page = writer.GetImportedPage(R2C7File, 1);
                var R2C7PDF = writer.GetImportedPage(R2C7File, 1);
                var R2C7 = new System.Drawing.Drawing2D.Matrix();
                R2C7.Translate(828f, 270f);
                R2C7.Rotate(90);
                writer.DirectContent.AddTemplate(R2C7Page, R2C7);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(180f, 540f);
                R3C1.Rotate(90);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(288f, 540f);
                R3C2.Rotate(90);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                PdfReader R3C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R3C3Page = writer.GetImportedPage(R3C3File, 1);
                var R3C3PDF = writer.GetImportedPage(R3C3File, 1);
                var R3C3 = new System.Drawing.Drawing2D.Matrix();
                R3C3.Translate(396f, 540f);
                R3C3.Rotate(90);
                writer.DirectContent.AddTemplate(R3C3Page, R3C3);

                PdfReader R3C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R3C4Page = writer.GetImportedPage(R3C4File, 1);
                var R3C4PDF = writer.GetImportedPage(R3C4File, 1);
                var R3C4 = new System.Drawing.Drawing2D.Matrix();
                R3C4.Translate(504f, 540f);
                R3C4.Rotate(90);
                writer.DirectContent.AddTemplate(R3C4Page, R3C4);

                PdfReader R3C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R3C5Page = writer.GetImportedPage(R3C5File, 1);
                var R3C5PDF = writer.GetImportedPage(R3C5File, 1);
                var R3C5 = new System.Drawing.Drawing2D.Matrix();
                R3C5.Translate(612f, 540f);
                R3C5.Rotate(90);
                writer.DirectContent.AddTemplate(R3C5Page, R3C5);

                PdfReader R3C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R3C6Page = writer.GetImportedPage(R3C6File, 1);
                var R3C6PDF = writer.GetImportedPage(R3C6File, 1);
                var R3C6 = new System.Drawing.Drawing2D.Matrix();
                R3C6.Translate(720f, 540f);
                R3C6.Rotate(90);
                writer.DirectContent.AddTemplate(R3C6Page, R3C6);

                PdfReader R3C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R3C7Page = writer.GetImportedPage(R3C7File, 1);
                var R3C7PDF = writer.GetImportedPage(R3C7File, 1);
                var R3C7 = new System.Drawing.Drawing2D.Matrix();
                R3C7.Translate(828f, 540f);
                R3C7.Rotate(90);
                writer.DirectContent.AddTemplate(R3C7Page, R3C7);


                //Row 4
                PdfReader R4C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(180f, 810f);
                R4C1.Rotate(90);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(288f, 810f);
                R4C2.Rotate(90);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                PdfReader R4C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R4C3Page = writer.GetImportedPage(R4C3File, 1);
                var R4C3PDF = writer.GetImportedPage(R4C3File, 1);
                var R4C3 = new System.Drawing.Drawing2D.Matrix();
                R4C3.Translate(396f, 810f);
                R4C3.Rotate(90);
                writer.DirectContent.AddTemplate(R4C3Page, R4C3);

                PdfReader R4C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R4C4Page = writer.GetImportedPage(R4C4File, 1);
                var R4C4PDF = writer.GetImportedPage(R4C4File, 1);
                var R4C4 = new System.Drawing.Drawing2D.Matrix();
                R4C4.Translate(504f, 810f);
                R4C4.Rotate(90);
                writer.DirectContent.AddTemplate(R4C4Page, R4C4);

                PdfReader R4C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R4C5Page = writer.GetImportedPage(R4C5File, 1);
                var R4C5PDF = writer.GetImportedPage(R4C5File, 1);
                var R4C5 = new System.Drawing.Drawing2D.Matrix();
                R4C5.Translate(612f, 810f);
                R4C5.Rotate(90);
                writer.DirectContent.AddTemplate(R4C5Page, R4C5);

                PdfReader R4C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R4C6Page = writer.GetImportedPage(R4C6File, 1);
                var R4C6PDF = writer.GetImportedPage(R4C6File, 1);
                var R4C6 = new System.Drawing.Drawing2D.Matrix();
                R4C6.Translate(720f, 810f);
                R4C6.Rotate(90);
                writer.DirectContent.AddTemplate(R4C6Page, R4C6);

                PdfReader R4C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R4C7Page = writer.GetImportedPage(R4C7File, 1);
                var R4C7PDF = writer.GetImportedPage(R4C7File, 1);
                var R4C7 = new System.Drawing.Drawing2D.Matrix();
                R4C7.Translate(828f, 810f);
                R4C7.Rotate(90);
                writer.DirectContent.AddTemplate(R4C7Page, R4C7);


                //Row 5
                PdfReader R5C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(180f, 1080f);
                R5C1.Rotate(90);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(288f, 1080f);
                R5C2.Rotate(90);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                PdfReader R5C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R5C3Page = writer.GetImportedPage(R5C3File, 1);
                var R5C3PDF = writer.GetImportedPage(R5C3File, 1);
                var R5C3 = new System.Drawing.Drawing2D.Matrix();
                R5C3.Translate(396f, 1080f);
                R5C3.Rotate(90);
                writer.DirectContent.AddTemplate(R5C3Page, R5C3);

                PdfReader R5C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R5C4Page = writer.GetImportedPage(R5C4File, 1);
                var R5C4PDF = writer.GetImportedPage(R5C4File, 1);
                var R5C4 = new System.Drawing.Drawing2D.Matrix();
                R5C4.Translate(504f, 1080f);
                R5C4.Rotate(90);
                writer.DirectContent.AddTemplate(R5C4Page, R5C4);

                PdfReader R5C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R5C5Page = writer.GetImportedPage(R5C5File, 1);
                var R5C5PDF = writer.GetImportedPage(R5C5File, 1);
                var R5C5 = new System.Drawing.Drawing2D.Matrix();
                R5C5.Translate(612f, 1080f);
                R5C5.Rotate(90);
                writer.DirectContent.AddTemplate(R5C5Page, R5C5);

                PdfReader R5C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R5C6Page = writer.GetImportedPage(R5C6File, 1);
                var R5C6PDF = writer.GetImportedPage(R5C6File, 1);
                var R5C6 = new System.Drawing.Drawing2D.Matrix();
                R5C6.Translate(720f, 1080f);
                R5C6.Rotate(90);
                writer.DirectContent.AddTemplate(R5C6Page, R5C6);

                PdfReader R5C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R5C7Page = writer.GetImportedPage(R5C7File, 1);
                var R5C7PDF = writer.GetImportedPage(R5C7File, 1);
                var R5C7 = new System.Drawing.Drawing2D.Matrix();
                R5C7.Translate(828f, 1080f);
                R5C7.Rotate(90);
                writer.DirectContent.AddTemplate(R5C7Page, R5C7);


                //Row 6
                PdfReader R6C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R6C1Page = writer.GetImportedPage(R6C1File, 1);
                var R6C1PDF = writer.GetImportedPage(R6C1File, 1);
                var R6C1 = new System.Drawing.Drawing2D.Matrix();
                R6C1.Translate(180f, 1350f);
                R6C1.Rotate(90);
                writer.DirectContent.AddTemplate(R6C1Page, R6C1);

                PdfReader R6C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R6C2Page = writer.GetImportedPage(R6C2File, 1);
                var R6C2PDF = writer.GetImportedPage(R6C2File, 1);
                var R6C2 = new System.Drawing.Drawing2D.Matrix();
                R6C2.Translate(288f, 1350f);
                R6C2.Rotate(90);
                writer.DirectContent.AddTemplate(R6C2Page, R6C2);

                PdfReader R6C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R6C3Page = writer.GetImportedPage(R6C3File, 1);
                var R6C3PDF = writer.GetImportedPage(R6C3File, 1);
                var R6C3 = new System.Drawing.Drawing2D.Matrix();
                R6C3.Translate(396f, 1350f);
                R6C3.Rotate(90);
                writer.DirectContent.AddTemplate(R6C3Page, R6C3);

                PdfReader R6C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R6C4Page = writer.GetImportedPage(R6C4File, 1);
                var R6C4PDF = writer.GetImportedPage(R6C4File, 1);
                var R6C4 = new System.Drawing.Drawing2D.Matrix();
                R6C4.Translate(504f, 1350f);
                R6C4.Rotate(90);
                writer.DirectContent.AddTemplate(R6C4Page, R6C4);

                PdfReader R6C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R6C5Page = writer.GetImportedPage(R6C5File, 1);
                var R6C5PDF = writer.GetImportedPage(R6C5File, 1);
                var R6C5 = new System.Drawing.Drawing2D.Matrix();
                R6C5.Translate(612f, 1350f);
                R6C5.Rotate(90);
                writer.DirectContent.AddTemplate(R6C5Page, R6C5);

                PdfReader R6C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R6C6Page = writer.GetImportedPage(R6C6File, 1);
                var R6C6PDF = writer.GetImportedPage(R6C6File, 1);
                var R6C6 = new System.Drawing.Drawing2D.Matrix();
                R6C6.Translate(720f, 1350f);
                R6C6.Rotate(90);
                writer.DirectContent.AddTemplate(R6C6Page, R6C6);

                PdfReader R6C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R6C7Page = writer.GetImportedPage(R6C7File, 1);
                var R6C7PDF = writer.GetImportedPage(R6C7File, 1);
                var R6C7 = new System.Drawing.Drawing2D.Matrix();
                R6C7.Translate(828f, 1350f);
                R6C7.Rotate(90);
                writer.DirectContent.AddTemplate(R6C7Page, R6C7);


                //Row 7
                PdfReader R7C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R7C1Page = writer.GetImportedPage(R7C1File, 1);
                var R7C1PDF = writer.GetImportedPage(R7C1File, 1);
                var R7C1 = new System.Drawing.Drawing2D.Matrix();
                R7C1.Translate(180f, 1620f);
                R7C1.Rotate(90);
                writer.DirectContent.AddTemplate(R7C1Page, R7C1);

                PdfReader R7C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R7C2Page = writer.GetImportedPage(R7C2File, 1);
                var R7C2PDF = writer.GetImportedPage(R7C2File, 1);
                var R7C2 = new System.Drawing.Drawing2D.Matrix();
                R7C2.Translate(288f, 1620f);
                R7C2.Rotate(90);
                writer.DirectContent.AddTemplate(R7C2Page, R7C2);

                PdfReader R7C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R7C3Page = writer.GetImportedPage(R7C3File, 1);
                var R7C3PDF = writer.GetImportedPage(R7C3File, 1);
                var R7C3 = new System.Drawing.Drawing2D.Matrix();
                R7C3.Translate(396f, 1620f);
                R7C3.Rotate(90);
                writer.DirectContent.AddTemplate(R7C3Page, R7C3);

                PdfReader R7C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R7C4Page = writer.GetImportedPage(R7C4File, 1);
                var R7C4PDF = writer.GetImportedPage(R7C4File, 1);
                var R7C4 = new System.Drawing.Drawing2D.Matrix();
                R7C4.Translate(504f, 1620f);
                R7C4.Rotate(90);
                writer.DirectContent.AddTemplate(R7C4Page, R7C4);

                PdfReader R7C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R7C5Page = writer.GetImportedPage(R7C5File, 1);
                var R7C5PDF = writer.GetImportedPage(R7C5File, 1);
                var R7C5 = new System.Drawing.Drawing2D.Matrix();
                R7C5.Translate(612f, 1620f);
                R7C5.Rotate(90);
                writer.DirectContent.AddTemplate(R7C5Page, R7C5);

                PdfReader R7C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R7C6Page = writer.GetImportedPage(R7C6File, 1);
                var R7C6PDF = writer.GetImportedPage(R7C6File, 1);
                var R7C6 = new System.Drawing.Drawing2D.Matrix();
                R7C6.Translate(720f, 1620f);
                R7C6.Rotate(90);
                writer.DirectContent.AddTemplate(R7C6Page, R7C6);

                PdfReader R7C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R7C7Page = writer.GetImportedPage(R7C7File, 1);
                var R7C7PDF = writer.GetImportedPage(R7C7File, 1);
                var R7C7 = new System.Drawing.Drawing2D.Matrix();
                R7C7.Translate(828f, 1620f);
                R7C7.Rotate(90);
                writer.DirectContent.AddTemplate(R7C7Page, R7C7);


                //Row 8
                PdfReader R8C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R8C1Page = writer.GetImportedPage(R8C1File, 1);
                var R8C1PDF = writer.GetImportedPage(R8C1File, 1);
                var R8C1 = new System.Drawing.Drawing2D.Matrix();
                R8C1.Translate(180f, 1890f);
                R8C1.Rotate(90);
                writer.DirectContent.AddTemplate(R8C1Page, R8C1);

                PdfReader R8C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R8C2Page = writer.GetImportedPage(R8C2File, 1);
                var R8C2PDF = writer.GetImportedPage(R8C2File, 1);
                var R8C2 = new System.Drawing.Drawing2D.Matrix();
                R8C2.Translate(288f, 1890f);
                R8C2.Rotate(90);
                writer.DirectContent.AddTemplate(R8C2Page, R8C2);

                PdfReader R8C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R8C3Page = writer.GetImportedPage(R8C3File, 1);
                var R8C3PDF = writer.GetImportedPage(R8C3File, 1);
                var R8C3 = new System.Drawing.Drawing2D.Matrix();
                R8C3.Translate(396f, 1890f);
                R8C3.Rotate(90);
                writer.DirectContent.AddTemplate(R8C3Page, R8C3);

                PdfReader R8C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R8C4Page = writer.GetImportedPage(R8C4File, 1);
                var R8C4PDF = writer.GetImportedPage(R8C4File, 1);
                var R8C4 = new System.Drawing.Drawing2D.Matrix();
                R8C4.Translate(504f, 1890f);
                R8C4.Rotate(90);
                writer.DirectContent.AddTemplate(R8C4Page, R8C4);

                PdfReader R8C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R8C5Page = writer.GetImportedPage(R8C5File, 1);
                var R8C5PDF = writer.GetImportedPage(R8C5File, 1);
                var R8C5 = new System.Drawing.Drawing2D.Matrix();
                R8C5.Translate(612f, 1890f);
                R8C5.Rotate(90);
                writer.DirectContent.AddTemplate(R8C5Page, R8C5);

                PdfReader R8C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R8C6Page = writer.GetImportedPage(R8C6File, 1);
                var R8C6PDF = writer.GetImportedPage(R8C6File, 1);
                var R8C6 = new System.Drawing.Drawing2D.Matrix();
                R8C6.Translate(720f, 1890f);
                R8C6.Rotate(90);
                writer.DirectContent.AddTemplate(R8C6Page, R8C6);

                PdfReader R8C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R8C7Page = writer.GetImportedPage(R8C7File, 1);
                var R8C7PDF = writer.GetImportedPage(R8C7File, 1);
                var R8C7 = new System.Drawing.Drawing2D.Matrix();
                R8C7.Translate(828f, 1890f);
                R8C7.Rotate(90);
                writer.DirectContent.AddTemplate(R8C7Page, R8C7);


                //Row 9
                PdfReader R9C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R9C1Page = writer.GetImportedPage(R9C1File, 1);
                var R9C1PDF = writer.GetImportedPage(R9C1File, 1);
                var R9C1 = new System.Drawing.Drawing2D.Matrix();
                R9C1.Translate(180f, 2160f);
                R9C1.Rotate(90);
                writer.DirectContent.AddTemplate(R9C1Page, R9C1);

                PdfReader R9C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R9C2Page = writer.GetImportedPage(R9C2File, 1);
                var R9C2PDF = writer.GetImportedPage(R9C2File, 1);
                var R9C2 = new System.Drawing.Drawing2D.Matrix();
                R9C2.Translate(288f, 2160f);
                R9C2.Rotate(90);
                writer.DirectContent.AddTemplate(R9C2Page, R9C2);

                PdfReader R9C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R9C3Page = writer.GetImportedPage(R9C3File, 1);
                var R9C3PDF = writer.GetImportedPage(R9C3File, 1);
                var R9C3 = new System.Drawing.Drawing2D.Matrix();
                R9C3.Translate(396f, 2160f);
                R9C3.Rotate(90);
                writer.DirectContent.AddTemplate(R9C3Page, R9C3);

                PdfReader R9C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R9C4Page = writer.GetImportedPage(R9C4File, 1);
                var R9C4PDF = writer.GetImportedPage(R9C4File, 1);
                var R9C4 = new System.Drawing.Drawing2D.Matrix();
                R9C4.Translate(504f, 2160f);
                R9C4.Rotate(90);
                writer.DirectContent.AddTemplate(R9C4Page, R9C4);

                PdfReader R9C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R9C5Page = writer.GetImportedPage(R9C5File, 1);
                var R9C5PDF = writer.GetImportedPage(R9C5File, 1);
                var R9C5 = new System.Drawing.Drawing2D.Matrix();
                R9C5.Translate(612f, 2160f);
                R9C5.Rotate(90);
                writer.DirectContent.AddTemplate(R9C5Page, R9C5);

                PdfReader R9C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R9C6Page = writer.GetImportedPage(R9C6File, 1);
                var R9C6PDF = writer.GetImportedPage(R9C6File, 1);
                var R9C6 = new System.Drawing.Drawing2D.Matrix();
                R9C6.Translate(720f, 2160f);
                R9C6.Rotate(90);
                writer.DirectContent.AddTemplate(R9C6Page, R9C6);

                PdfReader R9C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R9C7Page = writer.GetImportedPage(R9C7File, 1);
                var R9C7PDF = writer.GetImportedPage(R9C7File, 1);
                var R9C7 = new System.Drawing.Drawing2D.Matrix();
                R9C7.Translate(828f, 2160f);
                R9C7.Rotate(90);
                writer.DirectContent.AddTemplate(R9C7Page, R9C7);

                //Row 10
                PdfReader R10C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R10C1Page = writer.GetImportedPage(R10C1File, 1);
                var R10C1PDF = writer.GetImportedPage(R10C1File, 1);
                var R10C1 = new System.Drawing.Drawing2D.Matrix();
                R10C1.Translate(180f, 2430f);
                R10C1.Rotate(90);
                writer.DirectContent.AddTemplate(R10C1Page, R10C1);

                PdfReader R10C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R10C2Page = writer.GetImportedPage(R10C2File, 1);
                var R10C2PDF = writer.GetImportedPage(R10C2File, 1);
                var R10C2 = new System.Drawing.Drawing2D.Matrix();
                R10C2.Translate(288f, 2430f);
                R10C2.Rotate(90);
                writer.DirectContent.AddTemplate(R10C2Page, R10C2);

                PdfReader R10C3File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[2]) + ".pdf");
                PdfImportedPage R10C3Page = writer.GetImportedPage(R10C3File, 1);
                var R10C3PDF = writer.GetImportedPage(R10C3File, 1);
                var R10C3 = new System.Drawing.Drawing2D.Matrix();
                R10C3.Translate(396f, 2430f);
                R10C3.Rotate(90);
                writer.DirectContent.AddTemplate(R10C3Page, R10C3);

                PdfReader R10C4File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[3]) + ".pdf");
                PdfImportedPage R10C4Page = writer.GetImportedPage(R10C4File, 1);
                var R10C4PDF = writer.GetImportedPage(R10C4File, 1);
                var R10C4 = new System.Drawing.Drawing2D.Matrix();
                R10C4.Translate(504f, 2430f);
                R10C4.Rotate(90);
                writer.DirectContent.AddTemplate(R10C4Page, R10C4);

                PdfReader R10C5File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[4]) + ".pdf");
                PdfImportedPage R10C5Page = writer.GetImportedPage(R10C5File, 1);
                var R10C5PDF = writer.GetImportedPage(R10C5File, 1);
                var R10C5 = new System.Drawing.Drawing2D.Matrix();
                R10C5.Translate(612f, 2430f);
                R10C5.Rotate(90);
                writer.DirectContent.AddTemplate(R10C5Page, R10C5);

                PdfReader R10C6File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[5]) + ".pdf");
                PdfImportedPage R10C6Page = writer.GetImportedPage(R10C6File, 1);
                var R10C6PDF = writer.GetImportedPage(R10C6File, 1);
                var R10C6 = new System.Drawing.Drawing2D.Matrix();
                R10C6.Translate(720f, 2430f);
                R10C6.Rotate(90);
                writer.DirectContent.AddTemplate(R10C6Page, R10C6);

                PdfReader R10C7File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[6]) + ".pdf");
                PdfImportedPage R10C7Page = writer.GetImportedPage(R10C7File, 1);
                var R10C7PDF = writer.GetImportedPage(R10C7File, 1);
                var R10C7 = new System.Drawing.Drawing2D.Matrix();
                R10C7.Translate(828f, 2430f);
                R10C7.Rotate(90);
                writer.DirectContent.AddTemplate(R10C7Page, R10C7);

                itemTotal.RemoveRange(0, 7);

                cb.SetLineWidth(18f);

                //Cropmarks Horizontal
                cb.MoveTo(0f, 0);
                cb.LineTo(900f, 0);
                cb.Stroke();

                cb.MoveTo(0f, 270);
                cb.LineTo(900f, 270);
                cb.Stroke();

                cb.MoveTo(0f, 540);
                cb.LineTo(900f, 540);
                cb.Stroke();

                cb.MoveTo(0f, 810);
                cb.LineTo(900f, 810);
                cb.Stroke();

                cb.MoveTo(0f, 1080);
                cb.LineTo(900f, 1080);
                cb.Stroke();

                cb.MoveTo(0f, 1350);
                cb.LineTo(900f, 1350);
                cb.Stroke();

                cb.MoveTo(0f, 1620);
                cb.LineTo(900f, 1620);
                cb.Stroke();

                cb.MoveTo(0f, 1890);
                cb.LineTo(900f, 1890);
                cb.Stroke();

                cb.MoveTo(0f, 2160);
                cb.LineTo(900f, 2160);
                cb.Stroke();

                cb.MoveTo(0f, 2430);
                cb.LineTo(900f, 2430);
                cb.Stroke();

                cb.MoveTo(0f, 2700);
                cb.LineTo(900f, 2700);
                cb.Stroke();

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(27f, 0);
                cb.LineTo(873f, 0);
                cb.LineTo(873f, 2700);
                cb.LineTo(27f, 2700);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf4_3125x4_3125_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(328.32f, 328.32f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 328.32f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2628));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 2 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemPrint.RemoveRange(0, 2);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 8);
                        diffPerPage.Add("2 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 2);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 16);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                var R1C1 = new System.Drawing.Drawing2D.Matrix();
                R1C1.Translate(121.5f, 0f);
                writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                var R1C2 = new System.Drawing.Drawing2D.Matrix();
                R1C2.Translate(450f, 0f);
                writer.DirectContent.AddTemplate(R1C2Page, R1C2);

                //Row 2
                PdfReader R2C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R2C1Page = writer.GetImportedPage(R2C1File, 1);
                var R2C1PDF = writer.GetImportedPage(R2C1File, 1);
                var R2C1 = new System.Drawing.Drawing2D.Matrix();
                R2C1.Translate(121.5f, 328.5f);
                writer.DirectContent.AddTemplate(R2C1Page, R2C1);

                PdfReader R2C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R2C2Page = writer.GetImportedPage(R2C2File, 1);
                var R2C2PDF = writer.GetImportedPage(R2C2File, 1);
                var R2C2 = new System.Drawing.Drawing2D.Matrix();
                R2C2.Translate(450f, 328.5f);
                writer.DirectContent.AddTemplate(R2C2Page, R2C2);

                //Row 3
                PdfReader R3C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R3C1Page = writer.GetImportedPage(R3C1File, 1);
                var R3C1PDF = writer.GetImportedPage(R3C1File, 1);
                var R3C1 = new System.Drawing.Drawing2D.Matrix();
                R3C1.Translate(121.5f, 657f);
                writer.DirectContent.AddTemplate(R3C1Page, R3C1);

                PdfReader R3C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R3C2Page = writer.GetImportedPage(R3C2File, 1);
                var R3C2PDF = writer.GetImportedPage(R3C2File, 1);
                var R3C2 = new System.Drawing.Drawing2D.Matrix();
                R3C2.Translate(450f, 657f);
                writer.DirectContent.AddTemplate(R3C2Page, R3C2);

                //Row 4
                PdfReader R4C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R4C1Page = writer.GetImportedPage(R4C1File, 1);
                var R4C1PDF = writer.GetImportedPage(R4C1File, 1);
                var R4C1 = new System.Drawing.Drawing2D.Matrix();
                R4C1.Translate(121.5f, 985.5f);
                writer.DirectContent.AddTemplate(R4C1Page, R4C1);

                PdfReader R4C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R4C2Page = writer.GetImportedPage(R4C2File, 1);
                var R4C2PDF = writer.GetImportedPage(R4C2File, 1);
                var R4C2 = new System.Drawing.Drawing2D.Matrix();
                R4C2.Translate(450f, 985.5f);
                writer.DirectContent.AddTemplate(R4C2Page, R4C2);

                //Row 5
                PdfReader R5C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R5C1Page = writer.GetImportedPage(R5C1File, 1);
                var R5C1PDF = writer.GetImportedPage(R5C1File, 1);
                var R5C1 = new System.Drawing.Drawing2D.Matrix();
                R5C1.Translate(121.5f, 1314f);
                writer.DirectContent.AddTemplate(R5C1Page, R5C1);

                PdfReader R5C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R5C2Page = writer.GetImportedPage(R5C2File, 1);
                var R5C2PDF = writer.GetImportedPage(R5C2File, 1);
                var R5C2 = new System.Drawing.Drawing2D.Matrix();
                R5C2.Translate(450f, 1314f);
                writer.DirectContent.AddTemplate(R5C2Page, R5C2);

                //Row 6
                PdfReader R6C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R6C1Page = writer.GetImportedPage(R6C1File, 1);
                var R6C1PDF = writer.GetImportedPage(R6C1File, 1);
                var R6C1 = new System.Drawing.Drawing2D.Matrix();
                R6C1.Translate(121.5f, 1642.5f);
                writer.DirectContent.AddTemplate(R6C1Page, R6C1);

                PdfReader R6C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R6C2Page = writer.GetImportedPage(R6C2File, 1);
                var R6C2PDF = writer.GetImportedPage(R6C2File, 1);
                var R6C2 = new System.Drawing.Drawing2D.Matrix();
                R6C2.Translate(450f, 1642.5f);
                writer.DirectContent.AddTemplate(R6C2Page, R6C2);

                //Row 7
                PdfReader R7C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R7C1Page = writer.GetImportedPage(R7C1File, 1);
                var R7C1PDF = writer.GetImportedPage(R7C1File, 1);
                var R7C1 = new System.Drawing.Drawing2D.Matrix();
                R7C1.Translate(121.5f, 1971f);
                writer.DirectContent.AddTemplate(R7C1Page, R7C1);

                PdfReader R7C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R7C2Page = writer.GetImportedPage(R7C2File, 1);
                var R7C2PDF = writer.GetImportedPage(R7C2File, 1);
                var R7C2 = new System.Drawing.Drawing2D.Matrix();
                R7C2.Translate(450f, 1971f);
                writer.DirectContent.AddTemplate(R7C2Page, R7C2);

                //Row 8
                PdfReader R8C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                PdfImportedPage R8C1Page = writer.GetImportedPage(R8C1File, 1);
                var R8C1PDF = writer.GetImportedPage(R8C1File, 1);
                var R8C1 = new System.Drawing.Drawing2D.Matrix();
                R8C1.Translate(121.5f, 2299.5f);
                writer.DirectContent.AddTemplate(R8C1Page, R8C1);

                PdfReader R8C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                PdfImportedPage R8C2Page = writer.GetImportedPage(R8C2File, 1);
                var R8C2PDF = writer.GetImportedPage(R8C2File, 1);
                var R8C2 = new System.Drawing.Drawing2D.Matrix();
                R8C2.Translate(450f, 2299.5f);
                writer.DirectContent.AddTemplate(R8C2Page, R8C2);


                itemTotal.RemoveRange(0, 2);


                cb.SetLineWidth(18f);

                //Cropmarks Horizontal
                cb.MoveTo(103.5f, 0);
                cb.LineTo(796.5f, 0);
                cb.Stroke();

                cb.MoveTo(103.5f, 328.5f);
                cb.LineTo(796.5f, 328.5f);
                cb.Stroke();

                cb.MoveTo(103.5f, 657);
                cb.LineTo(796.5f, 657);
                cb.Stroke();

                cb.MoveTo(103.5f, 985.5f);
                cb.LineTo(796.5f, 985.5f);
                cb.Stroke();

                cb.MoveTo(103.5f, 1314);
                cb.LineTo(796.5f, 1314);
                cb.Stroke();

                cb.MoveTo(103.5f, 1642.5f);
                cb.LineTo(796.5f, 1642.5f);
                cb.Stroke();

                cb.MoveTo(103.5f, 1971);
                cb.LineTo(796.5f, 1971);
                cb.Stroke();

                cb.MoveTo(103.5f, 2299.5f);
                cb.LineTo(796.5f, 2299.5f);
                cb.Stroke();

                cb.MoveTo(103.5f, 2628);
                cb.LineTo(796.5f, 2628);
                cb.Stroke();


                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(121.5f, 0);
                cb.LineTo(778.5f, 0);
                cb.LineTo(778.5f, 2628);
                cb.LineTo(121.5f, 2628);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }

        public List<string> pdf4_25x1_6800(FormMain mainForm, string fileName, string[] art, int[] qty)
        {
            foreach (string file in art)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfReader inputFile = new PdfReader(file);
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(324f, 90f));
                doc1.SetMargins(0, 0, 0, 0);
                doc1.Open();
                doc1.NewPage();
                var imp = writer1.GetImportedPage(inputFile, 1);
                var tm = new System.Drawing.Drawing2D.Matrix();
                PdfImportedPage page = writer1.GetImportedPage(inputFile, 1);
                if (page.Height != 90f)
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

            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(900, 2700));
            doc.SetMargins(0, 0, 0, 0);
            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;

            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            item = art.ToList();
            itemQty = qty.ToList();

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
                    if (itemPrint.Count() % 2 == 0)
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[1]);
                        itemPrint.RemoveRange(0, 2);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 30);
                        diffPerPage.Add("2 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveRange(0, 2);
                    }
                    else
                    {
                        itemTotal.Add(itemPrint[0]);
                        itemTotal.Add(itemPrint[0]);
                        itemPrint.RemoveAt(0);
                        printed = (int)Math.Ceiling((double)itemQtyPrint[0] / 60);
                        diffPerPage.Add("1 Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                        itemQtyPrint.RemoveAt(0);
                    }
                    mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                }
            }

            string path = Settings.Default.tempDir;

            while (itemTotal.Count() > 0)
            {
                doc.NewPage();
                //Row 1
                float stepDistance = 0;
                for (int i = 1; i <= 30; i++)
                {

                    PdfReader R1C1File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[0]) + ".pdf");
                    PdfImportedPage R1C1Page = writer.GetImportedPage(R1C1File, 1);
                    var R1C1PDF = writer.GetImportedPage(R1C1File, 1);
                    var R1C1 = new System.Drawing.Drawing2D.Matrix();
                    R1C1.Translate(126f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C1Page, R1C1);

                    PdfReader R1C2File = new PdfReader(path + "\\" + Path.GetFileNameWithoutExtension(itemTotal[1]) + ".pdf");
                    PdfImportedPage R1C2Page = writer.GetImportedPage(R1C2File, 1);
                    var R1C2PDF = writer.GetImportedPage(R1C2File, 1);
                    var R1C2 = new System.Drawing.Drawing2D.Matrix();
                    R1C2.Translate(450f, stepDistance);
                    writer.DirectContent.AddTemplate(R1C2Page, R1C2);
                    stepDistance = stepDistance + 90;
                }
                stepDistance = 0;

                itemTotal.RemoveRange(0, 2);


                cb.SetLineWidth(18f);

                for (int i = 1; i <= 31; i++)
                {
                    //Cropmarks Horizontal
                    cb.MoveTo(103.5f, stepDistance);
                    cb.LineTo(796.5f, stepDistance);
                    cb.Stroke();
                    stepDistance = stepDistance + 90;
                }

                cb.SetColorFill(new CMYKColor(0f, 0f, 0f, 0f));
                cb.MoveTo(121.5f, 0);
                cb.LineTo(778.5f, 0);
                cb.LineTo(778.5f, 2700);
                cb.LineTo(121.5f, 2700);
                cb.Fill();
            }
            doc.Close();

            return diffPerPage;
        }
    }
}
