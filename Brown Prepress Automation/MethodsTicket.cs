using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelLibrary.SpreadSheet;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Xml.Linq;
using Brown_Prepress_Automation.Properties;

namespace Brown_Prepress_Automation
{
    class MethodsTicket
    {        
        MethodsCommon commonMethods = new MethodsCommon();
        MethodsMail mail = new MethodsMail();
        AveryLabel averyLabel = new AveryLabel();

        string gillRFont = "Fonts\\GIL_____.TTF";
        string gillBFont = "Fonts\\GILB____.TTF";

        public void shawPrintableTicket(string file, string name, List<string> numberUpList, List<int> lines, string formattedSize, List<string> diffPerPage, bool needAvery, bool dist)
        {
            List<string> filenameList = new List<string>();
            List<string> qtyList = new List<string>();
            List<string> sizeList = new List<string>();
            List<string> woList = new List<string>();
            List<string> soList = new List<string>();

            int i = 0;
            Workbook book = Workbook.Load(file);
            Worksheet sheet = book.Worksheets[0];
            int validCellsCheck = commonMethods.countValidCells(file, 1, 0, 0);
            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(612, 792));
            doc.SetMargins(0, 0, 1, 1);
            doc.Open();
            PdfContentByte cb = writer.DirectContent;
            BaseFont GillSansR = BaseFont.CreateFont(gillRFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            BaseFont GillSansB = BaseFont.CreateFont(gillBFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            //doc.NewPage();
            cb.BeginText();
            cb.SetFontAndSize(GillSansB, 13);
            cb.SetTextMatrix(62, 728);
            cb.ShowText(name);
            cb.EndText();
            //cb.SetLineWidth(2f);
            //cb.SetGrayStroke(0f);
            //cb.MoveTo(60, 798);
            //cb.LineTo(576, 798);
            //cb.Stroke();
            Paragraph paragraphTable = new Paragraph();
            Paragraph sheetsText = new Paragraph();
            paragraphTable.SpacingAfter = 30f;
            PdfPTable table = new PdfPTable(7);
            table.TotalWidth = 576f;
            float[] widths = new float[] { 2f, 1f, 1f, 1f, 1f, 1f, 3f };
            table.SetWidths(widths);
            PdfPCell labelid = new PdfPCell(new Phrase(sheet.Cells[0, 0].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(labelid);
            //PdfPCell style = new PdfPCell(new Phrase(sheet.Cells[0, 1].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            //table.AddCell(style);
            //PdfPCell type = new PdfPCell(new Phrase(sheet.Cells[0, 2].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            //table.AddCell(type);
            PdfPCell color = new PdfPCell(new Phrase(sheet.Cells[0, 3].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(color);
            PdfPCell size = new PdfPCell(new Phrase(sheet.Cells[0, 4].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(size);
            PdfPCell qty = new PdfPCell(new Phrase(sheet.Cells[0, 5].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(qty);
            PdfPCell attn = new PdfPCell(new Phrase(sheet.Cells[0, 6].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(attn);
            PdfPCell costcenter = new PdfPCell(new Phrase(sheet.Cells[0, 7].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(costcenter);
            PdfPCell finishing = new PdfPCell(new Phrase(sheet.Cells[0, 8].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(finishing);
            foreach (int line in lines)
            {
                string fileName = sheet.Cells[line, 0].StringValue;
                string fileNameOriginal = sheet.Cells[line, 0].StringValue;
                int index = fileName.IndexOf(" - ");
                if (index > 0)
                {
                    fileName = fileName.Substring(0, index);
                }
                PdfPCell labelidText = new PdfPCell(new Phrase(sheet.Cells[line, 0].StringValue));
                int fileNameLength = fileNameOriginal.Length - fileName.Length;
                if ((fileNameLength < 6) && (sheet.Cells[line, 2].StringValue.Trim() != ""))
                {
                    labelidText = new PdfPCell(new Phrase(fileName + " - " + sheet.Cells[line, 2].StringValue));
                }
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(labelidText);
                //PdfPCell styleText = new PdfPCell(new Phrase(sheet.Cells[line, 1].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                //table.AddCell(styleText);
                //PdfPCell typeText = new PdfPCell(new Phrase(sheet.Cells[line, 2].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                //table.AddCell(typeText);
                PdfPCell colorText = new PdfPCell(new Phrase(sheet.Cells[line, 3].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(colorText);
                PdfPCell sizeText = new PdfPCell(new Phrase(formattedSize));
                //PdfPCell sizeText = new PdfPCell(new Phrase(sheet.Cells[line, 4].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(sizeText);
                PdfPCell qtyText = new PdfPCell(new Phrase(sheet.Cells[line, 5].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(qtyText);
                PdfPCell attnText = new PdfPCell(new Phrase(sheet.Cells[line, 6].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(attnText);
                PdfPCell costcenterText = new PdfPCell(new Phrase(sheet.Cells[line, 7].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(costcenterText);
                PdfPCell finishingText = new PdfPCell();
                if (numberUpList.Any())
                {
                    finishingText = new PdfPCell(new Phrase(numberUpList[i] + sheet.Cells[line, 8].StringValue));
                }
                else
                {
                    finishingText = new PdfPCell(new Phrase(sheet.Cells[line, 8].StringValue));
                }
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(finishingText);

                filenameList.Add(fileName + " - " + sheet.Cells[line, 2].StringValue);
                sizeList.Add(formattedSize);
                qtyList.Add(sheet.Cells[line, 5].StringValue);
                woList.Add(sheet.Cells[line, 6].StringValue);
                soList.Add(sheet.Cells[line, 7].StringValue);
                i++;
            }
            table.SpacingBefore = 21f;
            table.SpacingAfter = 30f;
            paragraphTable.Add("    " + DateTime.Now.ToString());
            if (dist)
            {
                int countSheets = 1;

                while (diffPerPage.Count > 0 && diffPerPage[0] != "")
                {
                    sheetsText.Add("    Spread " + countSheets + " = " + diffPerPage[0] + "\r\n");

                    countSheets++;
                    diffPerPage.RemoveAt(0);
                }
            }

            doc.Add(paragraphTable);
            doc.Add(table);
            doc.Add(sheetsText);

            doc.Close();
            System.IO.Directory.CreateDirectory(Settings.Default.tempDir);
            if (Settings.Default.sendEmails == true)
            {
                mail.SendMailTicket(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", "shaw");
            }
            if (needAvery)
            {
                averyLabel.CreateAveryLabel(name, filenameList, qtyList, sizeList, woList, soList);
            }
            if (!Settings.Default.debugOn)
            {
                commonMethods.SendToPrinter(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", true, false);
            }
        }

        public void shawPrintableTicket6800(string file, string name, List<string> diffPerPage, List<int> lines, string formattedSize)
        {
            Workbook book = Workbook.Load(file);
            Worksheet sheet = book.Worksheets[0];
            int validCellsCheck = commonMethods.countValidCells(file, 1, 0, 0);
            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(612, 792));
            doc.SetMargins(0, 0, 1, 1);
            doc.Open();
            PdfContentByte cb = writer.DirectContent;
            BaseFont GillSansR = BaseFont.CreateFont(gillRFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            BaseFont GillSansB = BaseFont.CreateFont(gillBFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            //doc.NewPage();
            cb.BeginText();
            cb.SetFontAndSize(GillSansB, 13);
            cb.SetTextMatrix(62, 728);
            cb.ShowText(name);
            cb.EndText();
            //cb.SetLineWidth(2f);
            //cb.SetGrayStroke(0f);
            //cb.MoveTo(60, 798);
            //cb.LineTo(576, 798);
            //cb.Stroke();
            Paragraph paragraphTable = new Paragraph();
            Paragraph sheetsText = new Paragraph();
            paragraphTable.SpacingAfter = 30f;
            PdfPTable table = new PdfPTable(7);
            table.TotalWidth = 576f;
            float[] widths = new float[] { 2f, 1f, 1f, 1f, 1f, 1f, 2f };
            table.SetWidths(widths);
            PdfPCell labelid = new PdfPCell(new Phrase(sheet.Cells[0, 0].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(labelid);
            PdfPCell color = new PdfPCell(new Phrase(sheet.Cells[0, 3].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(color);
            PdfPCell size = new PdfPCell(new Phrase(sheet.Cells[0, 4].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(size);
            PdfPCell qty = new PdfPCell(new Phrase(sheet.Cells[0, 5].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(qty);
            PdfPCell attn = new PdfPCell(new Phrase(sheet.Cells[0, 6].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(attn);
            PdfPCell costcenter = new PdfPCell(new Phrase(sheet.Cells[0, 7].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(costcenter);
            PdfPCell finishing = new PdfPCell(new Phrase(sheet.Cells[0, 8].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(finishing);
            foreach (int line in lines)
            {
                string fileName = sheet.Cells[line, 0].StringValue;
                string fileNameOriginal = sheet.Cells[line, 0].StringValue;
                int index = fileName.IndexOf(" - ");
                if (index > 0)
                {
                    fileName = fileName.Substring(0, index);
                }
                PdfPCell labelidText = new PdfPCell(new Phrase(sheet.Cells[line, 0].StringValue));
                int fileNameLength = fileNameOriginal.Length - fileName.Length;
                if ((fileNameLength < 6) && (sheet.Cells[line, 2].StringValue.Trim() != ""))
                {
                    labelidText = new PdfPCell(new Phrase(fileName + " - " + sheet.Cells[line, 2].StringValue));
                }                
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(labelidText);
                PdfPCell colorText = new PdfPCell(new Phrase(sheet.Cells[line, 3].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(colorText);
                PdfPCell sizeText = new PdfPCell(new Phrase(formattedSize));
                //PdfPCell sizeText = new PdfPCell(new Phrase(sheet.Cells[line, 4].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(sizeText);
                PdfPCell qtyText = new PdfPCell(new Phrase(sheet.Cells[line, 5].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(qtyText);
                PdfPCell attnText = new PdfPCell(new Phrase(sheet.Cells[line, 6].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(attnText);
                PdfPCell costcenterText = new PdfPCell(new Phrase(sheet.Cells[line, 7].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(costcenterText);
                PdfPCell finishingText = new PdfPCell(new Phrase(sheet.Cells[line, 8].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(finishingText);
            }
            table.SpacingBefore = 21f;
            table.SpacingAfter = 30f;
            paragraphTable.Add("    " + DateTime.Now.ToString());
            int countSheets = 1;

            while (diffPerPage.Count > 0 && diffPerPage[0] != "")
            {
                sheetsText.Add("    Spread " + countSheets + " = " + diffPerPage[0] + "\r\n");

                countSheets++;
                diffPerPage.RemoveAt(0);
            }

            doc.Add(paragraphTable);
            doc.Add(table);
            doc.Add(sheetsText);
            diffPerPage.Clear();

            doc.Close();
            System.IO.Directory.CreateDirectory(Settings.Default.tempDir);
            if (Settings.Default.sendEmails == true)
            {
                mail.SendMailTicket(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", "shaw");
            }
            if (!Settings.Default.debugOn)
            {
                commonMethods.SendToPrinter(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", true, false);
            }
        }

        public void shawBoardPrintableTicket(string file, string name)
        {
            Workbook book = Workbook.Load(file);
            Worksheet sheet = book.Worksheets[0];
            int validCellsCheck = commonMethods.countValidCells(file, 1, 0, 0);            
            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(612, 792));
            doc.SetMargins(0, 0, 1, 1);
            doc.Open();
            PdfContentByte cb = writer.DirectContent;
            BaseFont GillSansR = BaseFont.CreateFont(gillRFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            BaseFont GillSansB = BaseFont.CreateFont(gillBFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            //doc.NewPage();
            cb.BeginText();
            cb.SetFontAndSize(GillSansB, 18);
            cb.SetTextMatrix(62, 728);
            cb.ShowText(name);
            cb.EndText();
            //cb.SetLineWidth(2f);
            //cb.SetGrayStroke(0f);
            //cb.MoveTo(60, 798);
            //cb.LineTo(576, 798);
            //cb.Stroke();
            Paragraph paragraphTable = new Paragraph();
            Paragraph sheetsText = new Paragraph();
            paragraphTable.SpacingAfter = 30f;
            PdfPTable table = new PdfPTable(9);
            table.TotalWidth = 576f;
            float[] widths = new float[] { 2f, 1f, 1f, 1f, 1f, 1f, 1f, 2f, 1f };
            table.SetWidths(widths);
            PdfPCell partnumber = new PdfPCell(new Phrase(sheet.Cells[0, 0].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(partnumber);
            PdfPCell coverid = new PdfPCell(new Phrase(sheet.Cells[0, 1].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(coverid);
            PdfPCell linerid = new PdfPCell(new Phrase(sheet.Cells[0, 2].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(linerid);
            PdfPCell flid = new PdfPCell(new Phrase(sheet.Cells[0, 3].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(flid);
            PdfPCell blid = new PdfPCell(new Phrase(sheet.Cells[0, 4].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(blid);
            PdfPCell size = new PdfPCell(new Phrase(sheet.Cells[0, 5].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(size);
            PdfPCell colors = new PdfPCell(new Phrase(sheet.Cells[0, 6].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(colors);
            PdfPCell material = new PdfPCell(new Phrase(sheet.Cells[0, 7].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(material);
            PdfPCell qty = new PdfPCell(new Phrase(sheet.Cells[0, 8].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(qty);
            for (int i = 1; i < validCellsCheck; i++)
            {
                PdfPCell partnumberText = new PdfPCell(new Phrase(sheet.Cells[i, 0].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(partnumberText);
                PdfPCell coverIDText = new PdfPCell(new Phrase(sheet.Cells[i, 1].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(coverIDText);
                PdfPCell linerIDText = new PdfPCell(new Phrase(sheet.Cells[i, 2].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(linerIDText);
                PdfPCell flIDText = new PdfPCell(new Phrase(sheet.Cells[i, 3].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(flIDText);
                PdfPCell blIDText = new PdfPCell(new Phrase(sheet.Cells[i, 4].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(blIDText);
                PdfPCell sizeText = new PdfPCell(new Phrase(sheet.Cells[i, 5].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(sizeText);
                PdfPCell colorsText = new PdfPCell(new Phrase(sheet.Cells[i, 6].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(colorsText);
                PdfPCell materialText = new PdfPCell(new Phrase(sheet.Cells[i, 7].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(materialText);
                PdfPCell qtyText = new PdfPCell(new Phrase(sheet.Cells[i, 8].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(qtyText);
            }
            table.SpacingBefore = 21f;
            table.SpacingAfter = 30f;
            paragraphTable.Add("    " + DateTime.Now.ToString());

            doc.Add(paragraphTable);
            doc.Add(table);
            doc.Add(sheetsText);

            doc.Close();
            System.IO.Directory.CreateDirectory(Settings.Default.tempDir);
            mail.SendMailTicket(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", "shaw");
            if (!Settings.Default.debugOn)
            {
                commonMethods.SendToPrinter(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", true, false);
            }
        }

        public void ddpTicket(string file, string name)
        {
            Workbook book = Workbook.Load(file);
            Worksheet sheet = book.Worksheets[0];
            int validCellsCheck = commonMethods.countValidCells(file, 1, 0, 0);
            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(612, 792));
            doc.SetMargins(0, 0, 1, 1);
            doc.Open();
            PdfContentByte cb = writer.DirectContent;
            BaseFont GillSansR = BaseFont.CreateFont(gillRFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            BaseFont GillSansB = BaseFont.CreateFont(gillBFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            //doc.NewPage();
            cb.BeginText();
            cb.SetFontAndSize(GillSansB, 18);
            cb.SetTextMatrix(62, 728);
            cb.ShowText(name);
            cb.EndText();
            //cb.SetLineWidth(2f);
            //cb.SetGrayStroke(0f);
            //cb.MoveTo(60, 798);
            //cb.LineTo(576, 798);
            //cb.Stroke();
            Paragraph paragraphTable = new Paragraph();
            Paragraph sheetsText = new Paragraph();
            paragraphTable.SpacingAfter = 30f;
            PdfPTable table = new PdfPTable(7);
            table.TotalWidth = 576f;
            float[] widths = new float[] { 1f, 1f, 2f, 2f, 1f, 0.75f, 0.5f };
            table.SetWidths(widths);
            PdfPCell labelid = new PdfPCell(new Phrase(sheet.Cells[5, 1].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(labelid);
            PdfPCell style = new PdfPCell(new Phrase(sheet.Cells[5, 2].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(style);
            PdfPCell type = new PdfPCell(new Phrase(sheet.Cells[5, 3].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(type);
            PdfPCell color = new PdfPCell(new Phrase(sheet.Cells[5, 4].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(color);
            PdfPCell size = new PdfPCell(new Phrase(sheet.Cells[5, 5].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(size);
            //PdfPCell qty = new PdfPCell(new Phrase(sheet.Cells[5, 6].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            //table.AddCell(qty);
            PdfPCell attn = new PdfPCell(new Phrase(sheet.Cells[5, 7].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(attn);
            //PdfPCell costcenter = new PdfPCell(new Phrase(sheet.Cells[5, 8].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            //table.AddCell(costcenter);
            PdfPCell finishing = new PdfPCell(new Phrase(sheet.Cells[5, 9].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(finishing);
            for (int i = 6; i < commonMethods.countValidCells(file, 6, 0, 0); i++)
            {
                PdfPCell labelidText = new PdfPCell(new Phrase(sheet.Cells[i, 1].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(labelidText);
                PdfPCell styleText = new PdfPCell(new Phrase(sheet.Cells[i, 2].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(styleText);
                PdfPCell typeText = new PdfPCell(new Phrase(sheet.Cells[i, 3].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(typeText);
                PdfPCell colorText = new PdfPCell(new Phrase(sheet.Cells[i, 4].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(colorText);
                PdfPCell sizeText = new PdfPCell(new Phrase(sheet.Cells[i, 5].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(sizeText);
                //PdfPCell qtyText = new PdfPCell(new Phrase(sheet.Cells[i, 6].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                //table.AddCell(qtyText);
                PdfPCell attnText = new PdfPCell(new Phrase(sheet.Cells[i, 7].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(attnText);
                //PdfPCell costcenterText = new PdfPCell(new Phrase(sheet.Cells[i, 8].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                //table.AddCell(costcenterText);
                PdfPCell finishingText = new PdfPCell(new Phrase(sheet.Cells[i, 9].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(finishingText);
            }
            table.SpacingBefore = 21f;
            table.SpacingAfter = 30f;
            paragraphTable.Add("    " + DateTime.Now.ToString());

            doc.Add(paragraphTable);
            doc.Add(table);

            doc.Close();
            System.IO.Directory.CreateDirectory(Settings.Default.tempDir);
            if (!Settings.Default.debugOn)
            {
                commonMethods.SendToPrinter(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", true, false);
            }
        }

        public void armstrongTicket6800(string file, string name, List<string> diffPerPage, List<int> lines)
        {
            Workbook book = Workbook.Load(file);
            Worksheet sheet = book.Worksheets[0];
            int validCellsCheck = commonMethods.countValidCells(file, 1, 0, 0);
            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(612, 792));
            doc.SetMargins(0, 0, 1, 1);
            doc.Open();
            PdfContentByte cb = writer.DirectContent;
            BaseFont GillSansR = BaseFont.CreateFont(gillRFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            BaseFont GillSansB = BaseFont.CreateFont(gillBFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            //doc.NewPage();
            cb.BeginText();
            cb.SetFontAndSize(GillSansB, 13);
            cb.SetTextMatrix(62, 728);
            cb.ShowText(name);
            cb.EndText();
            //cb.SetLineWidth(2f);
            //cb.SetGrayStroke(0f);
            //cb.MoveTo(60, 798);
            //cb.LineTo(576, 798);
            //cb.Stroke();
            Paragraph paragraphTable = new Paragraph();
            Paragraph sheetsText = new Paragraph();
            paragraphTable.SpacingAfter = 30f;
            PdfPTable table = new PdfPTable(6);
            table.TotalWidth = 576f;
            float[] widths = new float[] { 2f, 2f, 1f, 1f, 3f, 1f};
            table.SetWidths(widths);
            PdfPCell partNumber = new PdfPCell(new Phrase(sheet.Cells[0, 0].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(partNumber);
            PdfPCell armstrongPart = new PdfPCell(new Phrase(sheet.Cells[0, 1].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(armstrongPart);
            PdfPCell size = new PdfPCell(new Phrase(sheet.Cells[0, 2].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(size);
            PdfPCell qty = new PdfPCell(new Phrase(sheet.Cells[0, 3].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(qty);
            PdfPCell finish = new PdfPCell(new Phrase(sheet.Cells[0, 4].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(finish);
            PdfPCell soLine = new PdfPCell(new Phrase(sheet.Cells[0, 5].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(soLine);

            foreach (int line in lines)
            {
                PdfPCell partNumberText = new PdfPCell(new Phrase(sheet.Cells[line, 0].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(partNumberText);
                PdfPCell armstrongPartText = new PdfPCell(new Phrase(sheet.Cells[line, 1].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(armstrongPartText);
                PdfPCell sizeText = new PdfPCell(new Phrase(sheet.Cells[line, 2].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(sizeText);
                PdfPCell qtyText = new PdfPCell(new Phrase(sheet.Cells[line, 3].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(qtyText);
                PdfPCell finishText = new PdfPCell(new Phrase(sheet.Cells[line, 4].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(finishText);
                PdfPCell soLineText = new PdfPCell(new Phrase(sheet.Cells[line, 5].StringValue));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(soLineText);  
            }
            table.SpacingBefore = 21f;
            table.SpacingAfter = 30f;
            paragraphTable.Add("    " + DateTime.Now.ToString());
            int countSheets = 1;

            while (diffPerPage.Count > 0 && diffPerPage[0] != "")
            {
                sheetsText.Add("    Spread " + countSheets + " = " + diffPerPage[0] + "\r\n");

                countSheets++;
                diffPerPage.RemoveAt(0);
            }

            doc.Add(paragraphTable);
            doc.Add(table);
            doc.Add(sheetsText);
            diffPerPage.Clear();

            doc.Close();
            System.IO.Directory.CreateDirectory(Settings.Default.tempDir);
            if (Settings.Default.sendEmails == true)
            {
                mail.SendMailTicket(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", "armstrong");
            }
            if (!Settings.Default.debugOn)
            {
                commonMethods.SendToPrinter(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", true, false);
            }
        }
        
        public void armstrongTicket(string file, string name, List<int> lines, ModelArmstrong.ArmstrongSheet armstrongSheet)
        {
            Workbook book = Workbook.Load(file);
            Worksheet sheet = book.Worksheets[0];
            int validCellsCheck = commonMethods.countValidCells(file, 1, 0, 0);
            FileStream fs = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            writer.PdfVersion = PdfWriter.VERSION_1_3;
            doc.SetPageSize(new iTextSharp.text.Rectangle(612, 792));
            doc.SetMargins(0, 0, 1, 1);
            doc.Open();
            PdfContentByte cb = writer.DirectContent;
            BaseFont GillSansR = BaseFont.CreateFont(gillRFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            BaseFont GillSansB = BaseFont.CreateFont(gillBFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            //doc.NewPage();
            cb.BeginText();
            cb.SetFontAndSize(GillSansB, 13);
            cb.SetTextMatrix(62, 728);
            cb.ShowText(name + " 16x16");
            cb.EndText();
            //cb.SetLineWidth(2f);
            //cb.SetGrayStroke(0f);
            //cb.MoveTo(60, 798);
            //cb.LineTo(576, 798);
            //cb.Stroke();
            Paragraph paragraphTable = new Paragraph();
            Paragraph sheetsText = new Paragraph();
            paragraphTable.SpacingAfter = 30f;
            PdfPTable table = new PdfPTable(6);
            table.TotalWidth = 576f;
            float[] widths = new float[] { 2f, 2f, 1f, 1f, 3f, 1f };
            table.SetWidths(widths);
            PdfPCell partNumber = new PdfPCell(new Phrase(sheet.Cells[0, 0].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(partNumber);
            PdfPCell armstrongPart = new PdfPCell(new Phrase(sheet.Cells[0, 1].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(armstrongPart);
            PdfPCell size = new PdfPCell(new Phrase(sheet.Cells[0, 2].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(size);
            PdfPCell qty = new PdfPCell(new Phrase(sheet.Cells[0, 3].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(qty);
            PdfPCell finish = new PdfPCell(new Phrase(sheet.Cells[0, 4].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(finish);
            PdfPCell soLine = new PdfPCell(new Phrase(sheet.Cells[0, 5].StringValue));
            //labelid.Border = 0;
            //labelid.HorizontalAlignment = 0;
            table.AddCell(soLine);

            foreach (int line in lines)
            {
                PdfPCell partNumberText = new PdfPCell(new Phrase(armstrongSheet.PartNumber[line]));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(partNumberText);
                PdfPCell armstrongPartText = new PdfPCell(new Phrase(armstrongSheet.FileName[line]));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(armstrongPartText);
                PdfPCell sizeText = new PdfPCell(new Phrase(armstrongSheet.Size[line]));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(sizeText);
                PdfPCell qtyText = new PdfPCell(new Phrase(armstrongSheet.Quantity[line]));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(qtyText);
                PdfPCell finishText = new PdfPCell(new Phrase(armstrongSheet.Stock[line]));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(finishText);
                PdfPCell soLineText = new PdfPCell(new Phrase(armstrongSheet.SalesOrder[line]));
                //labelid.Border = 0;
                //labelid.HorizontalAlignment = 0;
                table.AddCell(soLineText);
            }
            table.SpacingBefore = 21f;
            table.SpacingAfter = 30f;
            paragraphTable.Add("    " + DateTime.Now.ToString());

            doc.Add(paragraphTable);
            doc.Add(table);
            doc.Add(sheetsText);

            doc.Close();
            System.IO.Directory.CreateDirectory(Settings.Default.tempDir);
            if (Settings.Default.sendEmails == true)
            {
                mail.SendMailTicket(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", "armstrong");
            }
            if (!Settings.Default.debugOn)
            {
                commonMethods.SendToPrinter(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(file) + ".pdf", true, false);
            }
        }
    }
}
