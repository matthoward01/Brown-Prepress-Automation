using Brown_Prepress_Automation.Properties;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Brown_Prepress_Automation
{
    class AveryLabel
    {
        PdfProcessing pdfProcessing = new PdfProcessing();
        string gillRFont = "Fonts\\GIL_____.TTF";
        string gillBFont = "Fonts\\GILB____.TTF";

        public void CreateAveryLabel (string name, List<string> fileName, List<string> qty, List<string> size, List<string> wo, List<string> so)
        {
            string averyFileName = Settings.Default.tempDir + "\\" + Path.GetFileName(name) + " Avery.pdf";

            FileStream fsAvery = new FileStream(Settings.Default.tempDir + "\\" + Path.GetFileName(name) + " Avery.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document docAvery = new Document();
            PdfWriter writerAvery = PdfWriter.GetInstance(docAvery, fsAvery);
            writerAvery.PdfVersion = PdfWriter.VERSION_1_3;
            docAvery.SetPageSize(new iTextSharp.text.Rectangle(216, 72));
            docAvery.SetMargins(0, 0, 0, 0);
            docAvery.Open();
            PdfContentByte cbAvery = writerAvery.DirectContent;
            BaseFont GillSansR = BaseFont.CreateFont(gillRFont, BaseFont.CP1252, BaseFont.EMBEDDED);
            BaseFont GillSansB = BaseFont.CreateFont(gillBFont, BaseFont.CP1252, BaseFont.EMBEDDED);

            int pdfPageCount = 0;

            for (int i = 0; i < fileName.Count(); i++)
            {
                float countFloat = (Int32.Parse(qty[i]) / 250);
                int count = (int)Math.Ceiling(countFloat);
                
                for (int z = 0; z <= count+1; z++)
                {
                    docAvery.NewPage();
                    cbAvery.BeginText();
                    cbAvery.SetFontAndSize(GillSansR, 10);
                    cbAvery.ShowTextAligned(Element.ALIGN_LEFT, fileName[i], 9, 50.4f, 0);
                    cbAvery.ShowTextAligned(Element.ALIGN_LEFT, "Size: " + size[i], 9, 38.16f, 0);
                    cbAvery.ShowTextAligned(Element.ALIGN_LEFT, "WO: " + wo[i], 9, 26.28f, 0);
                    cbAvery.ShowTextAligned(Element.ALIGN_LEFT, "QTY: " + qty[i], 9, 14.04f, 0);
                    cbAvery.ShowTextAligned(Element.ALIGN_RIGHT, so[i], 207, 14.04f, 0);
                    //cb.SetTextMatrix(62, 728);
                    //cb.ShowText(fileName[i]);
                    cbAvery.EndText();
                    pdfPageCount++;
                }
            }

            if (pdfPageCount % 24 != 0)
            {
                while (pdfPageCount % 24 != 0)
                {
                    docAvery.NewPage();
                    cbAvery.BeginText();
                    cbAvery.EndText();
                    pdfPageCount++;
                }
            }

            docAvery.Close();
            if (!Directory.Exists(Settings.Default.xmfHotfolders + "\\XMF 3 x 1 - AVERY\\"))
            {
                Directory.CreateDirectory(Settings.Default.xmfHotfolders + "\\XMF 3 x 1 - AVERY\\");
            }
            File.Copy(averyFileName, Settings.Default.xmfHotfolders + "\\XMF 3 x 1 - AVERY\\" + Path.GetFileName(averyFileName), true);

            //return averyFileName;
        }
    }
}
