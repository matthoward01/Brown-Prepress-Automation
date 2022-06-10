using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using Brown_Prepress_Automation.Properties;

namespace Brown_Prepress_Automation
{
    class NourisonOutput
    {
        MethodsCommon methods = new MethodsCommon();
        PdfProcessing pdfProcessing = new PdfProcessing();

        public void pdfNourisonPOP(FormMain mainForm, List<string> passedList, string workingFile)
        {
            int pageNumber = 1;
            int fileProgressStep = (int)Math.Ceiling(((double)100) / (passedList.Count / 10));
            while (passedList.Count > 0)
            {
                FileStream fs1 = new FileStream(Settings.Default.tempDir + "\\" + "Nourison " + Path.GetFileNameWithoutExtension(workingFile) + " - Pg" + pageNumber + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                Document doc1 = new Document();
                PdfWriter writer1 = PdfWriter.GetInstance(doc1, fs1);
                writer1.PdfVersion = PdfWriter.VERSION_1_3;
                float paperWidth = 3600;
                float paperHeight = 7128;
                doc1.SetPageSize(new iTextSharp.text.Rectangle(paperWidth, paperHeight));
                doc1.Open();
                doc1.NewPage();
                float yPosition = 414;
                for (int i = 1; i <= 5; i++)
                {
                    pdfProcessing.PdfPlacement(writer1, passedList[0], -9, yPosition, 0, 1);
                    pdfProcessing.PdfPlacement(writer1, passedList[1], 1800, yPosition, 0, 1);
                    /*PdfReader c1r1File = new PdfReader(passedList[0]);
                    PdfImportedPage c1r1Page = writer1.GetImportedPage(c1r1File, 1);
                    var c1r1Pdf = new System.Drawing.Drawing2D.Matrix();
                    c1r1Pdf.Translate(-9, yPosition);
                    writer1.DirectContent.AddTemplate(c1r1Page, c1r1Pdf);

                    PdfReader c2r1File = new PdfReader(passedList[1]);
                    PdfImportedPage c2r1Page = writer1.GetImportedPage(c2r1File, 1);
                    var c2r1Pdf = new System.Drawing.Drawing2D.Matrix();
                    c2r1Pdf.Translate(1800, yPosition);
                    writer1.DirectContent.AddTemplate(c2r1Page, c2r1Pdf);*/

                    yPosition = yPosition + 1260;
                    passedList.RemoveRange(0, 2);
                }
                PdfContentByte cb = writer1.DirectContent;
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.Circle(72f, 72f, 9f);
                cb.Circle(paperWidth - 72f, 72f, 9f);
                cb.Circle(72f, paperHeight - 72f, 9f);
                cb.Circle(paperWidth - 72f, paperHeight - 72f, 9f);
                cb.Fill();
                doc1.Close();

                System.IO.File.Copy(Settings.Default.tempDir + "\\" + "Nourison " + Path.GetFileNameWithoutExtension(workingFile) + " - Pg" + pageNumber + ".pdf", Settings.Default.nourisonHpOutput + "\\" + "Nourison " + Path.GetFileNameWithoutExtension(workingFile) + " - Pg" + pageNumber + ".pdf", true);
                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                pageNumber++;
            }

        }
    }
}
