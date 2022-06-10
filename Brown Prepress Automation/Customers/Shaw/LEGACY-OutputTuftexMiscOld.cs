using System;
using iTextSharp.text;
using ExcelLibrary.SpreadSheet;
using System.IO;
using iTextSharp.text.pdf;
using System.util;
using System.Reflection;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Brown_Prepress_Automation.Properties;
using Brown_Prepress_Automation;

public class AddPageMethodsMiscOld
{
    string agarRFont = "Fonts\\AGaramondPro-Regular.otf";
    string gillRFont = "Fonts\\GIL_____.TTF";
    string gillBFont = "Fonts\\GILB____.TTF";
    string gothamBoldFont = "Fonts\\Gotham-Bold.otf";
    string gothamBookFont = "Fonts\\Gotham-Book.otf";
    string gothamLightFont = "Fonts\\Gotham-Light.otf";
    string gothamMediumFont = "Fonts\\Gotham-Medium.otf";

    MethodsCommon methods = new MethodsCommon();

    public void AddPageNormal(Document doc, PdfContentByte cb, PdfWriter writer,
                                                               string imagePath1AddMethod, string imagePath2AddMethod,
                                                               string woNormal1, string woNormal2,
                                                               string styleNormal1, string styleNormal2,
                                                               string colorNormal1, string colorNormal2,
                                                               string dateNormal1, string dateNormal2,
                                                               string seqNormal1, string seqNormal2,
                                                               string labelNormal1, string labelNormal2,
                                                               string type)
    {

        var outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);

        //var logoimage = Path.Combine(outPutDirectory, "Fonts\\GIL_____.TTF");
        //string relFont = new Uri(logoimage).LocalPath;

        //Read and Set Posistion of the image file
        BaseFont GillSansR = BaseFont.CreateFont(gillRFont, BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont GillSansB = BaseFont.CreateFont(gillBFont, BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont AGarR = BaseFont.CreateFont(agarRFont, BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont GothamBold = BaseFont.CreateFont(gothamBoldFont, BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont GothamBook = BaseFont.CreateFont(gothamBookFont, BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont GothamLight = BaseFont.CreateFont(gothamLightFont, BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont GothamMedium = BaseFont.CreateFont(gothamMediumFont, BaseFont.CP1252, BaseFont.EMBEDDED);

            iTextSharp.text.Image image1AddMethod = iTextSharp.text.Image.GetInstance(new Uri(imagePath1AddMethod));
            image1AddMethod.ScaleToFit(864, 612);
            image1AddMethod.SetAbsolutePosition(0, 648);
        

            iTextSharp.text.Image image2AddMethod = iTextSharp.text.Image.GetInstance(new Uri(imagePath2AddMethod));
            image2AddMethod.ScaleToFit(864, 612);
            image2AddMethod.SetAbsolutePosition(0, 36);
        

        string imagePathBlockout = "Images\\Blank.jpg";
        string imagePathBlockout2 = "Images\\Blank.jpg";


        Workbook labelCheck = Workbook.Load("Type\\labels.xls");
        Worksheet labelSheetCheck = labelCheck.Worksheets[0];
        string createType1 = "";
        string createType2 = "";
        float createTypePositionX1 = 0;
        float createTypePositionY1 = 0;
        float createTypePositionX2 = 0;
        float createTypePositionY2 = 0;

        int labelTypeCount = 1;
        while (!labelSheetCheck.Cells[labelTypeCount, 1].IsEmpty)
        {
            labelTypeCount++;
        }
        for (int i = 1; i < labelTypeCount; i++)
        {
            if (Path.GetFileName(imagePath1AddMethod) == labelSheetCheck.Cells[i, 1].StringValue)
            {
                //createNumberofLines1 = labelSheetCheck.Cells[i, 1].StringValue;
                createType1 = labelSheetCheck.Cells[i, 0].StringValue;
                if (!labelSheetCheck.Cells[i, 2].IsEmpty)
                {
                    createTypePositionX1 = float.Parse(labelSheetCheck.Cells[i, 2].StringValue);
                    createTypePositionY1 = float.Parse(labelSheetCheck.Cells[i, 3].StringValue);
                }
                else
                {
                    createTypePositionX1 = 0;
                    createTypePositionY1 = 0;
                }
            }
            if (Path.GetFileName(imagePath2AddMethod) == labelSheetCheck.Cells[i, 1].StringValue)
            {
                //createNumberofLines2 = labelSheetCheck.Cells[i, 1].StringValue;
                createType2 = labelSheetCheck.Cells[i, 0].StringValue;
                if (!labelSheetCheck.Cells[i, 2].IsEmpty)
                {
                    createTypePositionX2 = float.Parse(labelSheetCheck.Cells[i, 2].StringValue);
                    createTypePositionY2 = float.Parse(labelSheetCheck.Cells[i, 3].StringValue);
                }
                else
                {
                    createTypePositionX2 = 0;
                    createTypePositionY2 = 0;
                }
            }
        }

        iTextSharp.text.Image image1BlockOut = iTextSharp.text.Image.GetInstance(new Uri(imagePathBlockout));
        image1BlockOut.ScaleToFit(180, 18);
        image1BlockOut.SetAbsolutePosition(90f + createTypePositionX1, 1089f + createTypePositionY1);

        iTextSharp.text.Image image2BlockOut = iTextSharp.text.Image.GetInstance(new Uri(imagePathBlockout));
        image2BlockOut.ScaleToFit(180, 18);
        image2BlockOut.SetAbsolutePosition(90f + createTypePositionX2, 477f + createTypePositionY2);

        iTextSharp.text.Image image1NewBlockOut = iTextSharp.text.Image.GetInstance(new Uri(imagePathBlockout));
        image1NewBlockOut.ScaleToFit(216, 18);
        image1NewBlockOut.SetAbsolutePosition(126f + createTypePositionX1, 1048.5f + createTypePositionY1);

        iTextSharp.text.Image image2NewBlockOut = iTextSharp.text.Image.GetInstance(new Uri(imagePathBlockout));
        image2NewBlockOut.ScaleToFit(216, 18);
        image2NewBlockOut.SetAbsolutePosition(126f + createTypePositionX2, 436.5f + createTypePositionY2);

        iTextSharp.text.Image image1NewGenericBlockOut = iTextSharp.text.Image.GetInstance(new Uri(imagePathBlockout));
        image1NewGenericBlockOut.ScaleToFit(216, 18);
        image1NewGenericBlockOut.SetAbsolutePosition(132f + createTypePositionX1, 1033f + createTypePositionY1);

        iTextSharp.text.Image image2NewGenericBlockOut = iTextSharp.text.Image.GetInstance(new Uri(imagePathBlockout));
        image2NewGenericBlockOut.ScaleToFit(216, 18);
        image2NewGenericBlockOut.SetAbsolutePosition(132f + createTypePositionX2, 421f + createTypePositionY2);

        iTextSharp.text.Image image1StyleBlockOut = iTextSharp.text.Image.GetInstance(new Uri(imagePathBlockout));
        image1StyleBlockOut.ScaleToFit(180, 18);
        image1StyleBlockOut.SetAbsolutePosition(90f, 1107f);

        iTextSharp.text.Image image2StyleBlockOut = iTextSharp.text.Image.GetInstance(new Uri(imagePathBlockout));
        image2StyleBlockOut.ScaleToFit(180, 18);
        image2StyleBlockOut.SetAbsolutePosition(90f, 495f);

        iTextSharp.text.Image image1BlockOutMisc = iTextSharp.text.Image.GetInstance(new Uri(imagePathBlockout2));
        image1BlockOutMisc.ScaleToFit(180, 54);
        image1BlockOutMisc.SetAbsolutePosition(29.7864f, 738f);

        iTextSharp.text.Image image2BlockOutMisc = iTextSharp.text.Image.GetInstance(new Uri(imagePathBlockout2));
        image2BlockOutMisc.ScaleToFit(180, 54);
        image2BlockOutMisc.SetAbsolutePosition(29.7864f, 126f);

        // Add a new page
        doc.NewPage();
        //Add Images to PDF Page
        if (File.Exists(Settings.Default.tuftexPdf + "\\" + Path.GetFileNameWithoutExtension(imagePath1AddMethod) + ".pdf") == false)
        {
            doc.Add(image1AddMethod);
        }
        else
        {
            PdfReader baseImage1 = new PdfReader(Settings.Default.tuftexPdf + "\\" + Path.GetFileNameWithoutExtension((imagePath1AddMethod)) + ".pdf");
            PdfImportedPage BaseImage1Page = writer.GetImportedPage(baseImage1, 1);
            var BaseImage1Var = new System.Drawing.Drawing2D.Matrix();
            BaseImage1Var.Translate(0f, 648f);
            writer.DirectContentUnder.AddTemplate(BaseImage1Page, BaseImage1Var);
        }
        if (createType1 != "NFA" ||
            labelNormal1 != "26172" ||
            labelNormal1 != "26173" ||
            labelNormal1 != "26178")
        {
            doc.Add(image1BlockOutMisc);
        }
        if (type == "PVD01")
        {
            cb.BeginText();
            cb.SetFontAndSize(GillSansB, 12);
            doc.Add(image1StyleBlockOut);
            cb.SetCMYKColorFill(0, 0, 0, 255);
            cb.SetTextMatrix(93.24f, 1111.4064f);
            cb.ShowText((styleNormal1).Replace(" Ii ", " II ").Replace(" Iii ", " III "));
            cb.EndText();
        }

        cb.BeginText();
        cb.SetFontAndSize(GillSansR, 12);
        if (createType1 == "NFA")
        {
            doc.Add(image1BlockOut);
            cb.SetCMYKColorFill(0, 0, 0, 255);
            cb.SetFontAndSize(AGarR, 14);
            cb.SetTextMatrix(126.864f, 962.64f);
        }
        else if (labelNormal1 == "26178")
        {
            cb.SetCMYKColorFill(0, 0, 0, 255);
            cb.SetFontAndSize(GothamBook, 12);
            doc.Add(image1NewBlockOut);
            cb.SetTextMatrix(135.36f + createTypePositionX1, 1053.36f + createTypePositionY1);
        }
        else if (labelNormal1 == "26172" ||
                 labelNormal1 == "26173")
        {
            cb.SetCMYKColorFill(0, 0, 0, 255);
            cb.SetFontAndSize(GothamBook, 12);
            doc.Add(image1NewGenericBlockOut);
            cb.SetTextMatrix(135.36f + createTypePositionX1, 1038.24f + createTypePositionY1);
        }
        else
        {
            doc.Add(image1BlockOut);
            cb.SetTextMatrix(93.24f + createTypePositionX1, 1094.4f + createTypePositionY1);
        }
        //cb.ShowText(createTypePositionY1);
        cb.ShowText(colorNormal1);
        cb.EndText();
        if (createType1 != "NFA" ||
            labelNormal1 != "26172" ||
            labelNormal1 != "26173" ||
            labelNormal1 != "26178")
        {
            if ((type == "normal") || (type == "ddp") || (type == "SEQ"))
            {
                cb.BeginText();
                cb.SetFontAndSize(GillSansR, 9);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(36f, 781.056f);
                cb.ShowText(woNormal1);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(GillSansR, 9);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(36f, 764.064f);
                cb.ShowText(dateNormal1);
                cb.EndText();
                if ((type != "SEQ"))
                {
                    cb.BeginText();
                    cb.SetFontAndSize(GillSansR, 9);
                    cb.SetCMYKColorFill(0, 0, 0, 255);
                    cb.SetTextMatrix(36f, 742.32f);
                    if (type == "ddp")
                    {
                        seqNormal1 = "Display Sequence: 0";
                    }
                    cb.ShowText(seqNormal1);
                    cb.EndText();
                }
            }
        }
        else if (labelNormal1 == "26172" ||
                 labelNormal1 == "26173" ||
                 labelNormal1 == "26178")
        {
            if ((type == "normal") || (type == "ddp") || (type == "SEQ"))
            {
                cb.BeginText();
                cb.SetFontAndSize(GothamBook, 8);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(92.52f, 727.49f);
                cb.ShowText(woNormal1);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(GothamBook, 8);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(120.6f, 727.49f);
                cb.ShowText(dateNormal1);
                cb.EndText();
                if ((type != "SEQ"))
                {
                    cb.BeginText();
                    cb.SetFontAndSize(GothamBook, 8);
                    cb.SetCMYKColorFill(0, 0, 0, 255);
                    cb.SetTextMatrix(92.52f, 737.514f);
                    if (type == "ddp")
                    {
                        seqNormal1 = "Display Sequence: 0";
                    }
                    cb.ShowText(seqNormal1);
                    cb.EndText();
                }
            }
        }
        else
        {
            if ((type == "normal") || (type == "ddp") || (type == "SEQ"))
            {
                cb.BeginText();
                cb.SetFontAndSize(AGarR, 9);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(48.24f, 880.704f);
                cb.ShowText(woNormal1);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(AGarR, 9);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(97.92f, 880.704f);
                cb.ShowText(dateNormal1);
                cb.EndText();

                if (type != "SEQ")
                {
                    cb.BeginText();
                    cb.SetFontAndSize(AGarR, 9);
                    cb.SetCMYKColorFill(0, 0, 0, 255);
                    cb.SetTextMatrix(142.272f, 880.704f);
                    if (type == "ddp")
                    {
                        seqNormal1 = "Display Sequence: 0";
                    }
                    cb.ShowText(seqNormal1);
                    cb.EndText();
                }
            }
        }
        if (File.Exists(Settings.Default.tuftexPdf + "\\" + Path.GetFileNameWithoutExtension(imagePath2AddMethod) + ".pdf") == false)
        {
            doc.Add(image2AddMethod);
        }
        else
        {
            PdfReader baseImage2 = new PdfReader(Settings.Default.tuftexPdf + "\\" + Path.GetFileNameWithoutExtension((imagePath2AddMethod)) + ".pdf");
            PdfImportedPage BaseImage2Page = writer.GetImportedPage(baseImage2, 1);
            var BaseImage2Var = new System.Drawing.Drawing2D.Matrix();
            BaseImage2Var.Translate(0f, 36f);
            writer.DirectContentUnder.AddTemplate(BaseImage2Page, BaseImage2Var);
        }
        if (createType2 != "NFA" ||
            labelNormal2 != "26172" ||
            labelNormal2 != "26173" ||
            labelNormal2 != "26178")
        {
            doc.Add(image2BlockOutMisc);
        }

        if (type == "PVD01")
        {
            cb.BeginText();
            cb.SetFontAndSize(GillSansB, 12);
            doc.Add(image2StyleBlockOut);
            cb.SetCMYKColorFill(0, 0, 0, 255);
            cb.SetTextMatrix(93.24f, 499.4064f);
            cb.ShowText(styleNormal2);
            cb.EndText();
        }

        cb.BeginText();
        cb.SetFontAndSize(GillSansR, 12);
        if (createType2 == "NFA")
        {
            doc.Add(image2BlockOut);
            cb.SetCMYKColorFill(0, 0, 0, 255);
            cb.SetFontAndSize(AGarR, 14);
            cb.SetTextMatrix(126.864f, 350.64f);
        }
        else if (labelNormal2 == "26178")
        {
            cb.SetCMYKColorFill(0, 0, 0, 255);
            cb.SetFontAndSize(GothamBook, 12);
            doc.Add(image2NewBlockOut);
            cb.SetTextMatrix(135.36f + createTypePositionX1, 441.36f + createTypePositionY1);
        }
        else if (labelNormal2 == "26172" ||
             labelNormal2 == "26173")
        {
            cb.SetCMYKColorFill(0, 0, 0, 255);
            cb.SetFontAndSize(GothamBook, 12);
            doc.Add(image2NewGenericBlockOut);
            cb.SetTextMatrix(135.36f + createTypePositionX1, 426.24f + createTypePositionY1);
        }
        else
        {
            doc.Add(image2BlockOut);
            cb.SetTextMatrix(93.24f + createTypePositionX2, 482.4f + createTypePositionY2);
        }
        cb.ShowText(colorNormal2);
        cb.EndText();
        if (createType2 != "NFA" ||
            labelNormal2 != "26172" ||
            labelNormal2 != "26173" ||
            labelNormal2 != "26178")
        {
            if ((type == "normal") || (type == "ddp") || (type == "SEQ"))
            {
                cb.BeginText();
                cb.SetFontAndSize(GillSansR, 9);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(36f, 169.056f);
                cb.ShowText(woNormal2);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(GillSansR, 9);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(36f, 152.064f);
                cb.ShowText(dateNormal2);
                cb.EndText();

                if (type != "SEQ")
                {
                    cb.BeginText();
                    cb.SetFontAndSize(GillSansR, 9);
                    cb.SetCMYKColorFill(0, 0, 0, 255);
                    cb.SetTextMatrix(36f, 130.32f);
                    if (type == "ddp")
                    {
                        seqNormal2 = "Display Sequence 0";
                    }
                    cb.ShowText(seqNormal2);
                    cb.EndText();
                }
            }
        }
        else if (labelNormal2 == "26172" ||
                 labelNormal2 == "26173" ||
                 labelNormal2 == "26178")
        {
            if ((type == "normal") || (type == "ddp") || (type == "SEQ"))
            {
                cb.BeginText();
                cb.SetFontAndSize(GothamBook, 8);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(92.52f, 115.49f);
                cb.ShowText(woNormal1);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(GothamBook, 8);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(120.6f, 115.49f);
                cb.ShowText(dateNormal1);
                cb.EndText();
                if ((type != "SEQ"))
                {
                    cb.BeginText();
                    cb.SetFontAndSize(GothamBook, 8);
                    cb.SetCMYKColorFill(0, 0, 0, 255);
                    cb.SetTextMatrix(92.52f, 125.514f);
                    if (type == "ddp")
                    {
                        seqNormal1 = "Display Sequence 0";
                    }
                    cb.ShowText(seqNormal1);
                    cb.EndText();
                }
            }
        }
        else
        {
            if ((type == "normal") || (type == "ddp"))
            {
                cb.BeginText();
                cb.SetFontAndSize(AGarR, 9);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(48.24f, 268.704f);
                cb.ShowText(woNormal2);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(AGarR, 9);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(97.92f, 268.704f);
                cb.ShowText(dateNormal2);
                cb.EndText();

                if (type != "SEQ")
                {
                    cb.BeginText();
                    cb.SetFontAndSize(AGarR, 9);
                    cb.SetCMYKColorFill(0, 0, 0, 255);
                    cb.SetTextMatrix(142.272f, 268.704f);
                    if (type == "ddp")
                    {
                        seqNormal2 = "Display Sequence: 0";
                    }
                    cb.ShowText(seqNormal2);
                    cb.EndText();
                }
            }
        }
        //Clear image variables
        image1AddMethod = null;
        image2AddMethod = null;
    }
}

