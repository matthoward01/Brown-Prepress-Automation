using System;
using iTextSharp.text;
using System.IO;
using iTextSharp.text.pdf;
using System.util;
using System.Reflection;
using System.Xml.Linq;
using System.Linq;
using System.Xml;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using Brown_Prepress_Automation.Properties;
using Brown_Prepress_Automation;

public class AddPageMethodsTuftex
{
    string gothamBoldFont = "Fonts\\Gotham-Bold.otf";
    string gothamBookFont = "Fonts\\Gotham-Book.otf";
    string gothamLightFont = "Fonts\\Gotham-Light.otf";
    string gothamMediumFont = "Fonts\\Gotham-Medium.otf";

    ///////////////////////////////////////////
    ////////////////OLD FONTS//////////////////
    ///////////////////////////////////////////
    string agarRFont = "Fonts\\AGaramondPro-Regular.otf";
    string gillRFont = "Fonts\\GIL_____.TTF";
    string gillBFont = "Fonts\\GILB____.TTF";
    string uniCBFont = "Fonts\\UniversLTStd-BoldCn.ttf";
    ///////////////////////////////////////////
    ///////////////////////////////////////////
    ///////////////////////////////////////////

    MethodsCommon methods = new MethodsCommon();

    public void AddPageNormal(Document doc,
                            PdfContentByte cb,
                            PdfWriter writer,
                            string woNumber,
                            string styleNumber,
                            string durability,
                            string styleWidth,
                            string stylePatternRepeat,
                            string styleName,
                            string styleFiber,
                            string styleBase,
                            string[] bugs,
                            string colorSequence,
                            string colorNumber,
                            string colorName)
    {

        BaseFont GothamBold = BaseFont.CreateFont(gothamBoldFont, BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont GothamBook = BaseFont.CreateFont(gothamBookFont, BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont GothamLight = BaseFont.CreateFont(gothamLightFont, BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont GothamMedium = BaseFont.CreateFont(gothamMediumFont, BaseFont.CP1252, BaseFont.EMBEDDED);

        ///////////////////////////////////////////
        ////////////////OLD FONTS//////////////////
        ///////////////////////////////////////////
        BaseFont GillSansR = BaseFont.CreateFont(gillRFont, BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont GillSansB = BaseFont.CreateFont(gillBFont, BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont AGarR = BaseFont.CreateFont(agarRFont, BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont UniCB = BaseFont.CreateFont(uniCBFont, BaseFont.CP1252, BaseFont.EMBEDDED);
        ///////////////////////////////////////////
        ///////////////////////////////////////////
        ///////////////////////////////////////////

        //Bug Images
        //xml Variable Sheet
        string[] styleSplit = Regex.Split(styleName, "\n");
        var xmlVariables = XDocument.Load("Variables.xml");

        //////////////////////
        //Correct Fiber Name//
        //////////////////////
        var fiberVariables = from c in xmlVariables.Root.Descendants("fiber")
                             where (string)c.Attribute("id") == styleFiber
                             select c.Element("text").Value;
        foreach (string fiberVariable in fiberVariables)
        {
            styleFiber = fiberVariable;
        }

        ////////////////////////
        //Get Position of Text//
        ////////////////////////
        var positionVariables = from c in xmlVariables.Root.Descendants("position")
                                where c.Attribute("sku").Value == Path.GetFileNameWithoutExtension(styleBase)
                                select new
                                {
                                    styleXPosition = (float)c.Element("style").Element("name").Attribute("x"),
                                    styleYPosition = (float)c.Element("style").Element("name").Attribute("y"),
                                    styleTextXPosition = (float)c.Element("style").Element("text").Attribute("x"),
                                    styleTextYPosition = (float)c.Element("style").Element("text").Attribute("y"),
                                    style = (string)c.Element("style").Element("name"),
                                    colorXPosition = (float)c.Element("color").Element("name").Attribute("x"),
                                    colorYPosition = (float)c.Element("color").Element("name").Attribute("y"),
                                    colorTextXPosition = (float)c.Element("color").Element("text").Attribute("x"),
                                    colorTextYPosition = (float)c.Element("color").Element("text").Attribute("y"),
                                    color = (string)c.Element("color").Element("name"),
                                    fiberXPosition = (float)c.Element("fiber").Element("name").Attribute("x"),
                                    fiberYPosition = (float)c.Element("fiber").Element("name").Attribute("y"),
                                    fiberTextXPosition = (float)c.Element("fiber").Element("text").Attribute("x"),
                                    fiberTextYPosition = (float)c.Element("fiber").Element("text").Attribute("y"),
                                    fiber = (string)c.Element("fiber").Element("name"),
                                    widthXPosition = (float)c.Element("width").Element("name").Attribute("x"),
                                    widthYPosition = (float)c.Element("width").Element("name").Attribute("y"),
                                    widthTextXPosition = (float)c.Element("width").Element("text").Attribute("x"),
                                    widthTextYPosition = (float)c.Element("width").Element("text").Attribute("y"),
                                    width = (string)c.Element("width").Element("name"),
                                    patternrepeatXPosition = (float)c.Element("pattern-repeat").Element("name").Attribute("x"),
                                    patternrepeatYPosition = (float)c.Element("pattern-repeat").Element("name").Attribute("y"),
                                    patternrepeatTextXPosition = (float)c.Element("pattern-repeat").Element("text").Attribute("x"),
                                    patternrepeatTextYPosition = (float)c.Element("pattern-repeat").Element("text").Attribute("y"),
                                    patternrepeat = (string)c.Element("pattern-repeat").Element("name"),
                                    patternrepeat2XPosition = (float)c.Element("pattern-repeat2").Element("name").Attribute("x"),
                                    patternrepeat2YPosition = (float)c.Element("pattern-repeat2").Element("name").Attribute("y"),
                                    patternrepeat2 = (string)c.Element("pattern-repeat2").Element("name"),
                                    bugPosition1X = (float)c.Element("bug-position1").Attribute("x"),
                                    bugPosition1Y = (float)c.Element("bug-position1").Attribute("y"),
                                    bugPosition2X = (float)c.Element("bug-position2").Attribute("x"),
                                    bugPosition2Y = (float)c.Element("bug-position2").Attribute("y"),
                                    wonumberTextXPosition = (float)c.Element("wo-number").Attribute("x"),
                                    wonumberTextYPosition = (float)c.Element("wo-number").Attribute("y"),
                                    dateTextXPosition = (float)c.Element("date").Attribute("x"),
                                    dateTextYPosition = (float)c.Element("date").Attribute("y"),
                                    sequencenumberTextXPosition = (float)c.Element("sequence-number").Attribute("x"),
                                    sequencenumberTextYPosition = (float)c.Element("sequence-number").Attribute("y")
                                };
        float styleXPosition = 0;
        float styleYPosition = 0;
        float styleTextXPosition = 0;
        float styleTextYPosition = 0;
        string style = "";
        float colorXPosition = 0;
        float colorYPosition = 0;
        float colorTextXPosition = 0;
        float colorTextYPosition = 0;
        string color = "";
        float fiberXPosition = 0;
        float fiberYPosition = 0;
        float fiberTextXPosition = 0;
        float fiberTextYPosition = 0;
        string fiber = "";
        float widthXPosition = 0;
        float widthYPosition = 0;
        float widthTextXPosition = 0;
        float widthTextYPosition = 0;
        string width = "";
        float patternrepeatXPosition = 0;
        float patternrepeatYPosition = 0;
        float patternrepeat2XPosition = 0;
        float patternrepeat2YPosition = 0;
        float patternrepeatTextXPosition = 0;
        float patternrepeatTextYPosition = 0;
        string patternrepeat = "";
        string patternrepeat2 = "";
        float bugPosition1X = 0;
        float bugPosition1Y = 0;
        float bugPosition2X = 0;
        float bugPosition2Y = 0;
        float wonumberTextXPosition = 0;
        float wonumberTextYPosition = 0;
        float dateTextXPosition = 0;
        float dateTextYPosition = 0;
        float sequencenumberTextXPosition = 0;
        float sequencenumberTextYPosition = 0;
        foreach (var positionVariable in positionVariables)
        {
            styleXPosition = positionVariable.styleXPosition;
            styleYPosition = positionVariable.styleYPosition;
            styleTextXPosition = positionVariable.styleTextXPosition;
            styleTextYPosition = positionVariable.styleTextYPosition;
            style = positionVariable.style;
            colorXPosition = positionVariable.colorXPosition;
            colorYPosition = positionVariable.colorYPosition;
            colorTextXPosition = positionVariable.colorTextXPosition;
            colorTextYPosition = positionVariable.colorTextYPosition;
            color = positionVariable.color;
            fiberXPosition = positionVariable.fiberXPosition;
            fiberYPosition = positionVariable.fiberYPosition;
            fiberTextXPosition = positionVariable.fiberTextXPosition;
            fiberTextYPosition = positionVariable.fiberTextYPosition;
            fiber = positionVariable.fiber;
            widthXPosition = positionVariable.widthXPosition;
            widthYPosition = positionVariable.widthYPosition;
            widthTextXPosition = positionVariable.widthTextXPosition;
            widthTextYPosition = positionVariable.widthTextYPosition;
            width = positionVariable.width;
            patternrepeatXPosition = positionVariable.patternrepeatXPosition;
            patternrepeatYPosition = positionVariable.patternrepeatYPosition;
            patternrepeatTextXPosition = positionVariable.patternrepeatTextXPosition;
            patternrepeatTextYPosition = positionVariable.patternrepeatTextYPosition;
            patternrepeat = positionVariable.patternrepeat;
            patternrepeat2XPosition = positionVariable.patternrepeat2XPosition;
            patternrepeat2YPosition = positionVariable.patternrepeat2YPosition;
            patternrepeat2 = positionVariable.patternrepeat2;
            bugPosition1X = positionVariable.bugPosition1X;
            bugPosition1Y = positionVariable.bugPosition1Y;
            bugPosition2X = positionVariable.bugPosition2X;
            bugPosition2Y = positionVariable.bugPosition2Y;
            wonumberTextXPosition = positionVariable.wonumberTextXPosition;
            wonumberTextYPosition = positionVariable.wonumberTextYPosition;
            dateTextXPosition = positionVariable.dateTextXPosition;
            dateTextYPosition = positionVariable.dateTextYPosition;
            sequencenumberTextXPosition = positionVariable.sequencenumberTextXPosition;
            sequencenumberTextYPosition = positionVariable.sequencenumberTextYPosition;
        }

        ////////////////////////
        //Get Correct Bug Text//     
        ////////////////////////
        string[] bugID = new string[] { "", "", "", "", "", "", "", "", "", "", "", "", "" };
        int bugIndex = 0;
        string[] bugsLine2 = new string[] { "", "", "", "", "", "", "", "", "", "", "", "", "" };
        while (bugIndex < bugs.Count())
        {
            if (bugs[bugIndex] != null)
            {
                var bugVariables = from c in xmlVariables.Root.Descendants("bug")
                                   where (string)c.Attribute("id") == bugs[bugIndex]
                                   select new
                                       {
                                           bugID = (string)c.Attribute("id"),
                                           bugText = (string)c.Element("text"),
                                           bugText2 = (string)c.Element("text2")
                                       };
                foreach (var bugVariable in bugVariables)
                {
                    if ((bugVariable.bugID != "SL32") || (bugVariable.bugID != "SL01") || (bugVariable.bugID != "SL02") || (bugVariable.bugID != "SH43") || (bugVariable.bugID != "SH75"))
                    {
                        bugID[bugIndex] = bugVariable.bugID;
                        bugs[bugIndex] = bugVariable.bugText;
                        bugsLine2[bugIndex] = bugVariable.bugText2;
                    }
                }
                bugIndex++;
            }
        }
        bugIndex = 0;

        ///////////////////
        // Add a new page//
        ///////////////////
        doc.NewPage();

        ////////////////////
        //Create the Label//
        ////////////////////
        PdfReader baseImage = new PdfReader(Settings.Default.tuftexPdf + "\\" + Path.GetFileNameWithoutExtension((styleBase)) + ".pdf");
        PdfImportedPage BaseImagePage = writer.GetImportedPage(baseImage, 1);
        var BaseImageVar = new System.Drawing.Drawing2D.Matrix();
        BaseImageVar.Translate(0f, 0f);
        writer.DirectContentUnder.AddTemplate(BaseImagePage, BaseImageVar);

        /////////////
        //Style//////
        /////////////
        cb.BeginText();
        if (Path.GetFileNameWithoutExtension(styleBase) == "000022021")
        {
            cb.SetFontAndSize(AGarR, 14);
        }
        else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026178")
        {
            cb.SetFontAndSize(GothamBold, 12);
        }
        else
        {
            cb.SetFontAndSize(GillSansB, 12);
        }

        cb.SetTextMatrix(styleXPosition, styleYPosition);
        cb.ShowText(style);
        cb.EndText();

        float styleAdjustment = 0;
        foreach (string s in styleSplit)
        {
            cb.BeginText();
            if (Path.GetFileNameWithoutExtension(styleBase) == "000022021")
            {
                cb.SetFontAndSize(AGarR, 14);
            }
            else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                     Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                     Path.GetFileNameWithoutExtension(styleBase) == "000026178")
            {
                cb.SetFontAndSize(GothamBold, 13);
            }
            else
            {
                cb.SetFontAndSize(GillSansB, 12);
            }
            int styleNameLength = s.Length;
            string styleModified = s;
            if (styleModified.Contains(styleNumber))
            {
                styleModified = styleModified.Substring(0, styleModified.Length - 5) + styleModified.Substring(styleModified.Length - 5).ToUpper();
            }
            cb.SetTextMatrix(styleTextXPosition, styleTextYPosition - styleAdjustment);
            if (styleSplit.Count() > 1)
            {
                if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                     Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                     Path.GetFileNameWithoutExtension(styleBase) == "000026178")
                {
                    cb.ShowText((styleModified.ToUpper()).Replace(" Ii ", " II ").Replace(" Iii ", " III ").Replace(" Ii", " II").Replace(" Iii", " III"));
                }
                else
                {
                    cb.ShowText((styleModified).Replace(" Ii ", " II ").Replace(" Iii ", " III ").Replace(" Ii", " II").Replace(" Iii", " III"));
                }
            }
            else
            {
                if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                     Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                     Path.GetFileNameWithoutExtension(styleBase) == "000026178")
                {
                    cb.ShowText((styleModified.ToUpper()).Replace(" Ii ", " II ").Replace(" Iii ", " III ").Replace(" Ii", " II").Replace(" Iii", " III") + " " + styleNumber);
                }
                else
                {
                    cb.ShowText((styleModified).Replace(" Ii ", " II ").Replace(" Iii ", " III ").Replace(" Ii", " II").Replace(" Iii", " III") + " " + styleNumber);
                }
            }
            //cb.ShowText((styleName + " " + styleNumber).Replace(" Ii ", " II ").Replace(" Iii ", " III "));
            cb.EndText();
            styleAdjustment = styleAdjustment + 16.992f;
            styleName = styleModified;
        }

        styleAdjustment = styleAdjustment - 16.992f;

        //////////////////////
        //Color///////////////
        //////////////////////
        cb.BeginText();
        if (Path.GetFileNameWithoutExtension(styleBase) == "000022021")
        {
            cb.SetFontAndSize(AGarR, 14);
        }
        else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026178")
        {
            cb.SetFontAndSize(GothamBold, 12);
        }
        else
        {
            cb.SetFontAndSize(GillSansB, 12);
        }
        cb.SetTextMatrix(colorXPosition, colorYPosition - styleAdjustment);
        cb.ShowText(color);
        cb.EndText();
        cb.BeginText();
        if (Path.GetFileNameWithoutExtension(styleBase) == "000022021")
        {
            cb.SetFontAndSize(AGarR, 14);
        }
        else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026178")
        {
            cb.SetFontAndSize(GothamBook, 12);
        }
        else
        {
            cb.SetFontAndSize(GillSansR, 12);
        }
        cb.SetTextMatrix(colorTextXPosition, colorTextYPosition - styleAdjustment);
        cb.ShowText(colorNumber + " " + colorName);
        cb.EndText();

        //////////////////////
        //Fiber///////////////
        //////////////////////
        cb.BeginText();
        if (Path.GetFileNameWithoutExtension(styleBase) == "000022021")
        {
            cb.SetFontAndSize(AGarR, 14);
        }
        else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026178")
        {
            if (fiber.Length < 40)
            {
                cb.SetFontAndSize(GothamBold, 12);
            }
            else
            {
                cb.SetFontAndSize(GothamBold, 11);
            }
        }
        else
        {
            cb.SetFontAndSize(GillSansB, 12);
        }
        cb.SetTextMatrix(fiberXPosition, fiberYPosition - styleAdjustment);
        cb.ShowText(fiber);
        cb.EndText();
        cb.BeginText();
        if (Path.GetFileNameWithoutExtension(styleBase) == "000022021")
        {
            cb.SetFontAndSize(AGarR, 14);
        }
        else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026178")
        {
            cb.SetFontAndSize(GothamBook, 12);
        }
        else
        {
            cb.SetFontAndSize(GillSansR, 12);
        }
        cb.SetTextMatrix(fiberTextXPosition, fiberTextYPosition - styleAdjustment);
        cb.ShowText(styleFiber
            .Replace("STAINMASTER", "Stainmaster")
            .Replace("Luxerell", "Luxerell\u2122")
            .Replace("LUXERELL", "Luxerell\u2122")
            .Replace("NYLON", "Nylon")
            .Replace("ANSO", "Anso")
            .Replace("EXTRABODY II", "ExtraBody II\u2122")
            .Replace("(TM)", "\u2122")
            .Replace("(R)", "\u00AE")
            );
        cb.EndText();

        //////////////////////
        //Width///////////////
        //////////////////////
        cb.BeginText();
        if (Path.GetFileNameWithoutExtension(styleBase) == "000022021")
        {
            cb.SetFontAndSize(AGarR, 14);
        }
        else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026178")
        {
            cb.SetFontAndSize(GothamBold, 12);
        }
        else
        {
            cb.SetFontAndSize(GillSansB, 12);
        }
        cb.SetTextMatrix(widthXPosition, widthYPosition - styleAdjustment);
        cb.ShowText(width);
        cb.EndText();
        cb.BeginText();
        if (Path.GetFileNameWithoutExtension(styleBase) == "000022021")
        {
            cb.SetFontAndSize(AGarR, 14);
        }
        else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026178")
        {
            cb.SetFontAndSize(GothamBook, 12);
        }
        else
        {
            cb.SetFontAndSize(GillSansR, 12);
        }
        cb.SetTextMatrix(widthTextXPosition, widthTextYPosition - styleAdjustment);
        cb.ShowText(styleWidth);
        cb.EndText();

        //////////////////////
        //Pattern Repeat//////
        //////////////////////
        cb.BeginText();
        if (Path.GetFileNameWithoutExtension(styleBase) == "000022021")
        {
            cb.SetFontAndSize(AGarR, 14);
        }
        else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026178")
        {
            cb.SetFontAndSize(GothamBold, 12);
        }
        else
        {
            cb.SetFontAndSize(GillSansB, 12);
        }
        cb.SetTextMatrix(patternrepeatXPosition, patternrepeatYPosition - styleAdjustment);
        cb.ShowText(patternrepeat);
        cb.EndText();
        cb.BeginText();
        if (Path.GetFileNameWithoutExtension(styleBase) == "000022021")
        {
            cb.SetFontAndSize(AGarR, 14);
        }
        else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                 Path.GetFileNameWithoutExtension(styleBase) == "000026178")
        {
            cb.SetFontAndSize(GothamBook, 12);
        }
        else
        {
            cb.SetFontAndSize(GillSansR, 12);
        }
        cb.SetTextMatrix(patternrepeatTextXPosition, patternrepeatTextYPosition - styleAdjustment);
        cb.ShowText(stylePatternRepeat);
        cb.EndText();
        ////////////////////
        //Line Two//////////
        ////////////////////        
        if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
            Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
            Path.GetFileNameWithoutExtension(styleBase) == "000026178")
        {
            cb.BeginText();
            cb.SetFontAndSize(GothamBold, 12);
            cb.SetTextMatrix(patternrepeat2XPosition, patternrepeat2YPosition - styleAdjustment);
            cb.ShowText(patternrepeat2);
            cb.EndText();
        }
        //**********************//
        //Performance Rating Bar//
        //**********************//

        if ((Path.GetFileNameWithoutExtension(styleBase) == "000026046") ||
            (Path.GetFileNameWithoutExtension(styleBase) == "000026047") ||
            (Path.GetFileNameWithoutExtension(styleBase) == "000026048"))
        {
            float ratingIncrement = 4.5558f;

            float pratingValueXvar = (((float.Parse(durability) * 10) - 10) * ratingIncrement);
            /*
            PdfReader pRating = new PdfReader(tuftexPdfPath + "prating.pdf");
            PdfImportedPage pRatingPage = writer.GetImportedPage(pRating, 1);
            var pRatingVar = new System.Drawing.Drawing2D.Matrix();
            pRatingVar.Translate(36.72f, 290f);
            writer.DirectContentUnder.AddTemplate(pRatingPage, pRatingVar);
            */
            PdfReader arrow = new PdfReader(Settings.Default.tuftexPdf + "\\" + "arrow.pdf");
            PdfImportedPage arrowPage = writer.GetImportedPage(arrow, 1);
            var arrowVar = new System.Drawing.Drawing2D.Matrix();
            arrowVar.Translate(35f + pratingValueXvar, 281f);
            writer.DirectContentUnder.AddTemplate(arrowPage, arrowVar);

            cb.BeginText();
            cb.SetFontAndSize(GillSansR, 7);
            cb.ShowTextAligned(Element.ALIGN_CENTER, float.Parse(durability).ToString("0.00"), 35f + pratingValueXvar + 4.6332f, 275, 0);
            cb.EndText();
            if ((Path.GetFileNameWithoutExtension(styleBase) != "000026046"))
            {
                bugPosition1Y = bugPosition1Y - 67;
            }
        }
        else if ((Path.GetFileNameWithoutExtension(styleBase) == "000026173"))
        {
            float ratingIncrement = 7.0045f;

            float pratingValueXvar = (((float.Parse(durability) * 10) - 10) * ratingIncrement);
            /*
            PdfReader pRating = new PdfReader(tuftexPdfPath + "prating.pdf");
            PdfImportedPage pRatingPage = writer.GetImportedPage(pRating, 1);
            var pRatingVar = new System.Drawing.Drawing2D.Matrix();
            pRatingVar.Translate(36.72f, 290f);
            writer.DirectContentUnder.AddTemplate(pRatingPage, pRatingVar);
            */
            PdfReader arrow = new PdfReader(Settings.Default.tuftexPdf + "\\" + "arrow.pdf");
            PdfImportedPage arrowPage = writer.GetImportedPage(arrow, 1);
            var arrowVar = new System.Drawing.Drawing2D.Matrix();
            arrowVar.Translate(50f + pratingValueXvar, 193.4f);
            writer.DirectContentUnder.AddTemplate(arrowPage, arrowVar);

            cb.BeginText();
            cb.SetFontAndSize(GothamBold, 5);
            cb.ShowTextAligned(Element.ALIGN_CENTER, float.Parse(durability).ToString("0.00"), 50f + pratingValueXvar + 4.6332f, 187.92f, 0);
            cb.EndText();
            if ((Path.GetFileNameWithoutExtension(styleBase) != "000026173"))
            {
                bugPosition1Y = bugPosition1Y - 67;
            }
        }

        //**********************//
        //**********BUGS********//
        //**********************//
        if ((bugs[0] != null) && (Path.GetFileNameWithoutExtension(styleBase) != "000022021"))
        {
            while (bugIndex < bugs.Count())
            {
                cb.BeginText();
                if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                    Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                    Path.GetFileNameWithoutExtension(styleBase) == "000026178")
                {
                    cb.SetFontAndSize(GothamBook, 11);
                }
                else
                {
                    cb.SetFontAndSize(GillSansB, 9);
                }
                cb.SetTextMatrix(bugPosition1X, bugPosition1Y);
                if (bugs[bugIndex] == "SL32")
                {
                    PdfReader bugImageSL32 = new PdfReader(Settings.Default.tuftexPdf + "\\" + "SL32.pdf");
                    PdfImportedPage bugImageSL32Page = writer.GetImportedPage(bugImageSL32, 1);
                    var bugImageSL32Var = new System.Drawing.Drawing2D.Matrix();
                    bugImageSL32Var.Translate(194.8464f, 65.6568f);
                    writer.DirectContent.AddTemplate(bugImageSL32Page, bugImageSL32Var);
                }
                else if (bugs[bugIndex] == "SL01")
                {
                    PdfReader bugImageSL01 = new PdfReader(Settings.Default.tuftexPdf + "\\" + "SL01.pdf");
                    PdfImportedPage bugImageSL01Page = writer.GetImportedPage(bugImageSL01, 1);
                    var bugImageSL01Var = new System.Drawing.Drawing2D.Matrix();
                    bugImageSL01Var.Translate(194.8464f, 65.6568f);
                    writer.DirectContent.AddTemplate(bugImageSL01Page, bugImageSL01Var);
                }
                else if (bugs[bugIndex] == "SL02")
                {
                    PdfReader bugImageSL02 = new PdfReader(Settings.Default.tuftexPdf + "\\" + "SL02.pdf");
                    PdfImportedPage bugImageSL02Page = writer.GetImportedPage(bugImageSL02, 1);
                    var bugImageSL02Var = new System.Drawing.Drawing2D.Matrix();
                    bugImageSL02Var.Translate(194.8464f, 65.6568f);
                    writer.DirectContent.AddTemplate(bugImageSL02Page, bugImageSL02Var);
                }
                else if (bugs[bugIndex] == "SH43")
                {
                    PdfReader bugImageSH43 = new PdfReader(Settings.Default.tuftexPdf + "\\" + "SH43.pdf");
                    PdfImportedPage bugImageSH43Page = writer.GetImportedPage(bugImageSH43, 1);
                    var bugImageSH43Var = new System.Drawing.Drawing2D.Matrix();
                    bugImageSH43Var.Translate(441f, 303.2784f);
                    writer.DirectContent.AddTemplate(bugImageSH43Page, bugImageSH43Var);
                    bugPosition1Y = bugPosition1Y - 63f;
                }
                else if (bugs[bugIndex] == "SH75")
                {
                    PdfReader bugImageSH75 = new PdfReader(Settings.Default.tuftexPdf + "\\" + "SH75.pdf");
                    PdfImportedPage bugImageSH75Page = writer.GetImportedPage(bugImageSH75, 1);
                    var bugImageSH75Var = new System.Drawing.Drawing2D.Matrix();
                    bugImageSH75Var.Translate(431.44f, 308f);
                    writer.DirectContent.AddTemplate(bugImageSH75Page, bugImageSH75Var);
                    bugPosition1Y = bugPosition1Y - 63f;
                }
                else if ((bugID[bugIndex] == "SA01") ||
                         (bugID[bugIndex] == "SA02") ||
                         (bugID[bugIndex] == "SA15") ||
                         (bugID[bugIndex] == "SH72") ||
                         (bugID[bugIndex] == "SH73") ||
                         (bugID[bugIndex] == "SH74"))
                {
                    if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                        Path.GetFileNameWithoutExtension(styleBase) == "000026173")
                    {
                        cb.SetFontAndSize(GothamBook, 11);
                        cb.SetTextMatrix(bugPosition2X, bugPosition2Y);
                        cb.ShowText(bugs[bugIndex].Replace("bulletdot", "\u2022" + " "));
                        bugPosition2Y = bugPosition2Y - 15f;
                    }
                    else
                    {
                        cb.SetFontAndSize(UniCB, 12);
                        cb.SetTextMatrix(bugPosition2X, bugPosition2Y);
                        cb.ShowText(bugs[bugIndex].Replace("bulletdot", "\u2022" + " "));
                        bugPosition2Y = bugPosition2Y - 15.84f;
                    }
                }
                else if (Path.GetFileNameWithoutExtension(styleBase) == "000019729" ||
                         Path.GetFileNameWithoutExtension(styleBase) == "000026046")
                {
                    cb.SetFontAndSize(GillSansB, 9);
                    cb.SetTextMatrix(bugPosition1X, bugPosition1Y);
                    cb.ShowText(bugs[bugIndex]);
                    bugPosition1Y = bugPosition1Y - 16.92f;
                }
                else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                         Path.GetFileNameWithoutExtension(styleBase) == "000026173")
                {
                    cb.SetFontAndSize(GothamBook, 11);
                    cb.SetTextMatrix(bugPosition1X, bugPosition1Y);
                    cb.ShowText(bugs[bugIndex]);
                    bugPosition1Y = bugPosition1Y - 20f;
                }
                else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                         Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                         Path.GetFileNameWithoutExtension(styleBase) == "000026178")
                {
                    cb.SetFontAndSize(GothamBook, 11);
                    cb.SetTextMatrix(bugPosition1X, bugPosition1Y);
                    cb.ShowText(bugs[bugIndex]);
                    bugPosition1Y = bugPosition1Y - 16.92f;
                }
                else if (bugID[bugIndex] == "SA22" ||
                         bugID[bugIndex] == "SH78")
                {
                    cb.ShowText(bugs[bugIndex]);
                    bugPosition1Y = bugPosition1Y - 11.136f;
                    cb.EndText();

                    cb.BeginText();
                    if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                        Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                        Path.GetFileNameWithoutExtension(styleBase) == "000026178")
                    {
                        cb.SetFontAndSize(GothamBook, 11);
                    }
                    else
                    {
                        cb.SetFontAndSize(GillSansB, 9);
                    }
                    cb.SetTextMatrix(bugPosition1X, bugPosition1Y);
                    cb.ShowText(bugsLine2[bugIndex]);
                    bugPosition1Y = bugPosition1Y - 17.136f;
                }
                else
                {
                    if (Path.GetFileNameWithoutExtension(styleBase) == "000026178")
                    {
                        cb.ShowText(bugs[bugIndex]);
                        bugPosition1Y = bugPosition1Y - 19f;
                    }
                    else
                    {
                        cb.ShowText(bugs[bugIndex]);
                        bugPosition1Y = bugPosition1Y - 17.136f;
                    }
                }
                cb.EndText();
                bugIndex++;
            }
            if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                Path.GetFileNameWithoutExtension(styleBase) == "000026173")
            {
                cb.BeginText();
                cb.SetFontAndSize(GothamLight, 8);
                cb.SetTextMatrix(443.52f, bugPosition2Y);
                cb.ShowText("See warranty brochure for warranty details.");
                cb.EndText();
            }
        }
        if (!woNumber.ToLower().Contains("create"))
        {
            //ExtraInfo
            cb.BeginText();
            if (Path.GetFileNameWithoutExtension(styleBase) == "000022021")
            {
                cb.SetFontAndSize(AGarR, 9);
            }
            else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                     Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                     Path.GetFileNameWithoutExtension(styleBase) == "000026178")
            {
                cb.SetFontAndSize(GothamBook, 8);
            }
            else
            {
                cb.SetFontAndSize(GillSansR, 9);
            }
            cb.SetCMYKColorFill(0, 0, 0, 255);
            cb.SetTextMatrix(wonumberTextXPosition, wonumberTextYPosition);
            cb.ShowText(woNumber);
            cb.EndText();
            cb.BeginText();
            if (Path.GetFileNameWithoutExtension(styleBase) == "000022021")
            {
                cb.SetFontAndSize(AGarR, 9);
            }
            else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                     Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                     Path.GetFileNameWithoutExtension(styleBase) == "000026178")
            {
                cb.SetFontAndSize(GothamBook, 8);
            }
            else
            {
                cb.SetFontAndSize(GillSansR, 9);
            }
            cb.SetCMYKColorFill(0, 0, 0, 255);
            cb.SetTextMatrix(dateTextXPosition, dateTextYPosition);
            cb.ShowText(DateTime.Now.ToString("MM/dd/yyyy"));
            cb.EndText();
            cb.BeginText();
            if (Path.GetFileNameWithoutExtension(styleBase) == "000022021")
            {
                cb.SetFontAndSize(AGarR, 9);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(sequencenumberTextXPosition, sequencenumberTextYPosition);
                cb.ShowText("DISPLAY SEQUENCE: " + colorSequence);
                cb.EndText();
            }
            else if (Path.GetFileNameWithoutExtension(styleBase) == "000026172" ||
                     Path.GetFileNameWithoutExtension(styleBase) == "000026173" ||
                     Path.GetFileNameWithoutExtension(styleBase) == "000026178")
            {
                cb.SetFontAndSize(GothamBook, 8);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(sequencenumberTextXPosition, sequencenumberTextYPosition);
                cb.ShowText("Display Sequence " + colorSequence);
                cb.EndText();
            }
            else
            {
                cb.SetFontAndSize(GillSansR, 9);
                cb.SetCMYKColorFill(0, 0, 0, 255);
                cb.SetTextMatrix(sequencenumberTextXPosition, sequencenumberTextYPosition);
                cb.ShowText("DISPLAY SEQUENCE: " + colorSequence);
                cb.EndText();
            }

        }
    }
}