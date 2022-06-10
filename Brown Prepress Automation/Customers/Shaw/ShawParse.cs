using Brown_Prepress_Automation.Properties;
using ExcelLibrary.SpreadSheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace Brown_Prepress_Automation
{
    class ShawParse
    {
        PdfProcessing pdfProcessing = new PdfProcessing();
        MethodsCommon methodsCommon = new MethodsCommon();


        //string pdfFile = "Wo's Combined.pdf";
        string path = @"C:\Users\mhoward\Desktop\Brown Automation Testing\WO Parsing\";

        public class WorkOrders
        {
            public string Wo { get; set; }
            public string AssemblyOutsource { get; set; }
            public string StyleNumber { get; set; }
            public string CustomerNumber { get; set; }
            public string CustomerSeqNumber { get; set; }
            public string Qty { get; set; }
            public string DeliveryVehicle { get; set; }
            public string BaseName { get; set; }
            public bool LookBook { get; set; }
            public bool PhotoBook { get; set; }
            public string Panels { get; set; }
            public List<string> Skus { get; set; } = new List<string>();
            public List<int> PageNumbers { get; set; } = new List<int>();
            public List<string> ColorNumbers { get; set; } = new List<string>();
        }
        public void buildShawSpreadSheet(string poFile)
        {
            string pdfFile = Path.GetFileName(poFile);
            path = Path.GetDirectoryName(poFile) + "\\";
            List<WorkOrders> parsedPdf = ParseWoPdf(pdfFile);
            string file = path + Path.GetFileNameWithoutExtension(pdfFile) + ".xls";
            int ssLine = 1;
            Workbook book = new Workbook();
            Worksheet sheet = new Worksheet("Sheet");
            sheet.Cells[0, 0] = new Cell("FileName");
            sheet.Cells[0, 1] = new Cell("Page");
            sheet.Cells[0, 2] = new Cell("Color");
            sheet.Cells[0, 3] = new Cell("Part Number");
            sheet.Cells[0, 4] = new Cell("Size");
            sheet.Cells[0, 5] = new Cell("Qty");
            sheet.Cells[0, 6] = new Cell("WO");
            sheet.Cells[0, 7] = new Cell("SO Line");
            sheet.Cells[0, 8] = new Cell("Stock and Finishing");

            foreach (WorkOrders wo in parsedPdf)
            {
                for (int i = 0; i < wo.Skus.Count(); i++)
                {
                    if (!wo.Wo.ToLower().Contains("b0"))
                    {
                        for (int z = 0; z < wo.ColorNumbers.Count(); z++)
                        {
                            sheet.Cells[ssLine, 0] = new Cell(wo.StyleNumber + " " +
                                //wo.BaseName + " " +
                                //wo.DeliveryVehicle + " " +
                                wo.CustomerNumber + " " +
                                wo.CustomerSeqNumber + " " +
                                wo.Skus[i].TrimStart('0')
                            );
                            sheet.Cells[ssLine, 1] = new Cell(wo.PageNumbers[z]);
                            sheet.Cells[ssLine, 2] = new Cell(wo.ColorNumbers[z]);
                            sheet.Cells[ssLine, 3] = new Cell("");
                            sheet.Cells[ssLine, 4] = new Cell("");
                            sheet.Cells[ssLine, 5] = new Cell(wo.Qty);
                            sheet.Cells[ssLine, 6] = new Cell(wo.Wo);
                            sheet.Cells[ssLine, 7] = new Cell("");
                            sheet.Cells[ssLine, 8] = new Cell("");
                            ssLine++;
                        }
                    }
                    else
                    {
                        sheet.Cells[ssLine, 0] = new Cell(wo.StyleNumber + " " +
                                //wo.BaseName + " " +
                                wo.DeliveryVehicle + " " +
                                wo.CustomerNumber + " " +
                                wo.CustomerSeqNumber + " " +
                                wo.Skus[i].TrimStart('0')
                            );
                        if (Int32.Parse(wo.Panels) > 1)
                        {
                            sheet.Cells[ssLine, 1] = new Cell("1");
                        }
                        sheet.Cells[ssLine, 3] = new Cell("");
                        sheet.Cells[ssLine, 4] = new Cell("");
                        sheet.Cells[ssLine, 5] = new Cell(wo.Qty);
                        sheet.Cells[ssLine, 6] = new Cell(wo.Wo);
                        sheet.Cells[ssLine, 7] = new Cell("");
                        sheet.Cells[ssLine, 8] = new Cell("");
                        ssLine++;
                        if (wo.LookBook || wo.PhotoBook)
                        {
                            sheet.Cells[ssLine, 0] = new Cell(wo.StyleNumber + " " +
                                //wo.BaseName + " " +
                                wo.DeliveryVehicle + " " +
                                wo.CustomerNumber + " " +
                                wo.CustomerSeqNumber + " " +
                                wo.Skus[i].TrimStart('0')
                            );
                            if (wo.PageNumbers.Count() != 0)
                            {
                                sheet.Cells[ssLine, 1] = new Cell(wo.PageNumbers[i]);
                            }
                            if (wo.ColorNumbers.Count() != 0)
                            {
                                sheet.Cells[ssLine, 2] = new Cell(wo.ColorNumbers[i]);
                            }
                            sheet.Cells[ssLine, 3] = new Cell("");
                            sheet.Cells[ssLine, 4] = new Cell("");
                            sheet.Cells[ssLine, 5] = new Cell(wo.Qty);
                            sheet.Cells[ssLine, 6] = new Cell(wo.Wo);
                            sheet.Cells[ssLine, 7] = new Cell("");
                            if (wo.LookBook)
                            {
                                sheet.Cells[ssLine, 8] = new Cell("Book - ");
                            }
                            if (wo.PhotoBook)
                            {
                                sheet.Cells[ssLine, 8] = new Cell("PhotoPack - ");
                            }
                            ssLine++;
                        }
                        else if (Int32.Parse(wo.Panels) > 1)
                        {
                            int boardPage = 3;
                            for (int z = 1; z < Int32.Parse(wo.Panels); z++)
                            {
                                sheet.Cells[ssLine, 0] = new Cell(wo.StyleNumber + " " +
                                    //wo.BaseName + " " +
                                    wo.DeliveryVehicle + " " +
                                    wo.CustomerNumber + " " +
                                    wo.CustomerSeqNumber + " " +
                                    wo.Skus[i].TrimStart('0')
                                );
                                sheet.Cells[ssLine, 1] = new Cell(boardPage);
                                sheet.Cells[ssLine, 2] = new Cell("");
                                sheet.Cells[ssLine, 3] = new Cell("");
                                sheet.Cells[ssLine, 4] = new Cell("");
                                sheet.Cells[ssLine, 5] = new Cell(wo.Qty);
                                sheet.Cells[ssLine, 6] = new Cell(wo.Wo);
                                sheet.Cells[ssLine, 7] = new Cell("");
                                sheet.Cells[ssLine, 8] = new Cell("");
                                ssLine++;
                                boardPage = boardPage + 2;                                
                            }
                        }
                    }
                }
            }
            for (int i = 0; i < 100; i++)
            {
                sheet.Cells[i, 20] = new Cell("");
            }
            book.Worksheets.Add(sheet);
            book.Save(file);
            //MessageBox.Show("Done");
        }

        public List<WorkOrders> ParseWoPdf(string fileName)
        {
            WorkOrders WorkOrder = new WorkOrders();
            List<WorkOrders> woList = new List<WorkOrders>();

            string wo = "";
            string assemblyOutsource = "";
            string styleNumber = "";
            string customerNo = "";
            string custSeqNo = "";
            string qty = "";
            string delVeh = "";
            string baseName = "";
            bool lookBook = false;
            bool printShop = true;
            bool photoBook = false;
            bool instructionLine = false;
            string panels = "1";
            int photoLineCheck = 1;
            List<string> skus = new List<string>();
            List<int> pageNumbers = new List<int>();
            List<string> colorNumbers = new List<string>();
            string file = path + fileName;
            string txtFile = path + Path.GetFileNameWithoutExtension(fileName) + ".txt";
            string text = pdfProcessing.pdfToText(file);
            File.WriteAllText(txtFile, text);
            string[] pdfText = text.Split('\n');
            foreach (string line in pdfText)
            {
                if (line.ToLower().Contains("bindery - instructions sheet"))
                {
                    printShop = false;
                }
                if (printShop)
                {
                    if (line.Contains("Work Order:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            if (splitLine[i] == "Order:")
                            {
                                //MessageBox.Show(splitLine[i + 1]);
                                wo = splitLine[i + 1];
                                WorkOrder.Wo = wo;
                            }
                        }
                    }
                    if (line.Contains("Assembly Outsource:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            if (splitLine[i] == "Outsource:")
                            {
                                //MessageBox.Show(splitLine[i + 1]);
                                assemblyOutsource = splitLine[i + 1];
                                WorkOrder.AssemblyOutsource = assemblyOutsource;
                            }
                        }
                    }
                    if (line.Contains("Sell Style:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            if (splitLine[i] == "Style:")
                            {
                                //MessageBox.Show(splitLine[i + 1]);
                                styleNumber = splitLine[i + 1];
                                WorkOrder.StyleNumber = styleNumber;
                            }
                        }
                    }
                    if (line.Contains("Customer No:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            int n;
                            if (i + 1 < splitLine.Count())
                            {
                                if ((splitLine[i] == "No:") && (int.TryParse(splitLine[i + 1], out n)))
                                {
                                    //MessageBox.Show(splitLine[i + 1]);
                                    customerNo = splitLine[i + 1].TrimStart('0');
                                    WorkOrder.CustomerNumber = customerNo;
                                }
                            }
                        }
                    }
                    if (line.Contains("Cust Seq No:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            int n;
                            if (i + 1 < splitLine.Count())
                            {
                                if ((splitLine[i] == "No:") && (int.TryParse(splitLine[i + 1], out n)))
                                {
                                    //MessageBox.Show(splitLine[i + 1]);
                                    custSeqNo = splitLine[i + 1];
                                    WorkOrder.CustomerSeqNumber = custSeqNo;
                                }
                            }
                        }
                    }
                    if (line.Contains("WO Qty:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            int n;
                            if (i + 1 < splitLine.Count())
                            {
                                if ((splitLine[i] == "Qty:") && (int.TryParse(splitLine[i + 1], out n)))
                                {
                                    //MessageBox.Show(splitLine[i + 1]);
                                    qty = splitLine[i + 1];
                                    WorkOrder.Qty = qty;
                                }
                            }
                        }
                    }
                    if (line.Contains("Del Veh:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            if (i + 1 < splitLine.Count())
                            {
                                if ((splitLine[i] == "Veh:"))
                                {
                                    //MessageBox.Show(splitLine[i + 1]);
                                    delVeh = splitLine[i + 1];
                                    WorkOrder.DeliveryVehicle = delVeh;
                                }
                            }
                        }
                    }
                    if (line.Contains("0000"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            int n;
                            if (i + 1 < splitLine.Count())
                            {
                                if ((splitLine[i].Length == 9) && (int.TryParse(splitLine[i], out n)) && (!splitLine.Contains("DUMMY")))
                                {
                                    //MessageBox.Show(splitLine[i]);
                                    //string skuLine = splitLine[i];
                                    //skus.Add(splitLine[i]);
                                    WorkOrder.Skus.Add(splitLine[i]);
                                }
                            }
                        }
                    }
                    if (wo != "")
                    {
                        if (!Directory.Exists(Settings.Default.tempDir + "Shaw\\WOs\\"))
                        {
                            Directory.CreateDirectory(Settings.Default.tempDir + "Shaw\\WOs\\");
                        }
                        if (!File.Exists(Settings.Default.tempDir + "Shaw\\WOs\\" + wo + ".xml"))
                        {
                            methodsCommon.xmlDownloadShaw(wo, Settings.Default.tempDir + "Shaw\\WOs\\" + wo + ".xml");
                        }

                        var xmlData = ShawSalsaParse(wo, styleNumber);
                        for (int i = 0; i < xmlData.colors.Count(); i++)
                        {

                            if (line.Contains(xmlData.colors[i]))
                            {
                                //colorNumbers.Add(xmlData.colors[i]);
                                //pageNumbers.Add(i + 1);
                                WorkOrder.ColorNumbers.Add(xmlData.colorNumbers[i]);
                                WorkOrder.PageNumbers.Add(i + 1);
                            }
                        }
                        WorkOrder.BaseName = baseName = xmlData.baseName;
                    }
                    if (line.Contains("Page:"))
                    {
                        string[] splitLine = line.Split(' ');
                        if (splitLine[4] == splitLine[6])
                        {
                            woList.Add(WorkOrder);
                            WorkOrder = new WorkOrders();
                            skus.Clear();
                        }
                    }
                }
                else
                {
                    if (line.Contains("Work Order:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            if (splitLine[i] == "Order:")
                            {
                                //MessageBox.Show(splitLine[i + 1]);
                                wo = splitLine[i + 1];
                                WorkOrder.Wo = wo;
                            }
                        }
                    }
                    if (line.Contains("Assembly Outsource:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            if (splitLine[i] == "Outsource:")
                            {
                                //MessageBox.Show(splitLine[i + 1]);
                                assemblyOutsource = splitLine[i + 1];
                                WorkOrder.AssemblyOutsource = assemblyOutsource;
                            }
                        }
                    }
                    if (line.Contains("Sell Style:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            if (splitLine[i] == "Style:")
                            {
                                //MessageBox.Show(splitLine[i + 1]);
                                styleNumber = splitLine[i + 1];
                                WorkOrder.StyleNumber = styleNumber;
                            }
                        }
                    }
                    if (line.Contains("Customer No:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            int n;
                            if (i + 1 < splitLine.Count())
                            {
                                if ((splitLine[i] == "No:") && (int.TryParse(splitLine[i + 1], out n)))
                                {
                                    if (splitLine[i - 1] == "Customer")
                                    {
                                        //MessageBox.Show(splitLine[i + 1]);
                                        customerNo = splitLine[i + 1].TrimStart('0');
                                        WorkOrder.CustomerNumber = customerNo;
                                    }
                                }
                            }
                        }
                    }
                    if (line.Contains("Sequence No:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            int n;
                            if (i + 1 < splitLine.Count())
                            {
                                if ((splitLine[i] == "No:") && (int.TryParse(splitLine[i + 1], out n)))
                                {
                                    if (splitLine[i - 1] == "Sequence")
                                    {
                                        //MessageBox.Show(splitLine[i + 1]);
                                        custSeqNo = splitLine[i + 1];
                                        WorkOrder.CustomerSeqNumber = custSeqNo;
                                    }
                                }
                            }
                        }
                    }
                    if (line.Contains("Quantity:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            int n;
                            if (i + 1 < splitLine.Count())
                            {
                                if ((splitLine[i] == "Quantity:") && (int.TryParse(splitLine[i + 1], out n)))
                                {
                                    //MessageBox.Show(splitLine[i + 1]);
                                    qty = splitLine[i + 1];
                                    WorkOrder.Qty = qty;
                                }
                            }
                        }
                    }
                    if (line.Contains("Del Veh No:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            if (i + 1 < splitLine.Count())
                            {
                                if ((splitLine[i] == "No:"))
                                {
                                    //MessageBox.Show(splitLine[i + 1]);
                                    delVeh = splitLine[i + 1];
                                    WorkOrder.DeliveryVehicle = delVeh;
                                }
                            }
                        }
                    }
                    if (line.Contains("Panels:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            int n;
                            if (i + 1 < splitLine.Count())
                            {
                                if ((splitLine[i] == "Panels:") && (int.TryParse(splitLine[i + 1], out n)) && (Int32.Parse(splitLine[i + 1]) > 1))
                                {
                                    if (!Directory.Exists(Settings.Default.tempDir + "Shaw\\WOs\\"))
                                    {
                                        Directory.CreateDirectory(Settings.Default.tempDir + "Shaw\\WOs\\");
                                    }
                                    if (!File.Exists(Settings.Default.tempDir + "Shaw\\WOs\\" + wo + ".xml"))
                                    {
                                        methodsCommon.xmlDownloadShaw(wo, Settings.Default.tempDir + "Shaw\\WOs\\" + wo + ".xml");
                                    }

                                    string xmlData = ShawSalsaParseBoard(wo);

                                    if (xmlData.ToLower().Contains("board"))
                                    {
                                        //MessageBox.Show(splitLine[i + 1]);
                                        panels = splitLine[i + 1];
                                        
                                    }
                                }
                            }
                        }
                    }
                    WorkOrder.Panels = panels;
                    if (line.Contains("Instructions:"))
                    {
                        /*string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            if (splitLine[i].ToLower() == "lookbook")
                            {
                                //MessageBox.Show(splitLine[i + 1]);
                                lookBook = true;
                            }                            
                        }*/
                        //instructionLine = true;
                        //WorkOrder.LookBook = lookBook;                        
                    }
                    if (instructionLine)
                    {
                        photoLineCheck++;
                    }
                    if (line.ToLower().Contains("waterfall"))
                        //if ((line.ToLower().Contains("waterfall")) && (photoLineCheck < 7))
                    {
                        photoBook = true;
                        WorkOrder.PhotoBook = photoBook;
                        //instructionLine = false;
                    }
                    if (line.ToLower().Contains("lookbook"))
                    {
                        lookBook = true;
                        WorkOrder.LookBook = lookBook; 
                    }
                    if (line.Contains("Bindery SKU:"))
                    {
                        string[] splitLine = line.Split(' ');
                        for (int i = 0; i < splitLine.Count(); i++)
                        {
                            if (splitLine[i] == "SKU:")
                            {
                                //MessageBox.Show(splitLine[i + 1]);
                                //skus.Add(splitLine[i + 1]);
                                WorkOrder.Skus.Add(splitLine[i + 1]);
                            }
                        }
                    }
                    if (line.Contains("Page:"))
                    {
                        string[] splitLine = line.Split(' ');
                        if (splitLine[4] == splitLine[6])
                        {
                            woList.Add(WorkOrder);
                            WorkOrder = new WorkOrders();
                            skus.Clear();
                            lookBook = false;
                            photoBook = false;
                            printShop = true;
                            instructionLine = false;
                            panels = "1";
                        }
                    }
                }
            }
            //buildShawSpreadSheet(woList);
            if (Directory.Exists(Settings.Default.tempDir + "Shaw\\WOs\\"))
            {
                Directory.Delete(Settings.Default.tempDir + "Shaw\\WOs\\", true);
            }
            if (File.Exists(path + Path.GetFileNameWithoutExtension(fileName) + ".txt"))
            {
                File.Delete(path + Path.GetFileNameWithoutExtension(fileName) + ".txt");
            }
            return woList;
        }

        public (string wo, string styleNo, string baseName, List<string> colorNumbers, List<string> colors) ShawSalsaParse(string wo, string style)
        {
            string file = Settings.Default.tempDir + "Shaw\\WOs\\" + wo + ".xml";
            //string file = "http://salsaprd.shawinc.com/SALSAWeb/ServiceRequest?service=getSpec&id=" + wo + "&nostatusupdate=true";
            string woStyle = string.Empty;
            string woReturn = string.Empty;
            string styleNoReturn = string.Empty;
            string baseNameReturn = string.Empty;
            List<string> colorNumbersReturn = new List<string>();
            List<string> colorNamesReturn = new List<string>();
            List<string> colorsCombinedReturn = new List<string>();

            XElement xml = XElement.Load(file);
            var woData = xml
                        .Descendants("style")
                        .Where(el => el.Attribute("id").Value == style)
                        .Select(woItems =>
                        {
                            return new
                            {
                                woNumber = woItems.Parent.Parent.Attribute("id").Value,
                                woStyleNumber = woItems.Parent.Element("style").Attribute("id").Value,
                                woBaseName = woItems.Element("base-name").Value,
                                woColorNumbers = woItems.Element("colors").Elements("color").Select(b => b.Element("color-id").Value).ToList(),
                                woColorNames = woItems.Element("colors").Elements("color").Select(b => b.Element("name").Value).ToList(),

                                //Infos = labelBase.Elements().Where(b => b.Name.LocalName.StartsWith("info")).Select(b => b.Value).ToList()
                            };
                        });
            foreach (var data in woData)
            {
                woReturn = data.woNumber;
                styleNoReturn = data.woStyleNumber;
                baseNameReturn = data.woBaseName;
                colorNumbersReturn = data.woColorNumbers;
                colorNamesReturn = data.woColorNames;
            }
            for (int i = 0; i < colorNumbersReturn.Count(); i++)
            {
                colorsCombinedReturn.Add(colorNumbersReturn[i] + " " + colorNamesReturn[i]);
            }
            return (woReturn, styleNoReturn, baseNameReturn, colorNumbersReturn, colorsCombinedReturn);

        }
        public string ShawSalsaParseBoard(string wo)
        {
            string file = Settings.Default.tempDir + "Shaw\\WOs\\" + wo + ".xml";
            //string file = "http://salsaprd.shawinc.com/SALSAWeb/ServiceRequest?service=getSpec&id=" + wo + "&nostatusupdate=true";

            XElement xml = XElement.Load(file);
            string boardType = xml.Element("board-info").Element("board-sku").Element("board-type").Value;
            
            return boardType;

        }
    }
}
