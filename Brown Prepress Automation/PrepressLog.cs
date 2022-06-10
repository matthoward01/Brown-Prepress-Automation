using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelLibrary.SpreadSheet;
using Brown_Prepress_Automation.Properties;
using System.IO;

namespace Brown_Prepress_Automation
{
    class PrepressLog
    {
        MethodsMySQL methodsMySQL = new MethodsMySQL();
        MethodsCommon methodsCommon = new MethodsCommon();
        MethodsMail methodsMail = new MethodsMail();

        string prepressHotfolder = Settings.Default.prepressHotfolder;

        public void Log(FormMain mainForm, string passedFile)
        {     
            List<string> partNumberList = new List<string>();
            List<string> remoteNameList = new List<string>();
            List<string> remotePgNumList = new List<string>();
            List<string> csrList = new List<string>();
            List<string> pdfProofList = new List<string>();
            List<string> constructionProofList = new List<string>();
            List<string> specsList = new List<string>();
            List<string> additionalCommentsList = new List<string>();
            List<string> errorList = new List<string>();
            bool isError = false;
            string workingFile = prepressHotfolder + "\\" + passedFile;
            Workbook book = Workbook.Load(workingFile);
            Worksheet sheet = book.Worksheets[0];
            int validCellsCheck = methodsCommon.countValidCells(workingFile, 1, 0, 0);
            for (int i = 1; i < validCellsCheck; i++)
            {
                string partNumber = sheet.Cells[i, 0].StringValue.Trim();
                string remoteName = sheet.Cells[i, 1].StringValue.Trim();
                string remotePgNum = sheet.Cells[i, 2].StringValue.Trim();
                string csr = methodsMySQL.SelectCSR(passedFile);
                string pdfProof = sheet.Cells[i, 3].StringValue.Trim();
                string constructionProof = sheet.Cells[i, 4].StringValue.Trim();
                string specs = sheet.Cells[i, 5].StringValue.Trim();
                string additionalComments = sheet.Cells[i, 6].StringValue.Trim();                

                partNumberList.Add(partNumber);
                remoteNameList.Add(remoteName);
                remotePgNumList.Add(remotePgNum);
                csrList.Add(csr);
                pdfProofList.Add(pdfProof);
                constructionProofList.Add(constructionProof);
                specsList.Add(specs);
                additionalCommentsList.Add(additionalComments);
                if (specs.Trim() == "")
                {
                    isError = true;
                    errorList.Add(partNumber + " does not have any Specifications.");
                }
                else if (remoteName.Trim() == "")
                {
                    isError = true;
                    errorList.Add(partNumber + " does not have a Remote Name.");
                }
                else if (partNumber.Trim() == "")
                {
                    isError = true;
                    errorList.Add("Line " + (i+1) + " is missing a Part Number or Proof Name.");
                }
                else if ((pdfProof.Trim() == "") && (constructionProof.Trim() == ""))
                {
                    isError = true;
                    errorList.Add(partNumber + " does not have Pdf Proof or Construction Proof selected.");
                }
                else
                {
                    methodsMySQL.InsertPrepressLog(partNumber, csr, pdfProof, constructionProof, specs, additionalComments);
                }
            }
            if (isError)
            {
                methodsMail.SendMailPrepressLogError(passedFile, errorList, (csrList[0]));
            }
            else
            {
                methodsMail.SendMailArtDept1(mainForm, passedFile, partNumberList, remoteNameList, remotePgNumList, csrList, pdfProofList, constructionProofList, specsList, additionalCommentsList, prepressHotfolder);
            }
            if (File.Exists(prepressHotfolder + "\\" + passedFile))
            {
                System.IO.File.Delete(prepressHotfolder + "\\" + passedFile);
            }
            if (File.Exists(prepressHotfolder + "\\" + Path.GetFileNameWithoutExtension(passedFile) + ".zip"))
            {
                FileInfo attachmentZip = new FileInfo(prepressHotfolder + "\\" + Path.GetFileNameWithoutExtension(passedFile) + ".zip");
                bool locked = true;
                while (locked == true)
                {
                    if (methodsCommon.IsFileLocked(attachmentZip))
                    {
                        //System.Threading.Thread.Sleep(5000);
                    }
                    else
                    {
                        locked = false;
                    }
                }
                attachmentZip.Delete();
            }
        }
    }
}
