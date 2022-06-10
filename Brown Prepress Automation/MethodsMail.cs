using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.IO;
using Brown_Prepress_Automation.Properties;
using System.Drawing;

namespace Brown_Prepress_Automation
{
    class MethodsMail
    {
        string attachmentsFolder = "\\\\192.168.19.5\\attachments\\";

        MethodsCommon methodsCommon = new MethodsCommon();

        public void SendMailShawTeam(string messageText, bool isError)
        {
            SmtpClient smtpClient = new SmtpClient();
            NetworkCredential basicCredential = new NetworkCredential(Settings.Default.username, Settings.Default.password, Settings.Default.exchangeServer);
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress(Settings.Default.fromEmail);
            // setup up the host, increase the timeout to 5 minutes
            smtpClient.Host = Settings.Default.exchangeServer;
            //smtpClient.Port = 587;
            smtpClient.UseDefaultCredentials = false;
            //smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.Credentials = basicCredential;
            //smtpClient.TargetName = "STARTTLS/smtp.office365.com";
            //smtpClient.EnableSsl = true;
            smtpClient.Timeout = (60 * 5 * 1000);
            message.From = fromAddress;
            if (isError == true)
            {
                message.Subject = "Shaw Label Maker Error Notification - " + DateTime.Now;
                message.IsBodyHtml = true;
                message.Body = "Check your spreadsheet \"" + messageText + "\". It has errors that must be resolved before futher processing. It has been moved to the error folder. Check the txt document for further details.";
            }
            else
            {
                SendMailShawColorCorr(messageText);
                message.Subject = "Shaw Label Maker Success Notification - " + DateTime.Now;
                message.IsBodyHtml = true;
                message.Body = "Your spreadsheet \"" + messageText + "\" has been processed and archived.";
            }
            message.To.Add(Settings.Default.fromEmail);
            if (Settings.Default.debugOn == false)
            {
                foreach (string email in Settings.Default.shawEmailList)
                {
                    message.To.Add(email);                    
                }
            }
            //if (attachments != null)
            //{
            //    foreach (string attachment in attachments)
            //    {
            //        message.Attachments.Add(new Attachment(attachment));
            //    }
            //}
            smtpClient.Send(message);
        }

        public void SendMailShawTeamNew(string messageText, List<string> errors, bool isError)
        {
            SmtpClient smtpClient = new SmtpClient();
            NetworkCredential basicCredential = new NetworkCredential(Settings.Default.username, Settings.Default.password, Settings.Default.exchangeServer);
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress(Settings.Default.fromEmail);
            // setup up the host, increase the timeout to 5 minutes
            smtpClient.Host = Settings.Default.exchangeServer;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = basicCredential;
            smtpClient.Timeout = (60 * 5 * 1000);
            message.From = fromAddress;
            //FileInfo attachment = new FileInfo(Settings.Default.shawHotfolder + "\\" + messageText);
            //attachment.CopyTo(Settings.Default.tempDir + "//" + messageText);
            if (isError == true)
            {
                message.Subject = "Shaw Label Maker Error Notification - " + DateTime.Now;
                message.IsBodyHtml = true;
                message.Body = "Check your spreadsheet \"" + messageText + "\". It has errors that must be resolved before futher processing. It has been moved to the error folder. Check the txt document for further details.<br /><br />";
                foreach (string e in errors)
                {
                    message.Body += "  " + e + "  <br />";
                }
                FormMain.Globals.shawCheckList.Clear();
            }
            else
            {
                SendMailShawColorCorr(messageText);
                message.Subject = "Shaw Label Maker Success Notification - " + DateTime.Now;
                message.IsBodyHtml = true;
                message.Body = "Your spreadsheet \"" + messageText + "\" has been processed and archived.";
                //message.Attachments.Add(new Attachment(Settings.Default.tempDir + "//" + messageText));
            }
            message.To.Add(Settings.Default.fromEmail);
            if (Settings.Default.debugOn == false)
            {
                foreach (string email in Settings.Default.shawEmailList)
                {
                    message.To.Add(email);
                }
            }
            //if (attachments != null)
            //{
            //    foreach (string attachment in attachments)
            //    {
            //        message.Attachments.Add(new Attachment(attachment));
            //    }
            //}
            smtpClient.Send(message);
        }

        public void SendMailShawColorCorr(string messageText)
        {
            SmtpClient smtpClient = new SmtpClient();
            NetworkCredential basicCredential = new NetworkCredential(Settings.Default.username, Settings.Default.password, Settings.Default.exchangeServer);
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress(Settings.Default.fromEmail);
            // setup up the host, increase the timeout to 5 minutes
            smtpClient.Host = Settings.Default.exchangeServer;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = basicCredential;
            smtpClient.Timeout = (60 * 5 * 1000);
            message.From = fromAddress;
            //FileInfo attachment = new FileInfo(Settings.Default.shawHotfolder + "\\" + messageText);
            //attachment.CopyTo(Settings.Default.tempDir + "//" + messageText);

            message.Subject = "Shaw Label Maker " + messageText +" Color Corrections - " + DateTime.Now;
            message.IsBodyHtml = true;
            foreach (string s in FormMain.Globals.shawCheckList)
            {
                message.Body += "<a href=\"https://shawfloors.widencollective.com/api/rest/asset/search/" + s + "?options=downloadUrl&metadata=roomsceneCX52colorcorrect\">";
                message.Body += s;
                message.Body += "</a><br>";
            }
            //message.Attachments.Add(new Attachment(Settings.Default.tempDir + "//" + messageText));

            message.To.Add(Settings.Default.fromEmail);

            //if (attachments != null)
            //{
            //    foreach (string attachment in attachments)
            //    {
            //        message.Attachments.Add(new Attachment(attachment));
            //    }
            //}
            smtpClient.Send(message);
            FormMain.Globals.shawCheckList.Clear();
        }

        public void SendMailShawTray(string messageText)
        {
            SmtpClient smtpClient = new SmtpClient();
            NetworkCredential basicCredential = new NetworkCredential(Settings.Default.username, Settings.Default.password, Settings.Default.exchangeServer);
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress(Settings.Default.fromEmail);
            // setup up the host, increase the timeout to 5 minutes
            smtpClient.Host = Settings.Default.exchangeServer;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = basicCredential;
            smtpClient.Timeout = (60 * 5 * 1000);
            message.From = fromAddress;
            //FileInfo attachment = new FileInfo(Settings.Default.shawHotfolder + "\\" + messageText);
            //attachment.CopyTo(Settings.Default.tempDir + "//" + messageText);

            message.Subject = "Shaw Label Maker - TRAY NEEDED - " + DateTime.Now;
            message.IsBodyHtml = true;
            message.Body += messageText;

            //message.Attachments.Add(new Attachment(Settings.Default.tempDir + "//" + messageText));

            message.To.Add(Settings.Default.fromEmail);

            //if (attachments != null)
            //{
            //    foreach (string attachment in attachments)
            //    {
            //        message.Attachments.Add(new Attachment(attachment));
            //    }
            //}
            smtpClient.Send(message);
            FormMain.Globals.shawCheckList.Clear();
        }

        public void SendMailTuftexTeam(string messageText, bool isError)
        {
            SmtpClient smtpClient = new SmtpClient();
            NetworkCredential basicCredential = new NetworkCredential(Settings.Default.username, Settings.Default.password, Settings.Default.exchangeServer);
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress(Settings.Default.fromEmail);
            // setup up the host, increase the timeout to 5 minutes
            smtpClient.Host = Settings.Default.exchangeServer;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = basicCredential;
            smtpClient.Timeout = (60 * 5 * 1000);
            message.From = fromAddress;
            if (isError == true)
            {
                message.Subject = "Shaw Label Maker Error Notification - " + DateTime.Now;
                message.IsBodyHtml = true;
                message.Body = "Check your spreadsheet \"" + messageText + "\". It has errors that must be resolved before futher processing. It has been moved to the error folder. Check the txt document for further details.";
            }
            else
            {
                message.Subject = "Shaw Label Maker Success Notification - " + DateTime.Now;
                message.IsBodyHtml = true;
                message.Body = "Your spreadsheet \"" + messageText + "\" has been processed and archived.";
            }
            message.To.Add(Settings.Default.fromEmail);
            if (Settings.Default.debugOn == false)
            {
                foreach (string email in Settings.Default.tuftexEmailList)
                {
                    message.To.Add(email);
                }
            }
            //if (attachments != null)
            //{
            //    foreach (string attachment in attachments)
            //    {
            //        message.Attachments.Add(new Attachment(attachment));
            //    }
            //}
            smtpClient.Send(message);
        }

        public void SendMailArmstrongTeam(string messageText, List<string> errors, bool isError)
        {
            SmtpClient smtpClient = new SmtpClient();
            NetworkCredential basicCredential = new NetworkCredential(Settings.Default.username, Settings.Default.password, Settings.Default.exchangeServer);
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress(Settings.Default.fromEmail);
            // setup up the host, increase the timeout to 5 minutes
            smtpClient.Host = Settings.Default.exchangeServer;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = basicCredential;
            smtpClient.Timeout = (60 * 5 * 1000);
            message.From = fromAddress;
            if (isError == true)
            {
                message.Subject = "Armstrong Hotfolder Error Notification - " + DateTime.Now;
                message.IsBodyHtml = true;
                message.Body = "Check your spreadsheet \"" + messageText + "\". It has errors that must be resolved before futher processing. It has been moved to the error folder. Check the txt document for further details.<br /><br />";
                if (errors.Count != 0)
                {
                    foreach (string e in errors)
                    {
                        message.Body += "  " + e + "  <br />";
                    }
                }
                /*message.Subject = "Armstrong Hotfolder Error Notification - " + DateTime.Now;
                message.IsBodyHtml = true;
                message.Body = "Check your spreadsheet \"" + messageText + "\". It has errors that must be resolved before futher processing. It has been moved to the error folder. Check the txt document for further details.";
                */
            }
            else
            {
                message.Subject = "Armstrong Hotfolder Success Notification - " + DateTime.Now;
                message.IsBodyHtml = true;
                message.Body = "Your spreadsheet \"" + messageText + "\" has been processed and archived.";
            }
            message.To.Add(Settings.Default.fromEmail);
            if (Settings.Default.debugOn == false)
            {
                foreach (string email in Settings.Default.armstrongEmailList)
                {
                    message.To.Add(email);
                }
            }
            //if (attachments != null)
            //{
            //    foreach (string attachment in attachments)
            //    {
            //        message.Attachments.Add(new Attachment(attachment));
            //    }
            //}
            smtpClient.Send(message);
        }

        public void SendMailArtDept1(FormMain mainForm, string fileName, List<string> partNumber, List<string> remoteName, List<string> remotePgNum, List<string> csr, List<string> pdfProof, List<string> constructionProof, List<string> specs, List<string> additionalComments, string hotfolder)
        {
            SmtpClient smtpClient = new SmtpClient();
            NetworkCredential basicCredential = new NetworkCredential(Settings.Default.username, Settings.Default.password, Settings.Default.exchangeServer);
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress(Settings.Default.fromEmail);
            // setup up the host, increase the timeout to 5 minutes
            smtpClient.Host = Settings.Default.exchangeServer;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = basicCredential;
            smtpClient.Timeout = (60 * 5 * 1000);
            message.From = fromAddress;
            message.To.Add("matt.howard@brownind.com");
            message.To.Add("emily.lovecraft@brownind.com");
            message.Subject = fileName + " " + DateTime.Now;
            message.IsBodyHtml = true;
            List<string> firstChars = new List<string>();
            string prevSpecs = "";
            string prevComments = "";
            for (int i = 0; i < partNumber.Count; i++)
            {
                message.Body += partNumber[i] + "<br><br>";
                if (remoteName[i].Trim() != "")
                {
                    message.Body += "Remote Job Name: " + remoteName[i] + "<br>";
                }

                if (remotePgNum[i].Trim() != "")
                {
                    message.Body += "Remote Page Number: " + remotePgNum[i] + "<br>";
                }
                
                if (pdfProof[i].ToLower() == "x")
                {
                    message.Body += "PDF Proof requested by " + csr[i] + "<br>";
                }

                if (constructionProof[i].ToLower() == "x")
                {
                    message.Body += "Construction Proof requested by " + csr[i] + "<br>";
                }

                firstChars = csr[i].Split(' ').ToList();

                message.Body += "<br>";

                if (specs[i] == "")
                {
                }
                else if (prevSpecs != specs[i])
                {
                    message.Body += "SPECIFICATIONS:<br>";

                    message.Body += specs[i] + "<br>";

                    message.Body += "<br>";
                    prevSpecs = specs[i];
                }
                else
                {
                    message.Body += "Same Specifications as Above<br><br>";
                }

                if (additionalComments[i] == "")
                {
                }
                else if (prevComments != additionalComments[i])
                {
                    message.Body += "ADDITIONAL COMMENTS:<br>";

                    message.Body += additionalComments[i] + "<br>";

                    message.Body += "<br>";
                    prevComments = additionalComments[i];
                }
                else
                {
                    message.Body += "Same Comments as Above<br><br>";
                }

                message.Body += "<hr><br>";
                
            }
            FileInfo xls = new FileInfo(hotfolder + "\\" + fileName);
            if (methodsCommon.IsFileLocked(xls))
            {
                mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + xls.Name + " is locked. Sleeping... \r\n", Color.Red, FontStyle.Regular); }));
                System.Threading.Thread.Sleep(5000);
            }
            else
            {                
                xls.CopyTo(Settings.Default.tempDir + "\\" + fileName, true);
                message.Attachments.Add(new Attachment(Settings.Default.tempDir + "\\" + fileName));
            }

            DirectoryInfo dinfoZip = new DirectoryInfo(hotfolder);
            FileInfo[] zipFiles = dinfoZip.GetFiles(Path.GetFileNameWithoutExtension(fileName) + ".zip.a*").ToArray();
            while (zipFiles.Count() != 0)
            {
                mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + " | Still " + zipFiles.Count() + " file(s). Sleeping for 5 seconds... \r\n", Color.Red, FontStyle.Regular); }));
                System.Threading.Thread.Sleep(5000);
                zipFiles = dinfoZip.GetFiles(Path.GetFileNameWithoutExtension(fileName) + ".zip.a*").ToArray();
            }

            if (File.Exists(hotfolder + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".zip"))
            {
                FileInfo zip = new FileInfo(hotfolder + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".zip");                
                if (zip.Length > 15000000)
                {
                    zip.CopyTo(attachmentsFolder + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".zip", true);
                    message.Body += "ATTACHMENTS:<br>";
                    message.Body += "<a href=\"http://192.168.19.5/prepress/uploads/attachments/" + Path.GetFileNameWithoutExtension(fileName) + ".zip" + "\">" + Path.GetFileNameWithoutExtension(fileName) + ".zip" + "</a><br>";
                    message.Body += "<br>";
                }
                else
                {
                    zip.CopyTo(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".zip", true);
                    FileInfo attachmentZip = new FileInfo(hotfolder + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".zip");
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
                    message.Attachments.Add(new Attachment(attachmentZip.FullName));
                }
            }
            smtpClient.Send(message);
        }

        public void SendMailPrepressLogError(string messageText, List<string> errors, string csr)
        {
            List<string> csrEmail = new List<string>();
            SmtpClient smtpClient = new SmtpClient();
            NetworkCredential basicCredential = new NetworkCredential(Settings.Default.username, Settings.Default.password, Settings.Default.exchangeServer);
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress(Settings.Default.fromEmail);
            // setup up the host, increase the timeout to 5 minutes
            smtpClient.Host = Settings.Default.exchangeServer;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = basicCredential;
            smtpClient.Timeout = (60 * 5 * 1000);
            message.From = fromAddress;

            message.Subject = "Prepress Hotfolder Error Notification - " + DateTime.Now;
            message.IsBodyHtml = true;
            message.Body = "Check your spreadsheet \"" + messageText + "\". It has errors that must be resolved before futher processing. <br />";
            if (errors.Count != 0)
            {
                foreach (string e in errors)
                {
                    message.Body += " -" + e + "  <br />";
                }
            }

            csrEmail = csr.Split(' ').ToList();

            message.To.Add(Settings.Default.fromEmail);
            if (Settings.Default.debugOn == false)
            {
                message.To.Add(csrEmail[0] + "." + csrEmail[1] + "@brownind.com");
            }

            smtpClient.Send(message);
        }

        public void SendMailCreateHotFolder(string messageText)
        {
            SmtpClient smtpClient = new SmtpClient();
            NetworkCredential basicCredential = new NetworkCredential(Settings.Default.username, Settings.Default.password, Settings.Default.exchangeServer);
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress(Settings.Default.fromEmail);
            // setup up the host, increase the timeout to 5 minutes
            smtpClient.Host = Settings.Default.exchangeServer;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = basicCredential;
            smtpClient.Timeout = (60 * 5 * 1000);
            message.From = fromAddress;
            message.Subject = "Shaw Label Maker Hotfolder Creation Notification - " + DateTime.Now;
            message.IsBodyHtml = true;
            message.Body = "Create a Hotfolder for " + messageText;
            message.To.Add(Settings.Default.fromEmail);
            //if (attachments != null)
            //{
            //    foreach (string attachment in attachments)
            //    {
            //        message.Attachments.Add(new Attachment(attachment));
            //    }
            //}
            smtpClient.Send(message);
        }
        
        public void SendMailTicket(string file, string customer)
        {
            SmtpClient smtpClient = new SmtpClient();
            NetworkCredential basicCredential = new NetworkCredential(Settings.Default.username, Settings.Default.password, Settings.Default.exchangeServer);
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress(Settings.Default.fromEmail);
            // setup up the host, increase the timeout to 5 minutes
            smtpClient.Host = Settings.Default.exchangeServer;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = basicCredential;
            smtpClient.Timeout = (60 * 5 * 1000);
            message.From = fromAddress;
            message.Subject = Path.GetFileNameWithoutExtension(file) + " Ticket PDF";
            message.IsBodyHtml = true;
            message.Body = Path.GetFileNameWithoutExtension(file);
            message.To.Add(Settings.Default.fromEmail);
            if (Settings.Default.debugOn == false)
            {
                if (customer.ToLower() == "shaw")
                {
                    foreach (string email in Settings.Default.shawEmailList)
                    {
                        message.To.Add(email);
                    }
                }
                if (customer.ToLower() == "armstrong")
                {
                    foreach (string email in Settings.Default.armstrongEmailList)
                    {
                        message.To.Add(email);
                    }
                }
            }
            message.Attachments.Add(new Attachment(file));

            //if (attachments != null)
            //{
            //    foreach (string attachment in attachments)
            //    {
            //        message.Attachments.Add(new Attachment(attachment));
            //    }
            //}
            smtpClient.Send(message);
        }
    }
}
