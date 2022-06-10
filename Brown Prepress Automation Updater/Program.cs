using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;

namespace Brown_Prepress_Automation_Updater
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo setupFile = new FileInfo("\\\\192.168.19.5\\Programs\\Brown Prepress Automation.msi");
            string appDir = System.IO.Directory.GetCurrentDirectory() + "\\";
            FileInfo appFile = new FileInfo("\\Temp\\Brown Prepress Automation.msi");
            if (appFile.Exists)
            {
                if (setupFile.LastWriteTime > appFile.LastWriteTime)
                {
                    Console.WriteLine("New Version Detected...");

                    Console.WriteLine("Uninstalling Old Version...");
                    setupFile.CopyTo("\\Temp\\Brown Prepress Automation.msi", true);
                    Process i = new Process();
                    Process u = new Process();
                    Process s = new Process();
                    u.StartInfo.FileName = "msiexec";
                    u.StartInfo.Arguments = "/x {8B28154F-5927-43D4-A873-2E6F52A2DC1E} /quiet";
                    u.Start();
                    u.WaitForExit();

                    Console.WriteLine("Installing New Version...");
                    i.StartInfo.FileName = "\\Temp\\Brown Prepress Automation.msi";
                    i.StartInfo.Arguments = "/quiet";
                    i.Start();
                    i.WaitForExit();

                    Console.WriteLine("Closing Updater and starting Brown Prepress Automation...");
                    s.StartInfo.FileName = appDir + "Brown Prepress Automation.exe";
                    s.Start();

                }
            }
        }
    }
}
