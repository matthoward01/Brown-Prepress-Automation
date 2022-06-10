using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;

namespace Brown_Prepress_Automation
{
    class Zip
    {
       public void zipDownload()
       {
        using (var client = new WebClient())
        {
            client.DownloadFile("http://something",  @"D:\Downloads\1.zip");
        }
       }
    }
}
