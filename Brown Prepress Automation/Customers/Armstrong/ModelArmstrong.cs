using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Brown_Prepress_Automation
{
    public class ModelArmstrong
    {
        public class ArmstrongSheet
        {
            public List<string> PartNumber { get; set; } = new List<string>();
            public List<string> FileName { get; set; } = new List<string>();
            public List<string> Size { get; set; } = new List<string>();
            public List<string> Quantity { get; set; } = new List<string>();
            public List<string> Stock { get; set; } = new List<string>();
            public List<string> SalesOrder { get; set; } = new List<string>();
        }

        public class ArmstrongParse
        {
            public List<string> BlList { get; set; } = new List<string>();
            public List<int> BlLines { get; set; } = new List<int>();
            public List<string> FlList { get; set; } = new List<string>();
            public List<int> FlLines { get; set; } = new List<int>();
            public List<int> LabelQty { get; set; } = new List<int>();
        }

        public class ArmstrongDB
        {
            public List<string> FileName { get; set; } = new List<string>();
            public List<string> FileNameAlt { get; set; } = new List<string>();
            public List<string> PartNumber { get; set; } = new List<string>();
            public List<string> Stock { get; set; } = new List<string>();
        }
    }
}
