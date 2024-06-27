using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListyApp.Models.Exports
{
    public class EschrankyOptions
    {
        public string WorkFolder { get; set; }
        public string DocxFileName { get; set; }
        public string OdcFileName { get; set; }
        public string ExportSubFolder { get; set; }
        public int PagesPerPdf { get; set; }
    }
}
