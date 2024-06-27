using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListyApp.Models.Exports
{
    public class PdfOptions
    {
        public string SpracovaniePath { get; set; }        
        public string TemplatePath { get; set; }
        public string OdcPath { get; set; }
        public EschrankyOptions EschrankyOptions { get; set; }
        public TlacSkOptions TlacSkOptions { get; set; }
        public TlacEuOptions TlacEuOptions { get; set; }

        public override string ToString()
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append($"PdfOptions:\nSpracovaniePath: {SpracovaniePath}\nTemplatePath: {TemplatePath}\nOdcPath: {OdcPath}\n");
            stringBuilder.Append($"EschrankyOptions:\n\tWorkFolder:{EschrankyOptions.WorkFolder}\n\tDocxFileName:{EschrankyOptions.DocxFileName}\n\tOdcFileName:{EschrankyOptions.OdcFileName}\n\tExportSubFolder:{EschrankyOptions.ExportSubFolder}\n");
            stringBuilder.Append($"TlacSkOptions:\n\tWorkFolder:{TlacSkOptions.WorkFolder}\n\tDocxFileName:{TlacSkOptions.DocxFileName}\n\tOdcFileName:{TlacSkOptions.OdcFileName}\n");
            stringBuilder.Append($"TlacEuOptions:\n\tWorkFolder:{TlacEuOptions.WorkFolder}\n\tDocxFileName:{TlacEuOptions.DocxFileName}\n\tOdcFileName:{TlacEuOptions.OdcFileName}\n");
            return stringBuilder.ToString();
        }
    }
}
