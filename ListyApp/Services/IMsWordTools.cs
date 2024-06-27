using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace ListyApp.Services
{
    public interface IMsWordTools
    {
        void ConvertWordToPdf(string docxPathFile, string pdfPathFile);
        void ConvertWordToPdf(List<string> docxPathFiles, string? pdfExportFolder = null);
        void PrintWordToPdf(string docxPathFile, string pdfPathFile, string printDriverName = "Microsoft Print to PDF");
        void PrintWordToPdf(List<string> docxPathFiles, string? pdfExportFolder = null, string printDriverName = "Microsoft Print to PDF");
        int PrintMergedWordToPdf(string docxPathFile, string pdfPathFile, string odcFileName, int firstRecord, int maxRecordCount, string printDriverName = "Microsoft Print to PDF");
        int ConvertMergedWordToSinglePdf(string docxPathFile, string pdfFolder, string pdfFileBaseName, string odcFileName, int pagePerPdf, string printDriverName = "Microsoft Print to PDF");
        void PrintMergedWordWithExcelToPdf(string docxPathFile, string pdfPathFile, string excelPathFile, int firstRecord, int maxRecordCount, string sheetData = "Data", string printDriverName = "Microsoft Print to PDF");
    }
}
