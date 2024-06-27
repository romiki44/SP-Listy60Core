using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Extensions.Logging;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace ListyApp.Services
{
    public class MsWordToolsException : Exception
    {
        public MsWordToolsException(string message, Exception innerException)
            : base(message, innerException)
        { }
    }

    public class MsWordTools: IMsWordTools
    {
        private ILogger<MsWordTools> logger;

        public bool PdfAutoIncrement { get; set; }

        public MsWordTools(ILogger<MsWordTools> logger)
        {
            this.logger = logger;
        }

        //https://stackoverflow.com/questions/607669/how-do-i-convert-word-files-to-pdf-programmatically
        //export wordov do pdf, pouziva sa interny office export
        public void ConvertWordToPdf(string docxPathFile, string pdfPathFile)
        {
            Application? app = null;
            Document? wordDoc = null;
            object missing = Missing.Value;
            string wordFileName = "";

            try
            {
                wordFileName = Path.GetFileName(docxPathFile);
                string pdfFileName = Path.GetFileName(pdfPathFile);
                bool isTemplate = Path.GetExtension(wordFileName).Equals(".dotx", StringComparison.OrdinalIgnoreCase);

                logger.LogInformation($"ConvertWordToPdf start processing single word document '{wordFileName}'.");

                app = new Application();
                app.Visible = true;
                Documents appDocuments = app.Documents;

                if (isTemplate)
                    wordDoc = appDocuments.Add(docxPathFile);
                else
                    wordDoc = appDocuments.Open(docxPathFile, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

                wordDoc.ExportAsFixedFormat(pdfPathFile, WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument);
                logger.LogInformation($"ConvertWordToPdf converted word document to pdf '{pdfFileName}'.");

                wordDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                wordDoc = null;
            }
            catch (Exception ex)
            {
                string errmsg = ex.InnerException != null ? $"{ex.Message}, {ex.InnerException.Message}" : ex.Message;
                string errtitle = $"ConvertWordToPdf failed converting '{wordFileName}' to pdf.";
                logger.LogError($"{errtitle} {errmsg}");
                throw new MsWordToolsException(errtitle, ex);
            }
            finally
            {
                app?.Quit();
                ReleaseObject(app);
            }
        }

        //export zoznamu wordov do pdf, pouziva sa interny office export
        //ak sa nezada priecinok pre export pdf, pouzije sa ten isty priecinok ako pre word...
        //teor by sa dal pouzivat v cyklus pre jeden word, ale chcel som sa vyhnut opakovanej inicializacii a uvolnovani zdrojov WordApp
        public void ConvertWordToPdf(List<string> docxPathFiles, string? pdfExportFolder = null)
        {
            Application? app = null;
            string wordFileName = "";

            try
            {
                app = new Application();
                app.Visible = true;
                Documents appDocuments = app.Documents;

                logger.LogInformation($"ConvertWordToPdf start processing multiple word documents [{docxPathFiles.Count}].");

                foreach (var docxPathFile in docxPathFiles)
                {
                    wordFileName = Path.GetFileName(docxPathFile);
                    string wordFileNameWithoutExt = Path.GetFileNameWithoutExtension(docxPathFile);
                    string wordFolder = Path.GetDirectoryName(docxPathFile)!;
                    bool isTemplate = Path.GetExtension(wordFileName).Equals(".dotx", StringComparison.OrdinalIgnoreCase);

                    if (string.IsNullOrEmpty(pdfExportFolder))
                        pdfExportFolder = wordFolder;

                    string pdfFileName = wordFileNameWithoutExt + ".pdf";
                    string pdfPathFile = Path.Combine(pdfExportFolder, pdfFileName);

                    Document? wordDoc = null;
                    if (isTemplate)
                        wordDoc = appDocuments.Add(docxPathFile);
                    else
                        wordDoc = appDocuments.Open(docxPathFile, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

                    wordDoc.ExportAsFixedFormat(pdfPathFile, WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument);
                    logger.LogInformation($"ConvertWordToPdf converted word document to pdf '{pdfFileName}'.");

                    wordDoc.Close();
                    wordDoc = null;
                }

                app.Quit();
            }
            catch (Exception ex)
            {
                string errmsg = ex.InnerException != null ? $"{ex.Message}, {ex.InnerException.Message}" : ex.Message;
                string errtitle = $"ConvertWordToPdf failed converting '{wordFileName}' to pdf.";
                logger.LogError($"{errtitle} {errmsg}");
                throw new MsWordToolsException(errtitle, ex);
            }
            finally
            {
                app?.Quit();
                ReleaseObject(app);
            }
        }

        //tlaci word do pdf cez pdf print driver
        public void PrintWordToPdf(string docxPathFile, string pdfPathFile, string printDriverName = "Microsoft Print to PDF")
        {
            Application? app = null;
            object missing = Missing.Value;
            object oTrue = true;
            string wordFileName = "";

            try
            {
                wordFileName = Path.GetFileName(docxPathFile);
                string pdfFileName = Path.GetFileName(pdfPathFile);
                bool isTemplate = Path.GetExtension(wordFileName).Equals(".dotx", StringComparison.OrdinalIgnoreCase);

                logger.LogInformation($"PrintWordToPdf start processing single word document '{wordFileName}'.");

                app = new Application();
                app.Visible = true;
                app.ActivePrinter = printDriverName;
                Documents appDocuments = app.Documents;

                Document? wordDoc = null;
                if (isTemplate)
                    wordDoc = appDocuments.Add(docxPathFile);
                else
                    wordDoc = appDocuments.Open(docxPathFile, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

                app.PrintOut(OutputFileName: pdfPathFile, PrintToFile: oTrue, Range: WdPrintOutRange.wdPrintAllDocument, Item: WdPrintOutItem.wdPrintDocumentWithMarkup, PageType: WdPrintOutPages.wdPrintAllPages);
                logger.LogInformation($"PrintWordToPdf is printing word document to pdf '{pdfFileName}'.");

                wordDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
            }
            catch (Exception ex)
            {
                string errmsg = ex.InnerException != null ? $"{ex.Message}, {ex.InnerException.Message}" : ex.Message;
                string errtitle = $"PrintWordToPdf failed print '{wordFileName}' to pdf";
                logger.LogError($"{errtitle} {errmsg}");
                throw new MsWordToolsException(errtitle, ex);
            }
            finally
            {
                app?.Quit();
                ReleaseObject(app);
            }
        }

        //tlaci zoznam wordov do pdf suborv cez pdf print driver
        public void PrintWordToPdf(List<string> docxPathFiles, string? pdfExportFolder = null, string printDriverName = "Microsoft Print to PDF")
        {
            Application? app = null;
            object missing = Missing.Value;
            object oTrue = true;
            string wordFileName = "";

            try
            {
                app = new Application();
                app.Visible = true;
                app.ActivePrinter = printDriverName;
                Documents appDocuments = app.Documents;

                logger.LogInformation($"PrintWordToPdf start processing multiple word documents [{docxPathFiles.Count}].");

                foreach (var docxPathFile in docxPathFiles)
                {
                    wordFileName = Path.GetFileName(docxPathFile);
                    string wordFileNameWithoutExt = Path.GetFileNameWithoutExtension(docxPathFile);
                    bool isTemplate = Path.GetExtension(wordFileName).Equals(".dotx", StringComparison.OrdinalIgnoreCase);
                    string? wordFolder = Path.GetDirectoryName(docxPathFile);

                    if (string.IsNullOrEmpty(pdfExportFolder))
                        pdfExportFolder = wordFolder;

                    string pdfFileName = wordFileNameWithoutExt + ".pdf";
                    string pdfPathFile = Path.Combine(pdfExportFolder, pdfFileName);

                    Document? wordDoc = null;
                    if (isTemplate)
                        wordDoc = appDocuments.Add(docxPathFile);
                    else
                        wordDoc = appDocuments.Open(docxPathFile, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

                    app.PrintOut(OutputFileName: pdfPathFile, PrintToFile: oTrue, Range: WdPrintOutRange.wdPrintAllDocument, Item: WdPrintOutItem.wdPrintDocumentWithMarkup, PageType: WdPrintOutPages.wdPrintAllPages);
                    logger.LogInformation($"PrintWordToPdf is printing word document to pdf '{pdfFileName}'.");

                    wordDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                }
            }
            catch (Exception ex)
            {
                string errmsg = ex.InnerException != null ? $"{ex.Message}, {ex.InnerException.Message}" : ex.Message;
                string errtitle = $"PrintWordToPdf failed print '{wordFileName}' to pdf";
                logger.LogError($"{errtitle} {errmsg}");
                throw new MsWordToolsException(errtitle, ex);
            }
            finally
            {
                app?.Quit();
                ReleaseObject(app);
            }
        }

        public int PrintMergedWordToPdf(string docxPathFile, string pdfPathFile, string odcFileName, int firstRecord, int maxRecordCount, string printDriverName = "Microsoft Print to PDF")
        {
            Application? app = null;
            object missing = Missing.Value;
            object oTrue = true;
            object oFalse = false;
            int activeRecordCount = 0;

            try
            {
                string wordFileName = Path.GetFileName(docxPathFile);
                string pdfFileName = Path.GetFileName(pdfPathFile);
                bool isTemplate = Path.GetExtension(wordFileName).Equals(".dotx", StringComparison.OrdinalIgnoreCase);

                logger.LogInformation($"PrintMergedWordToPdf start processing merged word document '{wordFileName}'.");

                app = new Application();
                app.Visible = true;
                app.ActivePrinter = printDriverName;
                Documents appDocuments = app.Documents;

                Document? wordDoc = null;
                if (isTemplate)
                    wordDoc = appDocuments.Add(docxPathFile);
                else
                    wordDoc = appDocuments.Open(docxPathFile, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

                wordDoc.MailMerge.OpenDataSource(Name: odcFileName, Format: WdOpenFormat.wdOpenFormatAuto, ConfirmConversions: oFalse, ReadOnly: oTrue);

                //kolko zaznamov sa bude tlacit? Prvy bude vzdy ako prvy, posledny bude posledny povoleny, alebo posledny existujuci....
                activeRecordCount = wordDoc.MailMerge.DataSource.RecordCount;
                int firstMergedRecord = firstRecord < 1 ? 1 : firstRecord;  //firstMergedRecord musi byt vacsi alebo 1
                int lastEnabledRecord = firstMergedRecord + maxRecordCount - 1;  //teoreticky posledny dovoleny record    
                int lastMergedRecord = activeRecordCount > lastEnabledRecord ? lastEnabledRecord : activeRecordCount;
                wordDoc.MailMerge.DataSource.FirstRecord = firstMergedRecord;
                wordDoc.MailMerge.DataSource.LastRecord = lastMergedRecord;

                wordDoc.MailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters;
                wordDoc.MailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;
                wordDoc.MailMerge.SuppressBlankLines = true;

                wordDoc.MailMerge.Execute(oFalse);
                logger.LogInformation($"PrintMergedWordToPdf merged [{firstMergedRecord} - {lastMergedRecord}] records in word document of [{activeRecordCount}] active records.");

                var activeMergeDoc = app.ActiveDocument;
                app.PrintOut(OutputFileName: pdfPathFile, PrintToFile: oTrue, Range: WdPrintOutRange.wdPrintAllDocument, Item: WdPrintOutItem.wdPrintDocumentWithMarkup, PageType: WdPrintOutPages.wdPrintAllPages);
                logger.LogInformation($"PrintMergedWordToPdf is printing merged word document to pdf '{pdfFileName}'.");

                wordDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                activeMergeDoc.Close(WdSaveOptions.wdDoNotSaveChanges);

                return activeRecordCount;
            }
            catch (Exception ex)
            {
                string errmsg = ex.InnerException != null ? $"{ex.Message}, {ex.InnerException.Message}" : ex.Message;
                string errtitle = $"PrintMergedWordToPdf failed print merged document '{docxPathFile}' to pdf.";
                logger.LogError($"{errtitle} {errmsg}");
                throw new MsWordToolsException(errtitle, ex);
            }
            finally
            {
                app?.Quit();
                ReleaseObject(app);
            }
        }

        public int ConvertMergedWordToSinglePdf(string docxPathFile, string pdfFolder, string pdfFileBaseName, string odcFileName, int pagePerPdf, string printDriverName = "Microsoft Print to PDF")
        {
            Application? app = null;
            object oMissing = Missing.Value;
            object oTrue = true;
            object oFalse = false;
            object oMsoTrue = MsoTriState.msoTrue;
            object oMsoFalse = MsoTriState.msoFalse;
            object oDocxPathFile = docxPathFile;
            int activeRecordCount = 0;

            try
            {
                string wordFileName = Path.GetFileName(docxPathFile);
                bool isTemplate = Path.GetExtension(wordFileName).Equals(".dotx", StringComparison.OrdinalIgnoreCase);

                logger.LogInformation($"ConvertMergedWordToSinglePdf start processing merged word document '{wordFileName}'.");

                app = new Application();
                app.Visible = true;
                app.ActivePrinter = printDriverName;
                Documents appDocuments = app.Documents;

                Document? wordDoc = null;
                if (isTemplate)
                    wordDoc = appDocuments.Add(docxPathFile);
                else
                    wordDoc = app.Documents.Open(ref oDocxPathFile, ref oMsoTrue, ref oMsoTrue, ref oMsoFalse, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                wordDoc.MailMerge.OpenDataSource(Name: odcFileName, Format: WdOpenFormat.wdOpenFormatAuto, ConfirmConversions: oFalse, ReadOnly: oTrue);
                logger.LogInformation($"MailMerge completed OpenDataSource");

                //kolko zaznamov sa bude tlacit? Prvy bude vzdy ako prvy, posledny bude posledny povoleny, alebo posledny existujuci....
                activeRecordCount = wordDoc.MailMerge.DataSource.RecordCount;
                wordDoc.MailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters;
                wordDoc.MailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;
                wordDoc.MailMerge.SuppressBlankLines = true;

                wordDoc.MailMerge.Execute(oFalse);
                var activeMergeDoc = app.ActiveDocument;
                logger.LogInformation($"Merged [{activeRecordCount}] active records into new document '{activeMergeDoc.Name}'");

                //app.PrintOut(OutputFileName: pdfPathFile, PrintToFile: oTrue, Range: WdPrintOutRange.wdPrintAllDocument, Item: WdPrintOutItem.wdPrintDocumentWithMarkup, PageType: WdPrintOutPages.wdPrintAllPages);

                //robim konverziu ExportAsFixedFormat(), pretoze ide vybrat pocet stran...pri PrintOut() to akosi nefunguje!
                int irec = 1;
                int maxPages = activeRecordCount * pagePerPdf;
                for (int fromPage = 1; fromPage <= maxPages; fromPage = fromPage + pagePerPdf)
                {
                    int toPage = fromPage + pagePerPdf - 1;
                    string pdfExportFileName = Path.GetFileNameWithoutExtension(pdfFileBaseName) + "_" + string.Format("{0:D4}", irec) + ".pdf";
                    string pdfExportPathFile = Path.Combine(pdfFolder, pdfExportFileName);
                    activeMergeDoc.ExportAsFixedFormat(pdfExportPathFile, WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportFromTo, fromPage, toPage);
                    logger.LogInformation($"Converted pages [{fromPage}-{toPage}] to single pdf '{pdfExportFileName}' of {irec}/{activeRecordCount}.");
                    irec++;
                }

                wordDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                activeMergeDoc.Close(WdSaveOptions.wdDoNotSaveChanges);

                return activeRecordCount;
            }
            catch (Exception ex)
            {
                string errmsg = ex.InnerException != null ? $"{ex.Message}, {ex.InnerException.Message}" : ex.Message;
                string errtitle = $"ConvertMergedWordToSinglePdf failed print merged document '{docxPathFile}' to pdf.";
                //logger.LogError($"{errtitle} {errmsg}");
                logger.LogError(ex.ToString());
                throw new MsWordToolsException(errtitle, ex);
            }
            finally
            {
                app?.Quit();
                ReleaseObject(app);
            }
        }

        public void PrintMergedWordWithExcelToPdf(string docxPathFile, string pdfPathFile, string excelPathFile, int firstRecord, int maxRecordCount, string sheetData = "Data", string printDriverName = "Microsoft Print to PDF")
        {
            Application? app = null;
            Document? wordDoc = null;
            object missing = Missing.Value;
            object oTrue = true;
            object oFalse = false;
            int activeRecordCount = 0;

            try
            {
                app = new Application();
                app.Visible = true;
                app.ActivePrinter = printDriverName;

                string wordFileName = Path.GetFileName(docxPathFile);
                string pdfFileName = Path.GetFileName(pdfPathFile);
                string excelFileName = Path.GetFileName(excelPathFile);
                bool isTemplate = Path.GetExtension(wordFileName).Equals(".dotx", StringComparison.OrdinalIgnoreCase);
                logger.LogInformation($"PrintMergedWordWithExcelToPdf start processing word document '{wordFileName}' merged with excelfile '{excelFileName}'.");

                Documents appDocuments = app.Documents;

                if (isTemplate)
                    wordDoc = appDocuments.Add(docxPathFile);
                else
                    wordDoc = appDocuments.Open(docxPathFile, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

                string odsName = excelPathFile;
                string odsConn = @"Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Initial Catalog=Listy60M1;Mode=Read;Extended Properties=""HDR=YES;IMEX=1;"";Data Source=" + excelPathFile;
                string sqlStatement = $"SELECT * FROM `{sheetData}$`";

                wordDoc.MailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters;
                wordDoc.MailMerge.OpenDataSource(Name: odsName, Connection: odsConn, SQLStatement: sqlStatement, LinkToSource: true, Format: WdOpenFormat.wdOpenFormatAuto, SubType: WdMergeSubType.wdMergeSubTypeOther);

                wordDoc.MailMerge.DataSource.ActiveRecord = WdMailMergeActiveRecord.wdFirstRecord;
                wordDoc.MailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;
                wordDoc.MailMerge.SuppressBlankLines = true;

                //kolko zaznamov sa bude tlacit? Prvy bude vzdy ako prvy, posledny bude posledny povoleny, alebo posledny existujuci....
                activeRecordCount = wordDoc.MailMerge.DataSource.RecordCount;
                int firstMergedRecord = firstRecord < 1 ? 1 : firstRecord;  //firstMergedRecord musi byt vacsi alebo 1
                int lastEnabledRecord = firstMergedRecord + maxRecordCount - 1;  //teoreticky posledny dovoleny record    
                int lastMergedRecord = activeRecordCount > lastEnabledRecord ? lastEnabledRecord : activeRecordCount;
                wordDoc.MailMerge.DataSource.FirstRecord = firstMergedRecord;
                wordDoc.MailMerge.DataSource.LastRecord = lastMergedRecord;

                wordDoc.MailMerge.Execute(oFalse);
                logger.LogInformation($"PrintMergedWordWithExcelToPdf merged [{firstMergedRecord} - {lastMergedRecord}] records in word document of [{activeRecordCount}] active records.");

                var activeMergeDoc = app.ActiveDocument;

                //pri tlaci do pdf range of pages, cize tlac len vybratych stranok proste nefunguje...ani vo worde!!!...vzdy tlaci cely dokument!!!                                
                app.PrintOut(OutputFileName: pdfPathFile, PrintToFile: oTrue, Range: WdPrintOutRange.wdPrintAllDocument, Item: WdPrintOutItem.wdPrintDocumentWithMarkup, PageType: WdPrintOutPages.wdPrintAllPages);
                logger.LogInformation($"PrintMergedWordWithExcelToPdf print merged word document to pdf '{pdfFileName}'.");

                //zavriem zdrojovy dokument + zmergovany dokument
                wordDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                activeMergeDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
            }
            catch (Exception ex)
            {
                string errmsg = ex.InnerException != null ? $"{ex.Message}, {ex.InnerException.Message}" : ex.Message;
                string errtitle = $"PrintMergedWordWithExcelToPdf failed print '{docxPathFile}' to pdf.";
                logger.LogError($"{errtitle} {errmsg}");
                throw new MsWordToolsException(errtitle, ex);
            }
            finally
            {
                app?.Quit();
                ReleaseObject(app);
            }

        }

        private void ReleaseObject(object? xlObject)
        {
            try
            {
                Marshal.ReleaseComObject(xlObject);
                xlObject = null;
                logger.LogInformation("MSWrdTools finished all processing with releasing msword-application objects");
            }
            catch (Exception ex)
            {
                logger.LogError($"MSWordTools ReleaseObject Error. {ex.Message}");
                xlObject = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
