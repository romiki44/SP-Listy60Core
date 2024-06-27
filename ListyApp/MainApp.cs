using ListyApp.Models;
using ListyApp.Models.Exports;
using ListyApp.Repositories;
using ListyApp.Services;
using McMaster.Extensions.CommandLineUtils;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Serilog;
using Serilog.Core;
using Serilog.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListyApp
{
    public class MainApp
    {
        private readonly ILogger<MainApp> logger;
        private readonly ISdsConfigRepo sdsConfigRepo;
        private readonly IMsWordTools msWordTools;
        private readonly PdfOptions pdfOptions;
        
        [System.ComponentModel.DataAnnotations.Required]
        [Option("-m|--mode", CommandOptionType.SingleValue, Description = "Input mode")]
        [AllowedValues("sql", "app")]
        public string Mode { get; set; }

        [Option("-s|--sprac", CommandOptionType.SingleValue, Description = "Spracovanie")]
        [System.ComponentModel.DataAnnotations.Required]
        public string Sprac { get; set; }

        public MainApp(ILogger<MainApp> logger, IOptions<PdfOptions> exportOptions, ISdsConfigRepo sdsConfigRepo, IMsWordTools msWordTools)
        {
            this.logger = logger;
            this.sdsConfigRepo = sdsConfigRepo;
            this.msWordTools = msWordTools;
            this.pdfOptions = exportOptions.Value;            
        }

        private async Task OnExecuteAsync()
        {
            await StartProcess();                                 
        }

        public async Task StartProcess()
        {
            logger.LogInformation("Program ListyApp started...");
            await Task.Delay(2000);

            //string docxPathFile = @"\\data\data\Public\Odd412\_Roman\Listy60Test\Template\Infolisty60_eschranky.dotx";
            //string pdfExportFolder = @"\\data\data\Public\Odd412\_Roman\Listy60Test\Spracovanie\202404\eschranky\export";
            //string pdfBaseName = "202404_";
            //string odcPathFile = @"\\data\data\Public\Odd412\_Roman\Listy60Test\Odc\uv_ExportDocEschranky.odc";
            //int pagePerPdf = 5;

            string docxPathFile = Path.Combine(pdfOptions.TemplatePath, pdfOptions.EschrankyOptions.DocxFileName);
            string pdfExportFolder = Path.Combine(pdfOptions.SpracovaniePath, Sprac, pdfOptions.EschrankyOptions.WorkFolder, pdfOptions.EschrankyOptions.ExportSubFolder);
            string pdfBaseName = $"{Sprac}_";
            string odcPathFile = Path.Combine(pdfOptions.OdcPath, pdfOptions.EschrankyOptions.OdcFileName);
            int pagesPerPdf = pdfOptions.EschrankyOptions.PagesPerPdf;

            logger.LogInformation($"Spracovanie: {Sprac}");
            logger.LogInformation($"DoxcTemplatePath: {docxPathFile}");
            logger.LogInformation($"PdfExportFolder: {pdfExportFolder}");
            logger.LogInformation($"OdcPath: {odcPathFile}");
            logger.LogInformation($"PagesPerPdf: {pagesPerPdf}");
            //return;

            bool deleteExistFiles = true;
            if (deleteExistFiles)
            {
                logger.LogInformation("Deleting files...");
                //FileTools.DeleteAllPdfFiles(pdfExportFolder);
            }

            logger.LogInformation("Starting process for converting word to pdf...");
            msWordTools.ConvertMergedWordToSinglePdf(docxPathFile, pdfExportFolder, pdfBaseName, odcPathFile, pagesPerPdf);
            logger.LogInformation("Program ListyApp finished...");
        }
    }
}
