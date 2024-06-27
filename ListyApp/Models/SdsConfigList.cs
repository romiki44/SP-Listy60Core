using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListyApp.Models
{
    public class SdsConfigList
    {
        List<SdsConfig> AppConfigItems = new List<SdsConfig>();

        public SdsConfigList(List<SdsConfig> AppConfigItems)
        {
            if (AppConfigItems != null)
                this.AppConfigItems = AppConfigItems;
        }

        public List<SdsConfig> AppConfigs => AppConfigItems;

        public SdsConfig? GetItem(string itemName)
        {
            return AppConfigItems.FirstOrDefault(s => s.ItemName == itemName);
        }

        public string? GetItemValue(string itemName)
        {
            return AppConfigItems.FirstOrDefault(s => s.ItemName == itemName)?.ItemValue;
        }

        public int GetItemViewable(string itemName)
        {
            var sdsConfig = AppConfigItems.FirstOrDefault(s => s.ItemName == itemName);
            if (sdsConfig != null)
                return sdsConfig.Viewable;

            return 0;
        }

        public bool GetItemEditable(string itemName)
        {
            var sdsConfig = AppConfigItems.FirstOrDefault(s => s.ItemName == itemName);
            if (sdsConfig != null)
                return sdsConfig.Editable;

            return false;
        }

        public string GetSpracovanie()
        {
            return GetItemValue(SdsConfigItemNames.Spracovanie) ?? string.Empty;
        }

        public string? GetSpracovanieFolder()
        {
            var globaNetPath = GetItemValue(SdsConfigItemNames.AppNetPath) ?? string.Empty;
            var spracovanieFolder = GetItemValue(SdsConfigItemNames.SpracFolder) ?? string.Empty;

            return Path.Combine(globaNetPath, spracovanieFolder);
        }

        public string? GetTemplateFolder()
        {
            var globaNetPath = GetItemValue(SdsConfigItemNames.AppNetPath) ?? string.Empty;
            var spracovanieFolder = GetItemValue(SdsConfigItemNames.TemplateFolder) ?? string.Empty;

            return Path.Combine(globaNetPath, spracovanieFolder);
        }

        public string? GetEschrankyWordTemplate()
        {
            var templatePath = GetTemplateFolder() ?? string.Empty;
            string wordEschrankyTemplateFileName = GetItemValue(SdsConfigItemNames.EschrankyWordTemplate) ?? string.Empty;

            return Path.Combine(templatePath, wordEschrankyTemplateFileName);
        }

        public string? GetTlacWordTemplate()
        {
            var templatePath = GetTemplateFolder() ?? string.Empty;
            string wordTlacTemplateFileName = GetItemValue(SdsConfigItemNames.TlacWordTemplateSk) ?? string.Empty;

            return Path.Combine(templatePath, wordTlacTemplateFileName);
        }

        public string? GetEschrankyPdfExportFolder()
        {
            string spracFolder = GetSpracovanieFolder() ?? string.Empty;
            string pdfExportSubFolder = GetItemValue(SdsConfigItemNames.EschrankyPdfExportFolder) ?? string.Empty;

            return Path.Combine(spracFolder, pdfExportSubFolder);
        }

        public string? GetTlacPdfExportFolder()
        {
            string spracFolder = GetSpracovanieFolder() ?? string.Empty;
            string pdfExportSubFolder = GetItemValue(SdsConfigItemNames.TlacPdfExportFolder) ?? string.Empty;

            return Path.Combine(spracFolder, pdfExportSubFolder);
        }

        public string? GetOdcFolder()
        {
            var globaNetPath = GetItemValue(SdsConfigItemNames.AppNetPath) ?? string.Empty;
            var odcSubFolder = GetItemValue(SdsConfigItemNames.OdcSubFolder) ?? string.Empty;

            return Path.Combine(globaNetPath, odcSubFolder);
        }

        public int GetEschrankyPagesPerPdf()
        {
            string pagescount = GetItemValue(SdsConfigItemNames.EschrankyPagesPerPdf) ?? string.Empty;
            if (int.TryParse(pagescount, out int pagesPerPages))
                return pagesPerPages;

            return 0;
        }

        public int GetTlacMaxRecordPerPdf()
        {
            string maxrecords = GetItemValue(SdsConfigItemNames.TlacMaxRecordPerPdf) ?? string.Empty;
            if (int.TryParse(maxrecords, out int maxRecordPerPage))
                return maxRecordPerPage;

            return 0;
        }
    }
}
