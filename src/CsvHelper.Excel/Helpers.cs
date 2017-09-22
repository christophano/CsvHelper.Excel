
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("CsvHelper.Excel.Specs")]

namespace CsvHelper.Excel
{
    using System.IO;
    using ClosedXML.Excel;

    internal static class Helpers
    {
        public static XLWorkbook GetOrCreateWorkbook(string path, string worksheetName)
        {
            if (!File.Exists(path))
            {
                var workbook = new XLWorkbook(XLEventTracking.Disabled);
                workbook.GetOrAddWorksheet(worksheetName);
                workbook.SaveAs(path);
                return workbook;
            }
            return new XLWorkbook(path, XLEventTracking.Disabled);
        }

        public static IXLWorksheet GetOrAddWorksheet(this XLWorkbook workbook, string sheetName)
        {
            if (!workbook.TryGetWorksheet(sheetName, out var worksheet))
            {
                worksheet = workbook.AddWorksheet(sheetName);
            }
            return worksheet;
        }
    }
}
