using System;

using Xl = Microsoft.Office.Interop.Excel;

using R5T.F0069;


public static class XlFileFormatExtensions
{
    internal static ExcelFileFormat ToExcelFileFormat(this Xl.XlFileFormat xlFileFormat)
    {
        return xlFileFormat switch
        {
            Xl.XlFileFormat.xlCSV => ExcelFileFormat.CSV,
            Xl.XlFileFormat.xlExcel8 => ExcelFileFormat.XLS,
            Xl.XlFileFormat.xlOpenXMLWorkbookMacroEnabled => ExcelFileFormat.XLSM,
            Xl.XlFileFormat.xlOpenXMLWorkbook => ExcelFileFormat.XLSX,
            _ => ExcelFileFormat.Other,
        };
    }
}
