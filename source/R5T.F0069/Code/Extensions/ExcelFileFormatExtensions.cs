using System;

using Xl = Microsoft.Office.Interop.Excel;

using R5T.F0069;



public static class ExcelFileFormatExtensions
{
    internal static Xl.XlFileFormat ToXlFileFormat(this ExcelFileFormat excelFileFormat)
    {
        return excelFileFormat switch
        {
            ExcelFileFormat.CSV => Xl.XlFileFormat.xlCSV,
            ExcelFileFormat.XLS => Xl.XlFileFormat.xlExcel8,
            ExcelFileFormat.XLSM => Xl.XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
            _ => Xl.XlFileFormat.xlOpenXMLWorkbook,
        };
    }
}
