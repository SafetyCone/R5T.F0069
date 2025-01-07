using System;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069.Extensions
{
    public static class XlWorksheetExtensions
    {
        public static Worksheet To_Worksheet(this Xl.Worksheet worksheet, Workbook workbook)
            => Instances.WorksheetOperator.From(worksheet, workbook);
    }
}
