using System;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069.Extensions
{
    public static class XlNameExtensions
    {
        public static Name To_Name(this Xl.Name name, Workbook workbook)
            => Instances.NameOperator.From(name, workbook);
    }
}
