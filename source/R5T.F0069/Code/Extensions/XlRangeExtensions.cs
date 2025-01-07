using System;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069.Extensions
{
    public static class XlRangeExtensions
    {
        public static Range To_Range(this Xl.Range range, Worksheet worksheet)
            => Instances.RangeOperator.From(range, worksheet);
    }
}
