using System;

using Xl = Microsoft.Office.Interop.Excel;

using R5T.F0069;

using Range = R5T.F0069.Range;


public static class RangeExtensions
{
    public static void Calculate(this Range range)
    {
        range.XlRange.Calculate();
    }

    /// <summary>
    /// Gets a range of the specified size with this range as the upper-left corner.
    /// </summary>
    public static Range GetRange(this Range range, RangeSize size)
    {
        var output = range.Worksheet.GetRange(range, size.Rows, size.Columns);
        return output;
    }

    /// <summary>
    /// Gets a column.
    /// </summary>
    /// <param name="index">The zero-based column index.</param>
    public static Range GetColumn(this Range range, int index)
    {
        var counter = 0;
        foreach (Xl.Range xlColumn in range.XlRange.Columns)
        {
            if(counter == index)
            {
                var output = new Range(xlColumn, range.Worksheet);
                return output;
            }
            counter++;
        }

        throw new Exception();
    }

    public static void SetName(this Range range, string name)
    {
        range.Workbook.AddNamedRange(range, name);
    }

    public static Range GetOffset(this Range range, int rows, int columns)
    {
        var xlRange = range.XlRange.Offset[rows, columns];

        var offsetRange = new Range(xlRange, range.Worksheet);
        return offsetRange;
    }
}
