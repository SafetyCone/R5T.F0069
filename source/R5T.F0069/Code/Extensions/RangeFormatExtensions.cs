using System;

using Xl = Microsoft.Office.Interop.Excel;

using Range = R5T.F0069.Range;



public static class RangeFormatExtensions
{
    public static void AlignHorizontalLeft(this Range range)
    {
        range.XlRange.HorizontalAlignment = Xl.XlHAlign.xlHAlignLeft;
    }

    public static void AlignHorizontalRight(this Range range)
    {
        range.XlRange.HorizontalAlignment = Xl.XlHAlign.xlHAlignRight;
    }

    public static void AlignHorizontalCenter(this Range range)
    {
        range.XlRange.HorizontalAlignment = Xl.XlHAlign.xlHAlignCenter;
    }

    public static void AlignVerticalTop(this Range range)
    {
        range.XlRange.VerticalAlignment = Xl.XlVAlign.xlVAlignTop;
    }

    public static void AlignVerticalBottom(this Range range)
    {
        range.XlRange.VerticalAlignment = Xl.XlVAlign.xlVAlignBottom;
    }

    public static void AlignVerticalCenter(this Range range)
    {
        range.XlRange.VerticalAlignment = Xl.XlVAlign.xlVAlignCenter;
    }

    public static void Bold(this Range range, bool value)
    {
        range.XlRange.Font.Bold = value;
    }

    public static void Bold(this Range range)
    {
        range.Bold(true);
    }
}
