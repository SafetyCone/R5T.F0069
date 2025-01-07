using System;

using R5T.F0069;

using Xl = Microsoft.Office.Interop.Excel;

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
    public static Range GetRange(this Range upperLeft, RangeSize size)
    {
        var output = upperLeft.Worksheet.GetRange(upperLeft, size.Rows, size.Columns);
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

    public static Range GetOffset(this Range range, int row_offset, int column_offset)
    {
        var xlRange = range.XlRange.Offset[row_offset, column_offset];

        var offsetRange = new Range(xlRange, range.Worksheet);
        return offsetRange;
    }
}


namespace R5T.F0069.Extensions
{
    public static class RangeExtensions
    {
        public static void Clear(this Range range)
            => Instances.RangeOperator.Clear(range);

        public static void Copy(this Range range)
            => Instances.RangeOperator.Copy(range);

        public static string Get_Address(this Range range)
            => Instances.RangeOperator.Get_Address(range);

        public static Range Get_End_Down(this Range range)
            => Instances.RangeOperator.Get_End_Down(range);

        public static Range Get_End_Left(this Range range)
            => Instances.RangeOperator.Get_End_Left(range);

        public static Range Get_End_Right(this Range range)
            => Instances.RangeOperator.Get_End_Right(range);

        public static Range Get_End_Up(this Range range)
            => Instances.RangeOperator.Get_End_Up(range);

        public static Range Get_Offset(this Range range, int row_offset, int column_offset)
            => Instances.RangeOperator.Get_Offset(range, row_offset, column_offset);

        public static Range Get_UpperLeft(this Range range)
            => Instances.RangeOperator.Get_UpperLeft(range);

        public static Range Get_UpperRight(this Range range)
            => Instances.RangeOperator.Get_UpperRight(range);

        public static Range Get_LowerLeft(this Range range)
            => Instances.RangeOperator.Get_LowerLeft(range);

        public static Range Get_LowerRight(this Range range)
            => Instances.RangeOperator.Get_LowerRight(range);

        public static void PasteSpecial_Formulas(this Range range)
            => Instances.RangeOperator.PasteSpecial_Formulas(range);

        public static void Select(this Range range)
            => Instances.RangeOperator.Select(range);

        public static void Set_Formula(this Range range,
            string formula)
            => Instances.RangeOperator.Set_Value(range, formula);

        public static DateTime Get_Value_DateTime(this Range range)
            => Instances.RangeOperator.Get_Value_DateTime(range);

        public static void Set_Value(this Range range,
            DateTime value)
            => Instances.RangeOperator.Set_Value(range, value);

        public static int Get_Value_Integer(this Range range)
            => Instances.RangeOperator.Get_Value_Integer(range);

        public static void Set_Value(this Range range,
            int value)
            => Instances.RangeOperator.Set_Value(range, value);

        public static decimal Get_Value_Decimal(this Range range)
            => Instances.RangeOperator.Get_Value_Decimal(range);

        public static void Set_Value(this Range range,
            decimal value)
            => Instances.RangeOperator.Set_Value(range, value);

        public static double Get_Value_Double(this Range range)
            => Instances.RangeOperator.Get_Value_Double(range);

        public static void Set_Value(this Range range,
            double value)
            => Instances.RangeOperator.Set_Value(range, value);

        public static string Get_Value_String(this Range range)
            => Instances.RangeOperator.Get_Value_String(range);

        public static void Set_Value(this Range range,
            string value)
            => Instances.RangeOperator.Set_Value(range, value);

        public static Range Set_Values_UsingUpperLeft(this Range range,
            object[,] values)
            => Instances.RangeOperator.Set_Values_UsingUpperLeftOf(
                range,
                values);
    }
}
