using System;

using R5T.T0132;
using R5T.T0143;

using R5T.F0069.Extensions;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069
{
    [FunctionalityMarker]
    public partial interface IRangeOperator : IFunctionalityMarker
    {
#pragma warning disable IDE1006 // Naming Styles

        [Ignore]
        public Internal.IRangeOperator _Internal => Internal.RangeOperator.Instance;

#pragma warning restore IDE1006 // Naming Styles

        public void Clear(Range range)
            => _Internal.Clear(range.XlRange);

        public void Copy(Range range)
            => _Internal.Copy(range.XlRange);

        public Range From(Xl.Range xlRange, Worksheet worksheet)
            => new Range(xlRange, worksheet);

        public string Get_Address(Range range)
            => _Internal.Get_Address(range.XlRange);

        public string Get_Name(Range range)
        {
            var worksheetName = range.Worksheet.Name;
            var rangeAddress = Instances.RangeOperator.Get_Address(range);

            var output = $"{worksheetName}{Instances.Strings.ExclamationPoint}{rangeAddress}";
            return output;
        }

        public Range Get_Offset(
            Range range,
            int row_offset,
            int column_offset)
            => _Internal.Get_Offset(range.XlRange, row_offset, column_offset)
                .To_Range(range.Worksheet);

        public Range Get_UpperLeft(Range range)
            => _Internal.Get_UpperLeft(range.XlRange)
                .To_Range(range.Worksheet);

        public Range Get_UpperRight(Range range)
            => _Internal.Get_UpperRight(range.XlRange)
                .To_Range(range.Worksheet);

        public Range Get_LowerLeft(Range range)
            => _Internal.Get_LowerLeft(range.XlRange)
                .To_Range(range.Worksheet);

        public Range Get_LowerRight(Range range)
            => _Internal.Get_LowerRight(range.XlRange)
                .To_Range(range.Worksheet);

        /// <summary>
        /// Determines whether a range is valid.
        /// </summary>
        /// <remarks>
        /// A range is valid if:
        /// <list type="bullet">
        /// <item>Its underlying <see cref="Range.XlRange"/> is not null.</item>
        /// </list>
        /// </remarks>
        public bool Is_Valid(Range range)
        {
            var output = Instances.NullOperator.Is_NotNull(range.XlRange);
            return output;
        }

        public Range Get_End_Down(Range range)
            => _Internal.Get_End_Down(range.XlRange)
                .To_Range(range.Worksheet);

        public Range Get_End_Left(Range range)
            => _Internal.Get_End_Left(range.XlRange)
                .To_Range(range.Worksheet);

        public Range Get_End_Right(Range range)
            => _Internal.Get_End_Right(range.XlRange)
                .To_Range(range.Worksheet);

        public Range Get_End_Up(Range range)
            => _Internal.Get_End_Up(range.XlRange)
                .To_Range(range.Worksheet);

        public void Get_Formula(Range range)
            => _Internal.Get_Formula(range.XlRange);

        public void Set_Formula(
            Range range,
            string formula)
            => _Internal.Set_Formula(range.XlRange, formula);

        public object[,] Get_Values(Range range)
            => _Internal.Get_Values(range.XlRange);

        public void PasteSpecial_Formulas(Range range)
            => _Internal.PasteSpecial_Formulas(range.XlRange);

        public void Select(Range range)
            => _Internal.Select(range.XlRange);

        public void Set_Values(
            Range range,
            object[,] values)
            => _Internal.Set_Values(
                range.XlRange,
                values);

        public Range Set_Values_UsingUpperLeftOf(
            Range range,
            object[,] values)
            => _Internal.Set_Values_UsingUpperLeftOf(
                range.XlRange,
                values)
                .To_Range(range.Worksheet);

        public DateTime Get_Value_DateTime(Range range)
            => _Internal.Get_Value_DateTime(range.XlRange);

        public void Set_Value(
            Range range,
            DateTime dateTime)
            => _Internal.Set_Value(
                range.XlRange,
                dateTime);

        public int Get_Value_Integer(Range range)
            => _Internal.Get_Value_Integer(range.XlRange);

        public void Set_Value(
            Range range,
            int value)
            => _Internal.Set_Value(range.XlRange, value);

        public decimal Get_Value_Decimal(Range range)
            => _Internal.Get_Value_Decimal(range.XlRange);

        public void Set_Value(
            Range range,
            decimal value)
            => _Internal.Set_Value(range.XlRange, value);

        public double Get_Value_Double(Range range)
            => _Internal.Get_Value_Double(range.XlRange);

        public void Set_Value(
            Range range,
            double value)
            => _Internal.Set_Value(range.XlRange, value);

        public string Get_Value_String(Range range)
            => _Internal.Get_Value_String(range.XlRange);

        public void Set_Value(
            Range range,
            string value)
            => _Internal.Set_Value(range.XlRange, value);

        public void Verify_IsValid(Range range)
        {
            var is_Valid = this.Is_Valid(range);
            if (!is_Valid)
            {
                throw new Exception("Range was not valid.");
            }
        }
    }
}
