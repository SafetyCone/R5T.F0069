using System;

using R5T.T0132;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069.Internal
{
    [FunctionalityMarker]
    public partial interface IRangeOperator : IFunctionalityMarker
    {
        public void Clear(Xl.Range range)
        {
            range.Clear();
        }

        public void Copy(Xl.Range range)
        {
            range.Copy();
        }

        public string Get_Address(Xl.Range range)
        {
            var output = range.Address[false, false];
            return output;
        }

        public int Get_ColumnCount(Xl.Range range)
            => this.Get_Count_OfColumns(range);

        public int Get_Count_OfColumns(Xl.Range range)
        {
            var output = range.Columns.Count;
            return output;
        }

        public int Get_RowCount(Xl.Range range)
            => this.Get_Count_OfRows(range);

        public int Get_Count_OfRows(Xl.Range range)
        {
            var output = range.Rows.Count;
            return output;
        }

        public Xl.Range Get_Offset(
            Xl.Range range,
            RangeSize rangeSize)
            => this.Get_Offset(
                range,
                rangeSize.Rows - 1,
                rangeSize.Columns - 1);

        public Xl.Range Get_Offset(
            Xl.Range range,
            int row_offset,
            int column_offset)
        {
            var output = range.Offset[row_offset, column_offset];
            return output;
        }

        public Xl.Range Get_UpperLeft(Xl.Range range)
        {
            var output = range.Cells[1, 1];
            return output;
        }

        public Xl.Range Get_UpperRight(Xl.Range range)
        {
            var upperLeft = this.Get_UpperLeft(range);

            var columns_count = this.Get_ColumnCount(range);

            var output = this.Get_Offset(upperLeft, 0, columns_count - 1);
            return output;
        }

        public Xl.Range Get_LowerLeft(Xl.Range range)
        {
            var upperLeft = this.Get_UpperLeft(range);

            var rows_count = this.Get_RowCount(range);

            var output = this.Get_Offset(upperLeft, rows_count - 1, 0);
            return output;
        }

        public Xl.Range Get_LowerRight(Xl.Range range)
        {
            var upperLeft = this.Get_UpperLeft(range);

            var rows_count = this.Get_RowCount(range);
            var columns_count = this.Get_ColumnCount(range);

            var output = this.Get_Offset(upperLeft, rows_count - 1, columns_count - 1);
            return output;
        }

        public Xl.Range Get_End_Down(Xl.Range range)
        {
            var output = range.End[Xl.XlDirection.xlDown];
            return output;
        }

        public Xl.Range Get_End_Left(Xl.Range range)
        {
            var output = range.End[Xl.XlDirection.xlToLeft];
            return output;
        }

        public Xl.Range Get_End_Right(Xl.Range range)
        {
            var output = range.End[Xl.XlDirection.xlToRight];
            return output;
        }

        public Xl.Range Get_End_Up(Xl.Range range)
        {
            var output = range.End[Xl.XlDirection.xlUp];
            return output;
        }

        public string Get_Formula(Xl.Range range)
        {
            var output = range.Formula;
            return output;
        }

        public void PasteSpecial_Formulas(Xl.Range range)
        {
            range.PasteSpecial(Xl.XlPasteType.xlPasteFormulas);
        }

        public void Select(Xl.Range range)
        {
            // Need to activate the worksheet before selecting.
            range.Worksheet.Activate();

            range.Select();
        }

        public void Set_Formula(
            Xl.Range range,
            string formula)
        {
            range.Formula = formula;
        }

        public object[,] Get_Values(Xl.Range range)
        {
            var output = range.Value as object[,];
            return output;
        }

        /// <summary>
        /// Set the values of the given range.
        /// </summary>
        public void Set_Values(
            Xl.Range range,
            object[,] values)
        {
            range.Value = values;
        }

        /// <summary>
        /// Set values starting using the given range as the upper-left starting point.
        /// </summary>
        public Xl.Range Set_Values_WithUpperLeft(
            Xl.Range upperLeft,
            object[,] values)
        {
            var rangeSize = Instances.RangeSizeOperator.Get_Size(values);

            var range = this.Get_Range_WithUpperLeft(
                upperLeft,
                rangeSize);

            this.Set_Values(
                range,
                values);

            return range;
        }

        /// <summary>
        /// Given a range, use the upper-left of the given range to get a range set to contain the given values.
        /// </summary>
        public Xl.Range Set_Values_UsingUpperLeftOf(
            Xl.Range range,
            object[,] values)
        {
            var upperLeft = this.Get_UpperLeft(range);

            var output = this.Set_Values_WithUpperLeft(
                upperLeft,
                values);

            return output;
        }

        /// <summary>
        /// Get a range of the given size, using the given range as the upper-left starting point.
        /// </summary>
        public Xl.Range Get_Range_WithUpperLeft(
            Xl.Range upperLeft,
            RangeSize rangeSize)
        {
            var worksheet = upperLeft.Worksheet;

            var lowerRight = this.Get_Offset(
                upperLeft,
                rangeSize);

            var output = Instances.WorksheetOperator_Internal.Get_Range(
                worksheet,
                upperLeft,
                lowerRight);

            return output;
        }

        public DateTime Get_Value_DateTime(Xl.Range range)
        {
            var output = range.Value2;
            return output;
        }

        public void Set_Value(
            Xl.Range range,
            DateTime dateTime)
        {
            range.Value2 = dateTime;
        }

        public int Get_Value_Integer(Xl.Range range)
        {
            var output = range.Value2;
            return output;
        }

        public void Set_Value(
            Xl.Range range,
            int value)
        {
            range.Value2 = value;
        }

        public decimal Get_Value_Decimal(Xl.Range range)
        {
            var output = range.Value2;
            return output;
        }

        public void Set_Value(
            Xl.Range range,
            decimal value)
        {
            range.Value2 = value;
        }

        public double Get_Value_Double(Xl.Range range)
        {
            var output = range.Value2;
            return output;
        }

        public void Set_Value(
            Xl.Range range,
            double value)
        {
            range.Value2 = value;
        }

        public string Get_Value_String(Xl.Range range)
        {
            var output = range.Value2;
            return output;
        }

        public void Set_Value(
            Xl.Range range,
            string value)
        {
            range.Value2 = value;
        }
    }
}
