using System;

using R5T.T0132;
using R5T.T0143;

using Xl = Microsoft.Office.Interop.Excel;

using R5T.F0069.Extensions;


namespace R5T.F0069
{
    [FunctionalityMarker]
    public partial interface IWorksheetOperator : IFunctionalityMarker
    {
#pragma warning disable IDE1006 // Naming Styles

        [Ignore]
        public Internal.IWorksheetOperator _Internal => Internal.WorksheetOperator.Instance;

#pragma warning restore IDE1006 // Naming Styles


        public bool Equals(Worksheet a, Worksheet b)
            => this.Equals_WithNullCheck(a, b);

        public bool Equals_WithNullCheck(Worksheet a, Worksheet b)
            => Instances.EqualityOperator.NullCheckDeterminesEquality_Else(a, b,
                this.Equals_WithoutNullCheck);

        public bool Equals_WithoutNullCheck(Worksheet a, Worksheet b)
            => _Internal.Equals_WithoutNullCheck(
                a.XlWorksheet,
                b.XlWorksheet);

        public Worksheet From(Xl.Worksheet xlWorksheet, Workbook workbook)
            => new Worksheet(xlWorksheet, workbook);

        public Range Get_Cell(
            Worksheet worksheet,
            int row_OneBased,
            int column_OneBased)
            => _Internal.Get_Cell(
                worksheet.XlWorksheet,
                row_OneBased,
                column_OneBased)
                .To_Range(worksheet);

        public int Get_HashCode(Worksheet worksheet)
            => _Internal.Get_HashCode(worksheet.XlWorksheet);

        public string Get_Name(Worksheet worksheet)
            => _Internal.Get_Name(worksheet.XlWorksheet);

        public Range Get_Range(
            Worksheet worksheet,
            Range upperLeft,
            Range lowerRight)
            => _Internal.Get_Range(
                worksheet.XlWorksheet,
                upperLeft.XlRange,
                lowerRight.XlRange)
                .To_Range(worksheet);

        public Range Get_Range_A1(Worksheet worksheet)
            => _Internal.Get_Range_A1(worksheet.XlWorksheet)
                .To_Range(worksheet);

        public bool Is_Named(
            Worksheet worksheet,
            string name)
            => _Internal.Is_Named(
                worksheet.XlWorksheet,
                name);

        /// <summary>
        /// Determines whether a worksheet is valid.
        /// </summary>
        /// <remarks>
        /// A worksheet is valid if:
        /// <list type="bullet">
        /// <item>Its underlying <see cref="Worksheet.XlWorksheet"/> is not null.</item>
        /// </list>
        /// </remarks>
        public bool Is_Valid(Worksheet worksheet)
        {
            var output = Instances.NullOperator.Is_NotNull(worksheet.XlWorksheet);
            return output;
        }

        public void Set_Name(
            Worksheet worksheet,
            string name)
        {
            _Internal.Set_Name(
                worksheet.XlWorksheet,
                name);
        }

        public void Verify_IsValid(Worksheet worksheet)
        {
            var is_Valid = this.Is_Valid(worksheet);
            if(!is_Valid)
            {
                throw new Exception("Worksheet was not valid.");
            }
        }
    }
}
