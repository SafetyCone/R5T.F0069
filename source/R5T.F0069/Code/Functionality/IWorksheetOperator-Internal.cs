using System;

using R5T.T0132;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069.Internal
{
    [FunctionalityMarker]
    public partial interface IWorksheetOperator : IFunctionalityMarker
    {
        /// <summary>
        /// Chooses <see cref="Equals_WithNullCheck(Xl.Worksheet, Xl.Worksheet)"/> as the default.
        /// </summary>
        public bool Equals(Xl.Worksheet a, Xl.Worksheet b)
            => this.Equals_WithNullCheck(a, b);

        public bool Equals_WithNullCheck(Xl.Worksheet a, Xl.Worksheet b)
            => Instances.EqualityOperator.NullCheckDeterminesEquality_Else(a, b,
                this.Equals_WithoutNullCheck);

        public bool Equals_WithoutNullCheck(Xl.Worksheet a, Xl.Worksheet b)
        {
            // Assume object (reference) equality.
            var output = a == b;
            return output;
        }

        public int Get_HashCode(Xl.Worksheet worksheet)
            // Use the object hashcode function.
            => worksheet.GetHashCode();

        public string Get_Name(Xl.Worksheet worksheet)
            => worksheet.Name;

        public Xl.Range Get_Range(
            Xl.Worksheet worksheet,
            Xl.Range upperLeft,
            Xl.Range lowerRight)
        {
            var output = worksheet.Range[upperLeft, lowerRight];
            return output;
        }

        public Xl.Range Get_Cell(
            Xl.Worksheet worksheet,
            int row_OneBased,
            int columnn_OneBased)
            => worksheet.Cells[row_OneBased, columnn_OneBased] as Xl.Range;

        public Xl.Range Get_Range_A1(Xl.Worksheet worksheet)
            => this.Get_Cell(worksheet, 1, 1);

        public bool Is_Named(
            Xl.Worksheet worksheet,
            string name)
        {
            var worksheet_name = this.Get_Name(worksheet);

            var output = worksheet_name == name;
            return output;
        }

        public void Set_Name(
            Xl.Worksheet worksheet,
            string name)
        {
            worksheet.Name = name;
        }
    }
}
