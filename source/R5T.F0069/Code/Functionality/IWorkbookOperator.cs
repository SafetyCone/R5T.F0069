using System;
using System.Collections.Generic;
using System.Linq;

using R5T.T0132;
using R5T.T0143;

using R5T.F0069.Extensions;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069
{
    /// <summary>
    /// An Excel workbook operator.
    /// </summary>
    [FunctionalityMarker]
    public partial interface IWorkbookOperator : IFunctionalityMarker
    {
#pragma warning disable IDE1006 // Naming Styles

        [Ignore]
        public Internal.IWorkbookOperator _Internal => Internal.WorkbookOperator.Instance;

#pragma warning restore IDE1006 // Naming Styles


        public Name Add_NamedRange(
            Workbook workbook,
            Range range,
            string name)
            => _Internal.Add_NamedRange(
                workbook.XlWorkbook,
                range.XlRange,
                name)
                .To_Name(workbook);

        public Name Set_NamedRange(
            Workbook workbook,
            Range range,
            string name)
            => _Internal.Set_NamedRange(
                workbook.XlWorkbook,
                range.XlRange,
                name)
                .To_Name(workbook);

        public bool Has_NamedRange(
            Workbook workbook,
            string name,
            out Name nameObj)
        {
            var output = _Internal.Has_NamedRange(
                workbook.XlWorkbook,
                name,
                out var xlName);

            nameObj = output
                ? xlName.To_Name(workbook)
                : default
                ;

            return output;
        }

        public IEnumerable<Worksheet> Enumerate_Worksheets(Workbook workbook)
            => _Internal.Enumerate_Worksheets(workbook.XlWorkbook)
                .Select(x => x.To_Worksheet(workbook))
                ;

        public bool Equals(Workbook a, Workbook b)
            => this.Equals_WithNullCheck(a, b);

        public bool Equals_WithNullCheck(Workbook a, Workbook b)
            => Instances.EqualityOperator.NullCheckDeterminesEquality_Else(a, b,
                this.Equals_WithoutNullCheck);

        public bool Equals_WithoutNullCheck(Workbook a, Workbook b)
            => _Internal.Equals_WithoutNullCheck(
                a.XlWorkbook,
                b.XlWorkbook);

        public Workbook From(Xl.Workbook xlWorkbook, Application application)
            => new Workbook(xlWorkbook, application);

        public int Get_HashCode(Workbook workbook)
            => _Internal.Get_HashCode(workbook.XlWorkbook);

        public string Get_Name(Workbook workbook)
            => _Internal.Get_Name(workbook.XlWorkbook);

        public Worksheet Get_Worksheet_First(Workbook workbook)
            => _Internal.Get_Worksheet_First(workbook.XlWorkbook)
                .To_Worksheet(workbook);

        public Worksheet[] Get_Worsheets(Workbook workbook)
            => this.Enumerate_Worksheets(workbook)
                .Now();

        /// <summary>
        /// Determines whether a workbook is valid.
        /// </summary>
        /// <remarks>
        /// A workbook is valid if:
        /// <list type="bullet">
        /// <item>Its underlying <see cref="Workbook.XlWorkbook"/> is not null.</item>
        /// </list>
        /// </remarks>
        public bool Is_Valid(Workbook workbook)
        {
            var output = Instances.NullOperator.Is_NotNull(workbook.XlWorkbook);
            return output;
        }

        public void Verify_IsValid(Workbook workbook)
        {
            var is_Valid = this.Is_Valid(workbook);
            if (!is_Valid)
            {
                throw new Exception("Workbook was not valid.");
            }
        }
    }
}
