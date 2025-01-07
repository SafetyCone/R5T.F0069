using System;
using System.Collections.Generic;
using System.Linq;

using R5T.T0132;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069.Internal
{
    /// <summary>
    /// An Excel workbook operator.
    /// </summary>
    [FunctionalityMarker]
    public partial interface IWorkbookOperator : IFunctionalityMarker
    {
        /// <summary>
        /// Enumerates the worksheet COM objects for the workbook.
        /// </summary>
        /// <remarks>
        /// Caller is responsible for release the returned worksheet COM object.
        /// This can be done by using the <see cref="Worksheet"/> type from the <see cref="F0069.IWorkbookOperator.Enumerate_Worksheets(Workbook)"/> method.
        /// </remarks>
        public IEnumerable<Xl.Worksheet> Enumerate_Worksheets(Xl.Workbook workbook)
        {
            // Use a for loop instead of a for-each loop since the worksheets collection will create an iterator that may retain a COM reference to the underlying COM worksheets collection.
            // Source: https://www.add-in-express.com/creating-addins-blog/release-excel-com-objects/
            var worksheets = workbook.Worksheets;

            var worksheets_start = 1; // 1-based.
            var worksheets_end = worksheets.Count + 1; // 1-based.

            for (int i = worksheets_start; i < worksheets_end; i++)
            {
                Xl.Worksheet worksheet = worksheets.Item[i];
                yield return worksheet;
            }

            Instances.MarshalOperator.Release_ComObject(worksheets);
        }

        /// <summary>
        /// Chooses <see cref="Equals_WithNullCheck(Xl.Workbook, Xl.Workbook)"/> as the default.
        /// </summary>
        public bool Equals(Xl.Workbook a, Xl.Workbook b)
            => this.Equals_WithNullCheck(a, b);

        public bool Equals_WithNullCheck(Xl.Workbook a, Xl.Workbook b)
            => Instances.EqualityOperator.NullCheckDeterminesEquality_Else(a, b,
                this.Equals_WithoutNullCheck);

        public bool Equals_WithoutNullCheck(Xl.Workbook a, Xl.Workbook b)
        {
            // Assume object (reference) equality.
            var output = a == b;
            return output;
        }

        public int Get_HashCode(Xl.Workbook workbook)
            // Use the object hashcode function.
            => workbook.GetHashCode();

        public string Get_Name(Xl.Workbook workbook)
            => workbook.Name;

        public Xl.Worksheet[] Get_Worksheets(Xl.Workbook workbook)
            => this.Enumerate_Worksheets(workbook)
            .Now();

        public Xl.Worksheet Get_Worksheet_First(Xl.Workbook workbook)
        {
            var output = workbook.Worksheets[1];
            return output;
        }

        public Xl.Name Add_NamedRange(
            Xl.Workbook workbook,
            Xl.Range range,
            string name)
        {
            var names = workbook.Names;

            var output = names.Add(name, range);

            Instances.MarshalOperator.Release_ComObject(names);

            return output;
        }

        public Xl.Name Set_NamedRange(
            Xl.Workbook workbook,
            Xl.Range range,
            string name)
        {
            var hasNamedRange = this.Has_NamedRange(
                workbook,
                name,
                out var xlName);

            if (hasNamedRange)
            {
                xlName.Delete();

                Instances.MarshalOperator.Release_ComObject(xlName);
            }

            var output = this.Add_NamedRange(
                workbook,
                range,
                name);

            return output;
        }

        public bool Has_NamedRange(
            Xl.Workbook workbook,
            string name,
            out Xl.Name xlName)
        {
            var names = workbook.Names;

            foreach (Xl.Name xlName_current in names)
            {
                if(xlName_current.Name == name)
                {
                    xlName = xlName_current;

                    return true;
                }
                else
                {
                    Instances.MarshalOperator.Release_ComObject(xlName_current);
                }
            }

            Instances.MarshalOperator.Release_ComObject(names);

            xlName = default;

            return false;
        }

        /// <summary>
        /// Does the workbook have a worksheet with the given name?
        /// If so, output the worksheet.
        /// </summary>
        public bool Has_Worksheet(
            Xl.Workbook workbook,
            string name,
            out Xl.Worksheet worksheet_OrNull)
        {
            var output = false;

            worksheet_OrNull = null;

            foreach (var worksheet in this.Enumerate_Worksheets(workbook))
            {
                var is_named = Instances.WorksheetOperator_Internal.Is_Named(
                    worksheet,
                    name);

                if(is_named)
                {
                    worksheet_OrNull = worksheet;

                    output = true;

                    break;
                }
            }

            return output;
        }
    }
}
