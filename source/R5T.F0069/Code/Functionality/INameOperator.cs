using System;

using R5T.T0132;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069
{
    /// <summary>
    /// For Excel named ranges.
    /// </summary>
    [FunctionalityMarker]
    public partial interface INameOperator : IFunctionalityMarker
    {
        public Name From(Xl.Name name, Workbook workbook)
            => new Name(name, workbook);

        /// <summary>
        /// Determines whether a name is valid.
        /// </summary>
        /// <remarks>
        /// A name is valid if:
        /// <list type="bullet">
        /// <item>Its underlying <see cref="Name.XlName"/> is not null.</item>
        /// </list>
        /// </remarks>
        public bool Is_Valid(Name name)
        {
            var output = Instances.NullOperator.Is_NotNull(name.XlName);
            return output;
        }
    }
}
