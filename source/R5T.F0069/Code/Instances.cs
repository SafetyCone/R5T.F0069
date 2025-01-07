using System;


namespace R5T.F0069
{
    public static class Instances
    {
        public static IApplicationOperator ApplicationOperator => F0069.ApplicationOperator.Instance;
        public static L0066.IConversionOperator ConversionOperator => L0066.ConversionOperator.Instance;
        public static L0066.IEnumerationOperator EnumerationOperator => L0066.EnumerationOperator.Instance;
        public static L0066.IEqualityOperator EqualityOperator => L0066.EqualityOperator.Instance;
        public static L0066.IMarshalOperator MarshalOperator => L0066.MarshalOperator.Instance;
        public static INameOperator NameOperator => F0069.NameOperator.Instance;
        public static L0066.INullOperator NullOperator => L0066.NullOperator.Instance;
        public static IRangeOperator RangeOperator => F0069.RangeOperator.Instance;
        public static Internal.IRangeOperator RangeOperator_Internal => Internal.RangeOperator.Instance;
        public static IRangeSizeOperator RangeSizeOperator => F0069.RangeSizeOperator.Instance;
        public static L0066.IStrings Strings => L0066.Strings.Instance;
        public static L0066.ISwitchOperator SwitchOperator => L0066.SwitchOperator.Instance;
        public static IValues Values => F0069.Values.Instance;
        public static IWorkbookOperator WorkbookOperator => F0069.WorkbookOperator.Instance;
        public static Internal.IWorkbookOperator WorkbookOperator_Internal => Internal.WorkbookOperator.Instance;
        public static IWorksheetOperator WorksheetOperator => F0069.WorksheetOperator.Instance;
        public static Internal.IWorksheetOperator WorksheetOperator_Internal => Internal.WorksheetOperator.Instance;
    }
}