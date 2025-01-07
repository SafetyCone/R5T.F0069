using System;

using Xl = Microsoft.Office.Interop.Excel;

using R5T.F0069;


public static class WorksheetExtensions
{
    public static void Calculate(this Worksheet worksheet)
    {
        worksheet.XlWorksheet.Calculate();
    }

    public static void Show(this Worksheet worksheet)
    {
        worksheet.XlWorksheet.Visible = Xl.XlSheetVisibility.xlSheetVisible;
    }

    public static void Hide(this Worksheet worksheet)
    {
        worksheet.XlWorksheet.Visible = Xl.XlSheetVisibility.xlSheetHidden;
    }

    public static void HideVeryHidden(this Worksheet worksheet)
    {
        worksheet.XlWorksheet.Visible = Xl.XlSheetVisibility.xlSheetVeryHidden;
    }

    public static void SetColumnWidths(this Worksheet worksheet, params double[] columnWidths)
    {
        var range = worksheet.GetA1Range();
        foreach (var columnWidth in columnWidths)
        {
            range.ColumnWidth = columnWidth;

            range = range.GetOffset(0, 1);
        }
    }
}


namespace R5T.F0069.Extensions
{
    public static class WorksheetExtensions
    {
        public static Range Get_Cell(this Worksheet worksheet,
            int row_OneBased,
            int column_OneBased)
            => Instances.WorksheetOperator.Get_Cell(
                worksheet,
                row_OneBased,
                column_OneBased);

        public static Range Get_Range(this Worksheet worksheet,
            Range upperLeft,
            Range lowerRight)
            => Instances.WorksheetOperator.Get_Range(
                worksheet,
                upperLeft,
                lowerRight);

        public static Range Get_Range_A1(this Worksheet worksheet)
            => Instances.WorksheetOperator.Get_Range_A1(worksheet);
    }
}