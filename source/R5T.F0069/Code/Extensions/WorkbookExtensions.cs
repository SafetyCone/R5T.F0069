using System;

using Xl = Microsoft.Office.Interop.Excel;

using R5T.F0069;

using Range = R5T.F0069.Range;


public static class WorkbookExtensions
{
    public static Worksheet NewWorksheet(this Workbook workbook, string name)
    {
        var worksheet = workbook.NewWorksheet();

        worksheet.Name = name;

        return worksheet;
    }

    public static bool HasWorksheet(this Workbook workbook, string name)
    {
        var output = false;
        foreach (Xl.Worksheet worksheet in workbook.XlWorkbook.Worksheets)
        {
            if (name == worksheet.Name)
            {
                output = true;
                break;
            }
        }

        return output;
    }

    public static void DeleteWorksheet(this Workbook workbook, string name)
    {
        var worksheet = workbook.GetWorksheet(name);

        worksheet.Delete();
    }

    public static void AddNamedRange(this Workbook workbook, Range range, string name)
    {
        workbook.XlWorkbook.Names.Add(name, range.XlRange);
    }

    public static bool HasNamedRange(this Workbook workbook, string name)
    {
        foreach (Xl.Name xlName in workbook.XlWorkbook.Names)
        {
            if(name == xlName.Name)
            {
                return true;
            }
        }

        return false;
    }

    public static Range GetNamedRange(this Workbook workbook, string name)
    {
        var xlName = workbook.XlWorkbook.Names.Item(name);
        var xlNamedRange = xlName.RefersToRange;
        var xlWorksheet = xlNamedRange.Worksheet;

        var worksheet = new Worksheet(xlWorksheet, workbook);
        var namedRange = new Range(xlNamedRange, worksheet);
        return namedRange;
    }

    /// <summary>
    /// Calculate the workbook.
    /// </summary>
    /// <remarks>
    /// Despite the fact the Xl.Workbook type has no calculation method, the application-level calcuation method is placed here since it is an error to calculate without a workbook present.
    /// </remarks>
    public static void Calculate(this Workbook workbook)
    {
        workbook.Application.XlApplication.Calculate();
    }

    public static void Save(this Workbook workbook)
    {
        workbook.XlWorkbook.Save();
    }
}


namespace R5T.F0069.Extensions
{
    public static class WorkbookExtensions
    {
        public static Worksheet Get_Worksheet_First(this Workbook workbook)
            => Instances.WorkbookOperator.Get_Worksheet_First(workbook);

        public static Name Set_NamedRange(this Workbook workbook,
            Range range,
            string name)
            => Instances.WorkbookOperator.Set_NamedRange(
                workbook,
                range,
                name);
    }
}