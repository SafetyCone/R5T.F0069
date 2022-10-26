using System;

using Xl = Microsoft.Office.Interop.Excel;



public static class XlSheetsExtensions
{
    internal static Xl.Worksheet AddWorksheet(this Xl.Sheets sheets)
    {
        var worksheet = sheets.Add() as Xl.Worksheet;
        return worksheet;
    }
}
