using System;

using Xl = Microsoft.Office.Interop.Excel;

using R5T.F0069;

using Instances = R5T.F0069.Instances;


public static class XlCalculationExtensions
{
    public static ExcelCalculationMode ToExcelCalculationMode(this Xl.XlCalculation xlCalculation)
    {
        var output = xlCalculation switch
        {
            Xl.XlCalculation.xlCalculationAutomatic => ExcelCalculationMode.Automatic,
            Xl.XlCalculation.xlCalculationManual => ExcelCalculationMode.Manual,
            Xl.XlCalculation.xlCalculationSemiautomatic => ExcelCalculationMode.SemiAutomatic,
            _ => throw Instances.SwitchOperator.Get_DefaultCaseException(xlCalculation),
        };

        return output;
    }
}
