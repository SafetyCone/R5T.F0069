using System;

using Xl = Microsoft.Office.Interop.Excel;

using R5T.F0069;

using Instances = R5T.F0069.Instances;


public static class ExcelCalculationModeExtensions
{
    public static Xl.XlCalculation ToXlCalculation(this ExcelCalculationMode mode)
    {
        var output = mode switch
        {
            ExcelCalculationMode.Automatic => Xl.XlCalculation.xlCalculationAutomatic,
            ExcelCalculationMode.Manual => Xl.XlCalculation.xlCalculationManual,
            ExcelCalculationMode.SemiAutomatic => Xl.XlCalculation.xlCalculationSemiautomatic,
            _ => throw Instances.SwitchOperator.Get_DefaultCaseException(mode),
        };

        return output;
    }
}
