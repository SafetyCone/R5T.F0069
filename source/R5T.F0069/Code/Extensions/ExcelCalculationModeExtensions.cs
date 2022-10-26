using System;

using Xl = Microsoft.Office.Interop.Excel;

using R5T.F0069;

using Instances = R5T.F0069.Instances;


public static class ExcelCalculationModeExtensions
{
    public static Xl.XlCalculation ToXlCalculation(this ExcelCalculationMode mode)
    {
        switch(mode)
        {
            case ExcelCalculationMode.Automatic:
                return Xl.XlCalculation.xlCalculationAutomatic;

            case ExcelCalculationMode.Manual:
                return Xl.XlCalculation.xlCalculationManual;

            case ExcelCalculationMode.SemiAutomatic:
                return Xl.XlCalculation.xlCalculationSemiautomatic;

            default:
                throw Instances.EnumerationOperator.SwitchDefaultCaseException(mode);
        }
    }
}
