using System;


namespace R5T.F0069.Extensions
{
    public static class ApplicationExtensions
    {
        public static void Set_CalculationMode(this Application application,
            ExcelCalculationMode calculationMode)
            => Instances.ApplicationOperator.Set_CalculationMode(
                application,
                calculationMode);
    }
}
