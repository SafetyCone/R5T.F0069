using System;

using R5T.T0132;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069.Internal
{
    [FunctionalityMarker]
    public partial interface IApplicationOperator : IFunctionalityMarker
    {
        public void Set_CalculationMode(
            Xl.Application application,
            Xl.XlCalculation calculationMode)
        {
            application.Calculation = calculationMode;
        }
    }
}
