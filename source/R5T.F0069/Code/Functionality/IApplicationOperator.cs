using System;

using R5T.T0132;
using R5T.T0143;


namespace R5T.F0069
{
    [FunctionalityMarker]
    public partial interface IApplicationOperator : IFunctionalityMarker
    {
#pragma warning disable IDE1006 // Naming Styles

        [Ignore]
        public Internal.IApplicationOperator _Internal => Internal.ApplicationOperator.Instance;

#pragma warning restore IDE1006 // Naming Styles


        public void Set_CalculationMode(
            Application application,
            ExcelCalculationMode calculationMode)
        {
            var xlCalculationMode = calculationMode.ToXlCalculation();

            _Internal.Set_CalculationMode(
                application.XlApplication,
                xlCalculationMode);
        }
    }
}
