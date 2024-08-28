using System;


namespace R5T.F0069
{
    public static class Instances
    {
        public static L0066.IConversionOperator ConversionOperator => L0066.ConversionOperator.Instance;
        public static L0066.IEnumerationOperator EnumerationOperator => L0066.EnumerationOperator.Instance;
        public static L0066.ISwitchOperator SwitchOperator => L0066.SwitchOperator.Instance;
        public static IValues Values => F0069.Values.Instance;
    }
}