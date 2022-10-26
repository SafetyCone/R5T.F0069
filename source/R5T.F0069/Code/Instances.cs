using System;

using R5T.F0000;


namespace R5T.F0069
{
    public static class Instances
    {
        public static IConversionOperator ConversionOperator { get; } = F0000.ConversionOperator.Instance;
        public static IEnumerationOperator EnumerationOperator { get; } = F0000.EnumerationOperator.Instance;
        public static IValues Values { get; } = F0069.Values.Instance;
    }
}