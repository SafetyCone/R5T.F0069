using System;


namespace R5T.F0069
{
    public class RangeOperator : IRangeOperator
    {
        #region Infrastructure

        public static IRangeOperator Instance { get; } = new RangeOperator();


        private RangeOperator()
        {
        }

        #endregion
    }
}


namespace R5T.F0069.Internal
{
    public class RangeOperator : IRangeOperator
    {
        #region Infrastructure

        public static IRangeOperator Instance { get; } = new RangeOperator();


        private RangeOperator()
        {
        }

        #endregion
    }
}