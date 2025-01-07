using System;


namespace R5T.F0069
{
    public class RangeSizeOperator : IRangeSizeOperator
    {
        #region Infrastructure

        public static IRangeSizeOperator Instance { get; } = new RangeSizeOperator();


        private RangeSizeOperator()
        {
        }

        #endregion
    }
}
