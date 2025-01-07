using System;


namespace R5T.F0069
{
    public class ApplicationOperator : IApplicationOperator
    {
        #region Infrastructure

        public static IApplicationOperator Instance { get; } = new ApplicationOperator();


        private ApplicationOperator()
        {
        }

        #endregion
    }
}


namespace R5T.F0069.Internal
{
    public class ApplicationOperator : IApplicationOperator
    {
        #region Infrastructure

        public static IApplicationOperator Instance { get; } = new ApplicationOperator();


        private ApplicationOperator()
        {
        }

        #endregion
    }
}