using System;


namespace R5T.F0069
{
    public class WorkbookOperator : IWorkbookOperator
    {
        #region Infrastructure

        public static IWorkbookOperator Instance { get; } = new WorkbookOperator();


        private WorkbookOperator()
        {
        }

        #endregion
    }
}


namespace R5T.F0069.Internal
{
    public class WorkbookOperator : IWorkbookOperator
    {
        #region Infrastructure

        public static IWorkbookOperator Instance { get; } = new WorkbookOperator();


        private WorkbookOperator()
        {
        }

        #endregion
    }
}