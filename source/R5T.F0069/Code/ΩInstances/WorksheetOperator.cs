using System;


namespace R5T.F0069
{
    public class WorksheetOperator : IWorksheetOperator
    {
        #region Infrastructure

        public static IWorksheetOperator Instance { get; } = new WorksheetOperator();


        private WorksheetOperator()
        {
        }

        #endregion
    }
}


namespace R5T.F0069.Internal
{
    public class WorksheetOperator : IWorksheetOperator
    {
        #region Infrastructure

        public static IWorksheetOperator Instance { get; } = new WorksheetOperator();


        private WorksheetOperator()
        {
        }

        #endregion
    }
}