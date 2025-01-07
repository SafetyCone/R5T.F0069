using System;

using R5T.T0142;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069
{
    [UtilityTypeMarker]
    public class Name : IDisposable
    {
        #region IDisposable

        private bool zDisposed = false; // To detect redundant calls.


        public void Dispose()
        {
            this.Dispose(true);

            GC.SuppressFinalize(this);
        }

        // Remove the virtual call if the class is sealed (or has no plans for subclassing, in which case this should be communicated by sealing the class).
        private void Dispose(bool disposing)
        {
            if (!this.zDisposed)
            {
                if (disposing)
                {
                    // Do nothing.
                    /// The <see cref="Xl.Range"/> object itself is managed, the Excel application it is the handle to is not.
                }

                Instances.MarshalOperator.Release_ComObject(this.XlName);

                this.XlName = null;
            }

            this.zDisposed = true;
        }

        ~Name()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            this.Dispose(false);
        }

        #endregion


        /// <summary>
        /// The underlying Excel COM automation range object.
        /// </summary>
        /// <remarks>
        /// Note: will never be null.
        /// See <see cref="INameOperator.Is_Valid(Name)"/>.
        /// </remarks>
        internal Xl.Name XlName { get; private set; }

        public Workbook Workbook { get; private set; }


        internal Name(Xl.Name xlName, Workbook workbook)
        {
            this.XlName = xlName;
            this.Workbook = workbook;
        }
    }
}
