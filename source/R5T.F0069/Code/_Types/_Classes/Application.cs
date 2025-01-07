using System;

using R5T.T0142;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069
{
    /// <summary>
    /// Represents an Excel application.
    /// </summary>
    [UtilityTypeMarker]
    public sealed class Application : IDisposable
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
                if(disposing)
                {
                    // Do nothing.
                    /// The <see cref="Xl.Application"/> object itself is managed, the Excel application it is the handle to is not.
                }

                this.XlApplication.DisplayAlerts = false;
                
                this.XlApplication.Quit();

                Instances.MarshalOperator.FinalRelease_ComObject(this.XlApplication);

                this.XlApplication = null;

                // Yes, this is needed twice.
                // The first call releases *our* disposable objects, such that then the COM runtime callable wrappers (RCWs) are then without references.
                // The second call releases the RCWs.
                // This double-call is placed here at the application level such that the expensive garbage collection process is only called once.
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            this.zDisposed = true;
        }

        ~Application()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            this.Dispose(false);
        }

        #endregion


        internal Xl.Application XlApplication { get; private set; }

        public bool DisplayAlerts
        {
            get
            {
                var output = this.XlApplication.DisplayAlerts;
                return output;
            }
            set
            {
                this.XlApplication.DisplayAlerts = value;
            }
        }
        public bool FreezePanes
        {
            get
            {
                var output = this.XlApplication.ActiveWindow.FreezePanes;
                return output;
            }
            set
            {
                this.XlApplication.ActiveWindow.FreezePanes = value;
            }
        }
        public bool ScreenUpdating
        {
            get
            {
                var output = this.XlApplication.ScreenUpdating;
                return output;
            }
            set
            {
                this.XlApplication.ScreenUpdating = value;
            }
        }
        public bool Visible
        {
            get
            {
                var output = this.XlApplication.Visible;
                return output;
            }
            set
            {
                this.XlApplication.Visible = value;
            }
        }
        public int ZoomPercent
        {
            get
            {
                var output = (int)this.XlApplication.ActiveWindow.Zoom;
                return output;
            }
            set
            {
                this.XlApplication.ActiveWindow.Zoom = value;
            }
        }


        public Application(bool visible)
        {
            this.XlApplication = new Xl.Application()
            {
                Visible = visible,
            };
        }

        /// <summary>
        /// Uses the <see cref="IValues.ApplicationVisibility_Default"/> value.
        /// </summary>
        public Application()
            : this(Instances.Values.ApplicationVisibility_Default)
        {
        }

        /// <summary>
        /// Identical to <see cref="Application.Dispose()"/>, but allows for use outside of a using statment.
        /// </summary>
        public void Quit()
        {
            this.Dispose();
        }

        public Workbook NewWorkbook()
        {
            var xlWorkbook = this.XlApplication.Workbooks.Add();

            var workbook = new Workbook(xlWorkbook, this);
            return workbook;
        }

        public Workbook OpenWorkbook(string workbookFilePath)
        {
            var xlWorkbook = this.XlApplication.Workbooks.Open(workbookFilePath);

            var workbook = new Workbook(xlWorkbook, this);
            return workbook;
        }
    }
}
