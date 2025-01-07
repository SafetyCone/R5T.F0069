using System;

using R5T.T0142;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069
{
    /// <summary>
    /// Wraps an Excel worksheet.
    /// </summary>
    /// <remarks>
    /// Implements value-based semantics.
    /// </remarks>
    [UtilityTypeMarker]
    public sealed class Worksheet : IDisposable, IEquatable<Worksheet>
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
                    /// The <see cref="Xl.Application"/> object itself is managed, the Excel application it is the handle to is not.
                }

                Instances.MarshalOperator.Release_ComObject(this.XlWorksheet);

                this.XlWorksheet = null;
            }

            this.zDisposed = true;
        }

        ~Worksheet()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            this.Dispose(false);
        }

        #endregion

        /// <summary>
        /// The underlying Excel COM automation worksheet object.
        /// </summary>
        /// <remarks>
        /// Note: will never be null.
        /// See <see cref="IWorksheetOperator.Is_Valid(Worksheet)"/>.
        /// </remarks>
        internal Xl.Worksheet XlWorksheet { get; private set; }

        public Workbook Workbook { get; private set; }

        public Application Application
        {
            get
            {
                var output = this.Workbook.Application;
                return output;
            }
        }
        public string Name
        {
            get
            {
                string output = Instances.WorksheetOperator.Get_Name(this);
                return output;
            }
            set
            {
                Instances.WorksheetOperator.Set_Name(this, value);
            }
        }


        internal Worksheet(Xl.Worksheet xlWorksheet, Workbook workbook)
        {
            this.XlWorksheet = xlWorksheet;
            this.Workbook = workbook;
        }

        internal Worksheet(Xl.Worksheet xlWorksheet, Workbook workbook, string name)
            : this(xlWorksheet, workbook)
        {
            this.Name = name;
        }

        public void Delete()
        {
            this.XlWorksheet.Delete();
        }

        public void Select()
        {
            this.Workbook.Select(); // Make sure this worksheet's workbook is selected first.

            this.XlWorksheet.Activate();
        }

        public Range GetA1Range()
        {
            var xlRange = this.XlWorksheet.Cells[1, 1] as Xl.Range;

            var range = new Range(xlRange, this);
            return range;
        }

        public Range GetRange(Range upperLeft, int numberOfRows, int numberOfColumns)
        {
            var xlLowerRight = this.XlWorksheet.Cells[upperLeft.Row + numberOfRows - 1, upperLeft.Column + numberOfColumns - 1] as Xl.Range;

            var xlRange = this.XlWorksheet.Range[upperLeft.XlRange, xlLowerRight];

            var range = new Range(xlRange, this);
            return range;
        }

        public bool Equals(Worksheet other)
        {
            var output = Instances.WorksheetOperator.Equals_WithNullCheck(
                this,
                other);

            return output;
        }

        public override bool Equals(object obj)
            => this.Equals(obj as Worksheet);

        public override int GetHashCode()
        {
            var output = Instances.WorksheetOperator.Get_HashCode(this);
            return output;
        }

        public override string ToString()
        {
            var output = Instances.WorksheetOperator.Get_Name(this);
            return output;
        }
    }
}
