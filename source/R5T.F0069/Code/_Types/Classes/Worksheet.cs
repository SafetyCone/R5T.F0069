using System;

using R5T.T0142;

using Xl = Microsoft.Office.Interop.Excel;


namespace R5T.F0069
{
    [UtilityTypeMarker]
    public class Worksheet
    {
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
                string output = this.XlWorksheet.Name;
                return output;
            }
            set
            {
                this.XlWorksheet.Name = value;
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
    }
}
