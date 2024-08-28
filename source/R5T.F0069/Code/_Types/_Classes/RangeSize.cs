using System;

using R5T.T0142;


namespace R5T.F0069
{
    [UtilityTypeMarker]
    public class RangeSize
    {
        public int Rows { get; set; }
        public int Columns { get; set; }


        public RangeSize()
        {
        }

        public RangeSize(int rows, int columns)
        {
            this.Rows = rows;
            this.Columns = columns;
        }
    }
}
