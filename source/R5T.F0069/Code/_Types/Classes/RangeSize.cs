using System;


namespace R5T.F0069
{
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
