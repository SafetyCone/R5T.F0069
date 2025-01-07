using System;

using R5T.T0132;


namespace R5T.F0069
{
    [FunctionalityMarker]
    public partial interface IRangeSizeOperator : IFunctionalityMarker
    {
        public int Get_RowCount(object[,] values)
        {
            var output = values.GetLength(0);
            return output;
        }

        public int Get_ColumnCount(object[,] values)
        {
            var output = values.GetLength(1);
            return output;
        }

        public RangeSize Get_Size(object[,] values)
        {
            var rows = this.Get_RowCount(values);
            var columns = this.Get_ColumnCount(values);

            var output = new RangeSize
            {
                Rows = rows,
                Columns = columns,
            };

            return output;
        }
    }
}
