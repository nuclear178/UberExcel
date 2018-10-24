namespace ExcelTools.IO
{
    public class RowObject
    {
        public RowObject(int columnIndex)
        {
            ColumnIndex = columnIndex;
        }

        public int ColumnIndex { get; }
    }
}