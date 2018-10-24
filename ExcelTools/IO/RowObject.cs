namespace ExcelTools.IO
{
    public class RowObject
    {
        public RowObject(int columnIndex, string name)
        {
            ColumnIndex = columnIndex;
            Name = name;
        }

        public int ColumnIndex { get; }
        public string Name { get; }
    }
}