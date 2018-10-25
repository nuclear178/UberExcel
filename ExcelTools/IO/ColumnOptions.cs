namespace ExcelTools.IO
{
    public class ColumnOptions
    {
        public ColumnOptions(int columnIndex, string name)
        {
            ColumnIndex = columnIndex;
            Name = name;
        }

        public int ColumnIndex { get; }
        public string Name { get; }
    }
}