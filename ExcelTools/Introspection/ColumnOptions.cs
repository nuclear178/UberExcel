namespace ExcelTools.Introspection
{
    public class ColumnOptions
    {
        public ColumnOptions(int index, string fullName)
        {
            Index = index;
            FullName = fullName;
        }

        public int Index { get; }
        public string FullName { get; }
    }
}