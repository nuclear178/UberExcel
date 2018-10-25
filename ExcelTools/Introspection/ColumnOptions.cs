using System;

namespace ExcelTools.Introspection
{
    public class ColumnOptions
    {
        public ColumnOptions(int index, string fullName, Type type)
        {
            Index = index;
            FullName = fullName;
            Type = type;
        }

        public int Index { get; }
        public string FullName { get; }
        public Type Type { get; }
    }
}