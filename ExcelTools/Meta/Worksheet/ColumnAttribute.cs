using System;

namespace ExcelTools.Meta.Worksheet
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ColumnAttribute : Attribute
    {
        public int ColumnIndex { get; }
        public string HeaderName { get; set; }

        public ColumnAttribute(int columnIndex)
        {
            ColumnIndex = columnIndex;
        }
    }
}