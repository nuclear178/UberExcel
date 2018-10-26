using System;

namespace ExcelTools.Meta.Worksheet
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ColumnAttribute : Attribute
    {
        public int ColumnIndex { get; set; }
        public string HeaderName { get; set; }

        public ColumnAttribute(int columnIndex = -1)
        {
            ColumnIndex = columnIndex;
        }
    }
}