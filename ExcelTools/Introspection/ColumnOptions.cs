using System;
using ExcelTools.Exceptions;

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
        private Type Type { get; }

        public object MapValue(object rawValue)
        {
            if (Type == typeof(string))
            {
                return rawValue.ToString();
            }

            if (Type == typeof(short))
            {
                return Convert.ToInt16(rawValue);
            }

            if (Type == typeof(int))
            {
                return Convert.ToInt32(rawValue);
            }

            if (Type == typeof(long))
            {
                return Convert.ToInt64(rawValue);
            }

            if (Type == typeof(double))
            {
                return Convert.ToDouble(rawValue);
            }

            if (Type == typeof(decimal))
            {
                return Convert.ToDecimal(rawValue);
            }

            throw ExcelWorksheetMapperException.UnsupportedColumnType(typeName: Type.FullName);
        }
    }
}