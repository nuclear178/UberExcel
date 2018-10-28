using System;
using ExcelTools.Exceptions;

namespace ExcelTools.Converters.Mappers
{
    public class DefaultValueMapper : IValueMapper
    {
        public object MapValue(Type type, object rawValue)
        {
            if (type == typeof(string))
                return rawValue.ToString();

            if (type == typeof(short))
                return Convert.ToInt16(rawValue);

            if (type == typeof(int))
                return Convert.ToInt32(rawValue);

            if (type == typeof(long))
                return Convert.ToInt64(rawValue);

            if (type == typeof(double))
                return Convert.ToDouble(rawValue);

            if (type == typeof(decimal))
                return Convert.ToDecimal(rawValue);

            throw ExcelWorksheetMapperException.UnsupportedColumnType(typeName: type.FullName);
        }
    }
}