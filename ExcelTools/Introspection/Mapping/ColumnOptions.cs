using System;
using ExcelTools.Converters;
using ExcelTools.Exceptions;

namespace ExcelTools.Introspection.Mapping
{
    public class ColumnOptions
    {
        private readonly Type _converterType;

        public ColumnOptions(int index, string fullName, Type converterType)
        {
            Index = index;
            FullName = fullName;
            _converterType = converterType;
        }

        public int Index { get; }
        public string FullName { get; }
        private Type Type { get; } // todo Interfere
        private bool HasCustomConverter => _converterType != null;

        public object MapValueFrom(object rawValue)
        {
            if (HasCustomConverter)
                return GetCustomConverter().Read(rawValue.ToString());

            if (Type == typeof(string))
                return rawValue.ToString();

            if (Type == typeof(short))
                return Convert.ToInt16(rawValue);

            if (Type == typeof(int))
                return Convert.ToInt32(rawValue);

            if (Type == typeof(long))
                return Convert.ToInt64(rawValue);

            if (Type == typeof(double))
                return Convert.ToDouble(rawValue);

            if (Type == typeof(decimal))
                return Convert.ToDecimal(rawValue);

            throw ExcelWorksheetMapperException.UnsupportedColumnType(typeName: Type.FullName);
        }

        public object MapValueTo(object rawValue)
        {
            return HasCustomConverter
                ? GetCustomConverter().Write(rawValue)
                : rawValue;
        }

        private IConverter GetCustomConverter()
        {
            if (!HasCustomConverter)
                throw new Exception("Custom converter is not defined.");

            return (IConverter) Activator.CreateInstance(_converterType);
        }
    }
}