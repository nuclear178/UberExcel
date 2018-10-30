using System;
using System.Linq;
using System.Reflection;
using ExcelTools.Converters;
using ExcelTools.Exceptions;

namespace ExcelTools.Introspection.Mapping
{
    public class ColumnOptions
    {
        private readonly string _fullName;
        private readonly Type _propType;
        private readonly Type _converterType;

        public ColumnOptions(int index, string fullName, Type propType, Type converterType)
        {
            Index = index;
            _fullName = fullName;
            _propType = propType;
            _converterType = converterType;
        }

        public int Index { get; }

        private bool HasCustomConverter => _converterType != null;

        public object GetValue(object obj)
        {
            foreach (string piece in _fullName.Split('.'))
            {
                if (obj == null) return null;

                Type type = obj.GetType();
                PropertyInfo property = type.GetProperty(piece);
                if (property == null) return null;

                obj = property.GetValue(obj, null);
            }

            return obj;
        }

        public void SetValue(object obj, object value)
        {
            string[] parts = _fullName.Split('.');
            for (var i = 0; i < parts.Length - 1; i++)
            {
                PropertyInfo property = obj.GetType().GetProperty(parts[i]);

                object propValue = property.GetValue(obj, null);
                if (propValue != null)
                {
                    obj = propValue;
                    continue;
                }

                object propObj = Activator.CreateInstance(property.PropertyType);
                property.SetValue(obj, propObj, null);

                obj = propObj;
            }

            obj.GetType()
                .GetProperty(parts.Last())
                .SetValue(obj, value, null);
        }

        public object MapValueFrom(object rawValue)
        {
            if (HasCustomConverter)
                return GetCustomConverter().Read(rawValue.ToString());

            if (_propType == typeof(string))
                return rawValue.ToString();

            if (_propType == typeof(short))
                return Convert.ToInt16(rawValue);

            if (_propType == typeof(int))
                return Convert.ToInt32(rawValue);

            if (_propType == typeof(long))
                return Convert.ToInt64(rawValue);

            if (_propType == typeof(double))
                return Convert.ToDouble(rawValue);

            if (_propType == typeof(decimal))
                return Convert.ToDecimal(rawValue);

            throw ExcelWorksheetMapperException.UnsupportedColumnType(typeName: _propType.FullName);
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