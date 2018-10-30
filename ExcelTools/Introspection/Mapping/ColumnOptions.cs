using System;
using System.Linq;
using System.Reflection;
using ExcelTools.Converters;
using ExcelTools.Converters.Mappers;

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

        public object MapValueFrom(object rawValue, IValueMapper mapper)
        {
            return HasCustomConverter
                ? GetCustomConverter().Read(rawValue.ToString())
                : mapper.MapValue(_propType, rawValue);
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