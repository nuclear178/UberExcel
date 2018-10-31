using System;
using System.Linq;
using System.Reflection;
using ExcelTools.Converters;
using ExcelTools.Converters.Mappers;

namespace ExcelTools.Introspection.Mapping
{
    public class ColumnOptions
    {
        public int Index { get; }

        private readonly string _fullName;
        private readonly Type _propType;

        private IConverter Converter { get; set; }
        private bool HasCustomConverter => Converter != null;

        public ColumnOptions(int index, string fullName, Type propType, Type converterType)
        {
            Index = index;
            _fullName = fullName;
            _propType = propType;
            InitializeCustomConverter(converterType);
        }

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

        public object MapFrom(object rawValue, IValueMapper mapper)
        {
            return HasCustomConverter
                ? Converter.Read(rawValue.ToString())
                : mapper.MapValue(_propType, rawValue);
        }

        public object MapTo(object rawValue)
        {
            return HasCustomConverter
                ? Converter.Write(rawValue)
                : rawValue;
        }

        private void InitializeCustomConverter(Type converterType)
        {
            if (_propType.IsEnum)
            {
                Converter = new EnumConverter(_propType);
            }
            else
            {
                if (converterType == null)
                    return;

                Converter = (IConverter) Activator.CreateInstance(converterType);
            }
        }
    }
}