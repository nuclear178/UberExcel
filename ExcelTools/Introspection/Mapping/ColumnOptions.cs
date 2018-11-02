using System;
using System.Linq;
using System.Reflection;
using ExcelTools.Converters;

namespace ExcelTools.Introspection.Mapping
{
    public class ColumnOptions
    {
        public int Index { get; }

        private readonly string _fullName;
        private readonly IConverter _converter;

        public ColumnOptions(int index, string fullName, IConverter converter)
        {
            Index = index;
            _fullName = fullName;
            _converter = converter;
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

        public object MapFrom(object rawValue)
        {
            return _converter.Read(rawValue.ToString());
        }

        public object MapTo(object rawValue)
        {
            return _converter.Write(rawValue);
        }
    }
}