using System;
using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace ExcelTools.Converters
{
    public class EnumConverter : IConverter
    {
        private readonly Type _enumType;

        public EnumConverter(Type enumType)
        {
            if (!enumType.IsEnum)
                throw new InvalidOperationException();

            _enumType = enumType;
        }

        public string Write(object value)
        {
            if (!value.GetType().IsEnum)
                throw new InvalidOperationException();

            FieldInfo field = _enumType.GetField(value.ToString());

            return field?.GetCustomAttribute<DisplayAttribute>()?.Name;
        }

        public object Read(string value)
        {
            foreach (FieldInfo field in _enumType.GetFields())
            {
                if (Attribute.GetCustomAttribute(field, typeof(DisplayAttribute)) is DisplayAttribute attribute)
                {
                    if (attribute.Name == value)
                    {
                        return field.GetValue(null);
                    }
                }
                else
                {
                    if (field.Name == value)
                    {
                        return field.GetValue(null);
                    }
                }
            }

            throw new ArgumentOutOfRangeException(nameof(value));
        }
    }
}