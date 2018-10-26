using System;
using ExcelTools.Converters;
using ExcelTools.Exceptions;

namespace ExcelTools.Meta.Worksheet
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ConverterAttribute : Attribute
    {
        public Type ConverterType { get; }

        public ConverterAttribute(Type converterType)
        {
            if (converterType == null)
                throw new ArgumentNullException(nameof(converterType));

            if (!typeof(IConverter).IsAssignableFrom(converterType))
                throw ExcelWorksheetMapperException.InvalidConverterType(converterType.FullName);

            ConverterType = converterType;
        }
    }
}