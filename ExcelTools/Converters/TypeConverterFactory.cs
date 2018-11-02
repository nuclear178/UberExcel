using System;

namespace ExcelTools.Converters
{
    public class TypeConverterFactory : ITypeConverterFactory
    {
        public IConverter Create(Type propType, Type converterType = null)
        {
            if (propType.IsEnum)
            {
                return new EnumConverter(propType);
            }

            if (converterType == null)
            {
                return new PrimitiveTypeConverter(
                    type: propType,
                    nullable: Nullable.GetUnderlyingType(propType) != null
                );
            }

            return (IConverter) Activator.CreateInstance(converterType);
        }
    }
}