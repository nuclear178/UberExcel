using System;

namespace ExcelTools.Converters
{
    public interface ITypeConverterFactory
    {
        IConverter Create(Type propType, Type converterType = null);
    }
}