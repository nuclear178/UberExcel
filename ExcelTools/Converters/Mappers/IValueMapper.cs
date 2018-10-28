using System;

namespace ExcelTools.Converters.Mappers
{
    public interface IValueMapper
    {
        object MapValue(Type type, object rawValue);
    }
}