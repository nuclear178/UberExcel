using ExcelTools.Introspection.Mapping;

namespace ExcelTools.Introspection
{
    public interface ITypeIntrospector
    {
        ObjectSchema Analyze();
    }
}