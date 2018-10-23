using System.Collections.Generic;
using System.Reflection;

namespace ExcelTools.IO
{
    public interface ITypeIntrospector
    {
        List<PropertyInfo> Analyze();
    }
}