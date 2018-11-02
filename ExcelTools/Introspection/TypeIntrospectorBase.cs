using System;
using ExcelTools.Converters;
using ExcelTools.Introspection.Mapping;

namespace ExcelTools.Introspection
{
    public abstract class TypeIntrospectorBase : ITypeIntrospector
    {
        protected readonly Type ObjType;
        protected readonly ObjectSchemaBuilder Mapping;

        protected TypeIntrospectorBase(Type objType, ITypeConverterFactory converterFactory)
        {
            ObjType = objType;
            Mapping = new ObjectSchemaBuilder(objType, converterFactory);
        }

        public abstract ObjectSchema Analyze();
    }
}