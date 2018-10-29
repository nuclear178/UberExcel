using System;
using ExcelTools.Introspection.Mapping;
using Newtonsoft.Json.Linq;

namespace ExcelTools.Introspection
{
    public class JsonTypeIntrospector : ITypeIntrospector
    {
        private readonly JObject _mappingJson;
        private readonly ObjectSchemaBuilder _mapping;

        public JsonTypeIntrospector(Type objType, JObject mappingJson)
        {
            _mappingJson = mappingJson;
            _mapping = new ObjectSchemaBuilder(objType);
        }

        public ObjectSchema Analyze()
        {
            return _mapping.Build();
        }
    }
}