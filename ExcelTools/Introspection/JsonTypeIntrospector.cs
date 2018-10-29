using System;
using System.Collections.Generic;
using System.Linq;
using ExcelTools.Introspection.Mapping;
using Newtonsoft.Json.Linq;

namespace ExcelTools.Introspection
{
    public class JsonTypeIntrospector : ITypeIntrospector
    {
        private readonly ObjectSchemaBuilder _mapping;
        private readonly JObject _mappingJson;

        public JsonTypeIntrospector(Type objType, JObject mappingJson)
        {
            _mappingJson = mappingJson;
            _mapping = new ObjectSchemaBuilder(objType);
        }

        public ObjectSchema Analyze()
        {
            Traverse(_mappingJson);

            return _mapping.Build();
        }

        private void Traverse(JToken currentObj, string parentName = null)
        {
            GetColumns(currentObj).ForEach(column =>
            {
                _mapping.AddColumn(
                    columnIndex: column["columnIndex"].ToObject<int>(),
                    columnName: column["propertyName"].ToString(),
                    columnType: null, // todo Move type interference into column obj...
                    addedIndex: out int _,
                    parentName: parentName
                );
            });

            GetIncludedColumns(currentObj).ForEach(including =>
            {
                _mapping.Include(
                    includingName: including["propertyName"].ToString(),
                    parentName: parentName
                );

                Traverse(including, including["propertyName"].ToString());
            });
        }

        private static List<JToken> GetColumns(JToken currentObj)
        {
            return currentObj["columns"].Children().ToList();
        }

        private static List<JToken> GetIncludedColumns(JToken parentObj)
        {
            return parentObj["includings"] == null
                ? new List<JToken>()
                : parentObj["includings"].Children().ToList();
        }
    }
}