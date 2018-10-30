using System;
using System.Collections.Generic;
using System.Linq;
using ExcelTools.Introspection.Mapping;
using ExcelTools.Meta.Json;
using Newtonsoft.Json.Linq;

namespace ExcelTools.Introspection
{
    public class JsonTypeIntrospector : ITypeIntrospector
    {
        private readonly ObjectSchemaBuilder _mapping;
        private readonly JToken _mappingJson;

        public JsonTypeIntrospector(Type objType, JToken mappingJson)
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
                    columnIndex: column[JsonPropertiesTokens.ColumnIndex].ToObject<int>(),
                    columnName: column[JsonPropertiesTokens.PropertyName].ToString(),
                    addedIndex: out int _,
                    parentName: parentName
                );

                if (column[JsonPropertiesTokens.ConverterClass] != null)
                {
                    _mapping.AddConverter(
                        columnIndex: column[JsonPropertiesTokens.ColumnIndex].ToObject<int>(),
                        converterType: Type.GetType(column[JsonPropertiesTokens.ConverterClass].ToString())
                    );
                }
            });

            GetIncludedColumns(currentObj).ForEach(including =>
            {
                _mapping.Include(
                    includingName: including[JsonPropertiesTokens.PropertyName].ToString(),
                    parentName: parentName
                );

                Traverse(including, including[JsonPropertiesTokens.PropertyName].ToString());
            });
        }

        private static List<JToken> GetColumns(JToken currentObj)
        {
            return currentObj[JsonPropertiesTokens.Columns] == null
                ? new List<JToken>()
                : currentObj[JsonPropertiesTokens.Columns].Children().ToList();
        }

        private static List<JToken> GetIncludedColumns(JToken parentObj)
        {
            return parentObj[JsonPropertiesTokens.Includings] == null
                ? new List<JToken>()
                : parentObj[JsonPropertiesTokens.Includings].Children().ToList();
        }
    }
}