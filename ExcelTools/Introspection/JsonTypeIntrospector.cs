using System;
using System.Collections.Generic;
using System.Linq;
using ExcelTools.Converters;
using ExcelTools.Introspection.Mapping;
using ExcelTools.Meta.Json;
using Newtonsoft.Json.Linq;

namespace ExcelTools.Introspection
{
    public class JsonTypeIntrospector : TypeIntrospectorBase
    {
        private readonly JToken _mappingJson;

        public JsonTypeIntrospector(Type objType, ITypeConverterFactory converterFactory, JToken mappingJson)
            : base(objType, converterFactory)
        {
            _mappingJson = mappingJson;
        }

        public override ObjectSchema Analyze()
        {
            Traverse(_mappingJson);

            return Mapping.Build();
        }

        private void Traverse(JToken currentObj, string parentName = null)
        {
            GetColumns(currentObj).ForEach(column =>
            {
                Mapping.AddColumn(
                    columnIndex: column[JsonPropertiesTokens.ColumnIndex].ToObject<int>(),
                    columnName: column[JsonPropertiesTokens.PropertyName].ToString(),
                    addedIndex: out int _,
                    parentName: parentName
                );

                if (column[JsonPropertiesTokens.ConverterClass] != null)
                {
                    Mapping.AddConverter(
                        columnIndex: column[JsonPropertiesTokens.ColumnIndex].ToObject<int>(),
                        converterType: Type.GetType(column[JsonPropertiesTokens.ConverterClass].ToString())
                    );
                }
            });

            GetIncludedColumns(currentObj).ForEach(including =>
            {
                Mapping.Include(
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