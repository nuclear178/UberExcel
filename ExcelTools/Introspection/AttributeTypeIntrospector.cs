using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelTools.Converters;
using ExcelTools.Introspection.Mapping;
using ExcelTools.Meta.Attributes;

namespace ExcelTools.Introspection
{
    public class AttributeTypeIntrospector : TypeIntrospectorBase
    {
        public AttributeTypeIntrospector(Type objType, ITypeConverterFactory converterFactory)
            : base(objType, converterFactory)
        {
        }

        public override ObjectSchema Analyze()
        {
            Traverse(ObjType);

            return Mapping.Build();
        }

        // ReSharper disable once SuggestBaseTypeForParameter
        private void Traverse(Type currentType, PropertyInfo parentObj = null)
        {
            GetColumns(currentType).ForEach(column =>
            {
                var columnOptions = column.GetCustomAttribute<ColumnAttribute>();
                Mapping.AddColumn(
                    columnIndex: columnOptions.ColumnIndex,
                    columnName: column.Name,
                    addedIndex: out int addedIndex,
                    parentName: parentObj?.Name
                );

                if (!Attribute.IsDefined(column, typeof(ConverterAttribute))) return;

                var converterOptions = column.GetCustomAttribute<ConverterAttribute>();
                Mapping.AddConverter(
                    columnIndex: addedIndex,
                    converterType: converterOptions.ConverterType
                );
            });

            GetIncludedColumns(currentType).ForEach(including =>
            {
                var includingOptions = including.GetCustomAttribute<IncludeAttribute>();
                Mapping.Include(
                    includingName: including.Name,
                    parentName: parentObj?.Name,
                    offset: includingOptions.Offset
                );

                Traverse(currentType: including.PropertyType, parentObj: including);
            });
        }

        private static List<PropertyInfo> GetColumns(Type rowType)
        {
            return rowType.GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(ColumnAttribute)))
                .ToList();
        }

        private static List<PropertyInfo> GetIncludedColumns(Type parentType)
        {
            return parentType.GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(IncludeAttribute)))
                .ToList();
        }
    }
}