using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelTools.Introspection.Mapping;
using ExcelTools.Meta.Attributes;

namespace ExcelTools.Introspection
{
    public class AttributeTypeIntrospector : ITypeIntrospector
    {
        private readonly Type _objType;
        private readonly ObjectSchemaBuilder _mapping;

        public AttributeTypeIntrospector(Type objType)
        {
            _objType = objType;
            _mapping = new ObjectSchemaBuilder(objType);
        }

        public ObjectSchema Analyze()
        {
            Traverse(_objType);

            return _mapping.Build();
        }

        // ReSharper disable once SuggestBaseTypeForParameter
        private void Traverse(Type currentType, PropertyInfo parentObj = null)
        {
            GetColumns(currentType).ForEach(column =>
            {
                var columnOptions = column.GetCustomAttribute<ColumnAttribute>();
                _mapping.AddColumn(
                    columnIndex: columnOptions.ColumnIndex,
                    columnName: column.Name,
                    addedIndex: out int addedIndex,
                    parentName: parentObj?.Name
                );

                if (!Attribute.IsDefined(column, typeof(ConverterAttribute))) return;

                var converterOptions = column.GetCustomAttribute<ConverterAttribute>();
                _mapping.AddConverter(
                    columnIndex: addedIndex,
                    converterType: converterOptions.ConverterType
                );
            });

            GetIncludedColumns(currentType).ForEach(including =>
            {
                var includingOptions = including.GetCustomAttribute<IncludeAttribute>();
                _mapping.IncludeWithOffset(
                    offset: includingOptions.Offset,
                    includingName: including.Name,
                    parentName: parentObj?.Name
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