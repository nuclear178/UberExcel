using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelTools.Introspection.Mapping;
using ExcelTools.Meta.Worksheet;

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
            GetColumns(currentType).ForEach(currentColumn =>
            {
                var columnOptions = currentColumn.GetCustomAttribute<ColumnAttribute>();
                _mapping.AddColumn(
                    columnIndex: columnOptions.ColumnIndex,
                    columnName: currentColumn.Name,
                    columnType: currentColumn.PropertyType,
                    addedIndex: out int addedIndex,
                    parentName: parentObj?.Name
                );

                if (!Attribute.IsDefined(currentColumn, typeof(ConverterAttribute))) return;

                var converterOptions = currentColumn.GetCustomAttribute<ConverterAttribute>();
                _mapping.AddConverter(
                    columnIndex: addedIndex,
                    converterType: converterOptions.ConverterType
                );
            });

            GetIncludedColumns(currentType).ForEach(included =>
            {
                var includingOptions = included.GetCustomAttribute<IncludeAttribute>();
                _mapping.Include(
                    offset: includingOptions.Offset,
                    name: included.Name,
                    parentName: parentObj?.Name
                );

                Traverse(currentType: included.PropertyType, parentObj: included);
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