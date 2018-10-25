using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelTools.Meta.Worksheet;

namespace ExcelTools.Introspection
{
    public class AttributeBasedIntrospector : ITypeIntrospector
    {
        private readonly Type _rootType;
        private readonly ObjectSchemaBuilder _mapping;

        public AttributeBasedIntrospector(Type rootType)
        {
            _rootType = rootType;
            _mapping = new ObjectSchemaBuilder();
        }

        public ObjectSchema Analyze()
        {
            Traverse(_rootType);

            return _mapping.Build();
        }

        // ReSharper disable once SuggestBaseTypeForParameter
        private void Traverse(Type currentType, PropertyInfo parentObj = null)
        {
            GetColumns(currentType).ForEach(currentColumn =>
            {
                var columnOptions = currentColumn.GetCustomAttribute<Column>();
                _mapping.AddColumn(
                    columnIndex: columnOptions.ColumnIndex,
                    columnName: currentColumn.Name,
                    parentName: parentObj?.Name);
            });

            GetIncludedColumns(currentType).ForEach(included =>
            {
                var includingOptions = included.GetCustomAttribute<Include>();
                _mapping.Include(
                    offset: includingOptions.Offset,
                    name: included.Name,
                    parentName: parentObj?.Name);

                Traverse(currentType: included.PropertyType, parentObj: included);
            });
        }

        private static List<PropertyInfo> GetColumns(Type rowType)
        {
            return rowType.GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(Column)))
                .ToList();
        }

        private static List<PropertyInfo> GetIncludedColumns(Type parentType)
        {
            return parentType.GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(Include)))
                .ToList();
        }
    }
}