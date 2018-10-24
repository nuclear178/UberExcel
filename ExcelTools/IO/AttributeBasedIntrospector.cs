using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelTools.Meta.Worksheet;

namespace ExcelTools.IO
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
                var columnMeta = currentColumn.GetCustomAttribute<Column>();
                _mapping.Add(
                    columnIndex: columnMeta.ColumnIndex,
                    columnName: currentColumn.Name,
                    parentName: parentObj?.Name);
            });
            GetNestedColumns(currentType).ForEach(included =>
            {
                _mapping.Include(included.Name, parentObj?.Name);
                Traverse(currentType: included.PropertyType, parentObj: included);
            });
        }

        private static List<PropertyInfo> GetColumns(Type rowType)
        {
            return rowType.GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(Column)))
                .ToList();
        }

        private static List<PropertyInfo> GetNestedColumns(Type parentType)
        {
            return parentType.GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(Include)))
                .ToList();
        }
    }
}