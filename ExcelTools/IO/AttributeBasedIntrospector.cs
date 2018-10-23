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
        private readonly List<PropertyInfo> _columns;

        public AttributeBasedIntrospector(Type rootType)
        {
            _rootType = rootType;
            _columns = new List<PropertyInfo>();
        }

        public List<PropertyInfo> Analyze()
        {
            Traverse(_rootType);
            return _columns;
        }

        private void Traverse(Type currentType)
        {
            _columns.AddRange(GetColumns(currentType));
            GetNestedColumnsTypes(currentType).ForEach(Traverse);
        }

        private static IEnumerable<PropertyInfo> GetColumns(Type rowType)
        {
            return rowType.GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(Column)));
        }

        private static List<Type> GetNestedColumnsTypes(Type parentType)
        {
            return parentType.GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(Include)))
                .Select(included => included.PropertyType)
                .ToList();
        }
    }
}