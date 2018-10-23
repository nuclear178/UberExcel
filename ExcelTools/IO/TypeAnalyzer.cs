using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelTools.Meta.Worksheet;

namespace ExcelTools.IO
{
    public class TypeAnalyzer
    {
        private readonly Type _rootType;
        private readonly List<PropertyInfo> _columns;

        public TypeAnalyzer(Type rootType)
        {
            _rootType = rootType;
            _columns = new List<PropertyInfo>();
        }

        public List<PropertyInfo> Analyze()
        {
            Traverse(_rootType);
            return _columns;
        }

        private void Traverse(Type rowType)
        {
            _columns.AddRange(GetColumns(rowType));
            GetNestedColumns(rowType)
                .Select(prop => prop.PropertyType)
                .ToList()
                .ForEach(Traverse);
        }

        private static IEnumerable<PropertyInfo> GetColumns(Type rowType)
        {
            return rowType.GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(Column)));
        }

        private static IEnumerable<PropertyInfo> GetNestedColumns(Type rowType)
        {
            return rowType.GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(Include)));
        }
    }
}