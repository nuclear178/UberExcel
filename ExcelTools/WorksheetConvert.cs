using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelTools.Meta.Worksheet;

namespace ExcelTools
{
    public static class WorksheetConvert
    {
        public static string SerializeObject<T>(List<T> value)
        {
            var rowType = typeof(T);

            var columns = GetColumns(rowType);
            while (true)
            {
                var nestedColumns = GetNestedColumns(rowType);
                if (nestedColumns.Any())
                {
                }
            }
        }

        public static IEnumerable<T> DeserializeObject<T>()
        {
            return new List<T>();
        }

        private static List<PropertyInfo> GetColumns(Type rowType)
        {
            return rowType.GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(Column)))
                .Where(prop => prop.GetType().IsValueType)
                .ToList();
        }

        private static List<PropertyInfo> GetNestedColumns(Type rowType)
        {
            return rowType.GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(Include)))
                .Where(prop => prop.GetType().IsClass)
                .ToList();
        }
    }
}