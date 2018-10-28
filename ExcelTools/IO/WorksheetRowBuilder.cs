using System;
using System.Reflection;
using ExcelTools.Introspection.Mapping;
using OfficeOpenXml;

namespace ExcelTools.IO
{
    public class WorksheetRowBuilder
    {
        private readonly ObjectSchema _schema;

        public WorksheetRowBuilder(ObjectSchema schema)
        {
            _schema = schema;
        }

        public void Build(ExcelRange cells, int rowIndex, object obj)
        {
            foreach (ColumnOptions column in _schema)
            {
                object rawValue = GetPropValue(column.FullName, obj);
                cells[rowIndex, column.Index].Value = column.MapValueTo(rawValue);
            }
        }

        private static object GetPropValue(string qualifiedName, object obj)
        {
            foreach (string part in qualifiedName.Split('.'))
            {
                if (obj == null) return null;

                Type type = obj.GetType();
                PropertyInfo property = type.GetProperty(part);
                if (property == null) return null;

                obj = property.GetValue(obj, null);
            }

            return obj;
        }
    }
}