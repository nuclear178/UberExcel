using System;
using System.Reflection;
using OfficeOpenXml;

namespace ExcelTools.Introspection
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
                cells[rowIndex, column.Index].Value = rawValue;
            }
        }

        private static object GetPropValue(string qualifiedName, object obj)
        {
            foreach (string part in qualifiedName.Split('.'))
            {
                if (obj == null) return null;

                Type type = obj.GetType();
                PropertyInfo info = type.GetProperty(part);
                if (info == null) return null;

                obj = info.GetValue(obj, null);
            }

            return obj;
        }
    }
}