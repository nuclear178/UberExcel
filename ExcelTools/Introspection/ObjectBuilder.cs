using System;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;

namespace ExcelTools.Introspection
{
    public class ObjectBuilder<T> where T : new()
    {
        private readonly ObjectSchema _schema;

        public ObjectBuilder(ObjectSchema schema)
        {
            _schema = schema;
        }

        public T Build(ExcelRange cells, int rowIndex)
        {
            var obj = new T();
            foreach (ColumnOptions column in _schema)
            {
                object rawValue = cells[rowIndex, column.Index].Value;
                
            }

            return obj;
        }

        public static void SetPropValue(string qualifiedName, object obj, object value)
        {
            var parts = qualifiedName.Split('.');
            for (var i = 0; i < parts.Length - 1; i++)
            {
                PropertyInfo property = obj.GetType().GetProperty(parts[i]);

                object propValue = property.GetValue(obj, null);
                if (propValue != null)
                {
                    obj = propValue;
                    continue;
                }

                object propObj = Activator.CreateInstance(property.PropertyType);
                property.SetValue(obj, propObj, null);

                obj = propObj;
            }

            obj.GetType()
                .GetProperty(parts.Last())
                .SetValue(obj, value, null);
        }

        private double ParseDouble(string value)
        {
            return double.Parse(value);
        }

        private static int ParseInt(string value)
        {
            return int.Parse(value);
        }
    }
}