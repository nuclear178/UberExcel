using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelTools.Exceptions;

namespace ExcelTools.IO
{
    public class ObjectSchemaBuilder
    {
        private readonly Dictionary<int, string> _columns;

        public ObjectSchemaBuilder()
        {
            _columns = new Dictionary<int, string>();
        }

        public void Add(int columnIndex, string columnName, string parentName = null)
        {
            if (_columns.ContainsKey(columnIndex))
                throw ExcelWorksheetMapperException.ColumnIndexAlreadyExists(columnIndex,
                    alreadyContainedName: _columns[columnIndex]);

            string name = parentName == null ? columnName : $"{parentName}.{columnName}";

            _columns[columnIndex] = name;
        }

        public ObjectSchema Build()
        {
            foreach (string columnsValue in _columns.Values)
            {
                Console.WriteLine(columnsValue);
            }

            return new ObjectSchema(new List<PropertyInfo>());
        }
    }
}