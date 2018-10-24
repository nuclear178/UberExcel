using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelTools.Exceptions;

namespace ExcelTools.IO
{
    public class ObjectSchemaBuilder
    {
        private readonly Dictionary<int, string> _columns;
        private readonly List<string> _includings;

        public ObjectSchemaBuilder()
        {
            _columns = new Dictionary<int, string>();
            _includings = new List<string>();
        }

        public void Add(int columnIndex, string columnName, string parentName = null)
        {
            if (_columns.ContainsKey(columnIndex))
                throw ExcelWorksheetMapperException.ColumnIndexAlreadyExists(columnIndex,
                    alreadyContainedName: _columns[columnIndex]);

            string name = parentName == null ? columnName : IncludeName(parentName, columnName);

            _columns[columnIndex] = name;
        }

        public void Include(string includedName, string parentName = null)
        {
            if (parentName == null)
            {
                _includings.Add(includedName);
            }
            else
            {
                for (var i = 0; i < _includings.Count; i++)
                {
                    if (_includings[i].EndsWith(parentName))
                    {
                        _includings[i] = $"{_includings[i]}.{includedName}";
                    }
                }
            }
        }

        public ObjectSchema Build()
        {
            /*Console.WriteLine("::");
            foreach (string columnsValue in _columns.Values)
            {
                Console.WriteLine(columnsValue);
            }

            Console.WriteLine("::");
            _includings.ForEach(Console.WriteLine);*/

            return new ObjectSchema(new List<PropertyInfo>());
        }

        private string IncludeName(string parentName, string columnName)
        {
            string oldName = _includings.SingleOrDefault(include => include.EndsWith(parentName));
            return $"{oldName}.{columnName}";
        }
    }
}