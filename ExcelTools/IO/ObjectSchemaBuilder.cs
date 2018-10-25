using System.Collections.Generic;
using System.Linq;
using ExcelTools.Exceptions;

namespace ExcelTools.IO
{
    public class ObjectSchemaBuilder
    {
        private readonly Dictionary<int, string> _columns;
        private readonly Dictionary<string, Including> _includings;

        public ObjectSchemaBuilder()
        {
            _columns = new Dictionary<int, string>();
            _includings = new Dictionary<string, Including>();
        }

        public ObjectSchema Build()
        {
            return new ObjectSchema(
                _columns.Select(pair => new ColumnOptions(pair.Key, pair.Value))
            );
        }

        public void AddColumn(int columnIndex, string columnName, string parentName = null)
        {
            columnIndex = parentName == null ? columnIndex : GetColumnWithOffset(parentName, columnIndex);
            if (_columns.ContainsKey(columnIndex))
                throw ExcelWorksheetMapperException.ColumnIndexAlreadyExists(
                    addedColumn: columnName,
                    index: columnIndex,
                    alreadyContainedName: _columns[columnIndex]);

            columnName = parentName == null ? columnName : IncludeName(parentName, columnName);

            _columns[columnIndex] = columnName;
        }

        public void Include(int offset, string name, string parentName = null)
        {
            if (parentName == null)
            {
                _includings.Add(name, new Including
                {
                    Name = name,
                    Offset = offset
                });

                return;
            }

            foreach (string includingKey in _includings.Keys)
            {
                if (!_includings[includingKey].Name.EndsWith(parentName))
                    continue;

                _includings[includingKey].Append(name, offset);
            }
        }

        private string IncludeName(string parentName, string columnName)
        {
            string oldName = _includings.Values
                .SingleOrDefault(including => including.Name.EndsWith(parentName))?.Name;

            return $"{oldName}.{columnName}";
        }

        private int GetColumnWithOffset(string parentName, int columnIndex)
        {
            var offset = _includings.Values
                .SingleOrDefault(including => including.Name.EndsWith(parentName))?.Offset;

            if (!offset.HasValue) return columnIndex;

            return columnIndex + offset.Value;
        }

        private class Including
        {
            public string Name { get; set; }
            public int Offset { get; set; }

            public void Append(string name, int offset)
            {
                Name += $".{name}";
                Offset += offset;
            }

            public override string ToString()
            {
                return $"{nameof(Name)}: {Name}, {nameof(Offset)}: {Offset}";
            }
        }
    }
}