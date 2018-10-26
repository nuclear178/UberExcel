using System;
using System.Collections.Generic;
using System.Linq;
using ExcelTools.Exceptions;

namespace ExcelTools.Introspection
{
    public class ObjectSchemaBuilder
    {
        private readonly Dictionary<int, Column> _columns;
        private readonly List<Including> _includings;
        private readonly Type _type;

        public ObjectSchemaBuilder(Type type)
        {
            _type = type;
            _columns = new Dictionary<int, Column>();
            _includings = new List<Including>();
        }

        public ObjectSchema Build()
        {
            return new ObjectSchema(
                _columns.Values.Select(column => new ColumnOptions(
                    column.Index,
                    column.FullName,
                    column.Type)),
                _type
            );
        }

        public void AddColumn(int columnIndex, string columnName, Type columnType, string parentName = null)
        {
            columnIndex = parentName == null ? columnIndex : GetColumnWithOffset(parentName, columnIndex);
            if (_columns.ContainsKey(columnIndex))
                throw ExcelWorksheetMapperException.ColumnIndexAlreadyExists(
                    addedColumn: columnName,
                    index: columnIndex,
                    alreadyContainedName: _columns[columnIndex].FullName);

            columnName = parentName == null ? columnName : IncludeName(parentName, columnName);

            _columns[columnIndex] = new Column(columnIndex, columnName, columnType);
        }

        public void Include(int offset, string name, string parentName = null)
        {
            if (parentName == null)
            {
                _includings.Add(new Including
                {
                    Name = name,
                    Offset = offset
                });

                return;
            }

            _includings.SingleOrDefault(incl => incl.Name.EndsWith(parentName))
                ?.Complete(name, offset);
        }

        private string IncludeName(string parentName, string columnName)
        {
            string oldName = _includings
                .SingleOrDefault(including => including.Name.EndsWith(parentName))?.Name;

            return $"{oldName}.{columnName}";
        }

        private int GetColumnWithOffset(string parentName, int columnIndex)
        {
            var offset = _includings
                .SingleOrDefault(including => including.Name.EndsWith(parentName))?.Offset;

            if (!offset.HasValue) return columnIndex;

            return columnIndex + offset.Value;
        }

        private class Including
        {
            public string Name { get; set; }
            public int Offset { get; set; }

            public void Complete(string name, int offset)
            {
                Name += $".{name}";
                Offset += offset;
            }
        }

        private class Column
        {
            public int Index { get; }
            public string FullName { get; }
            public Type Type { get; }

            public Column(int index, string fullName, Type type)
            {
                Index = index;
                FullName = fullName;
                Type = type;
            }
        }
    }
}