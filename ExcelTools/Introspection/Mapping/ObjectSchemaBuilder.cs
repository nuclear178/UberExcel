using System;
using System.Collections.Generic;
using System.Linq;
using ExcelTools.Exceptions;

namespace ExcelTools.Introspection.Mapping
{
    public class ObjectSchemaBuilder
    {
        private readonly Type _type;
        private readonly Dictionary<int, Column> _columns;
        private readonly List<Including> _includings;

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
                    column.ConverterType)),
                _type
            );
        }

        public void AddColumn(
            int columnIndex,
            string columnName,
            out int addedIndex,
            string parentName = null)
        {
            columnIndex = parentName == null
                ? columnIndex
                : GetColumnIndexWithOffset(parentName, columnIndex);
            if (_columns.ContainsKey(columnIndex))
                throw ExcelWorksheetMapperException.ColumnIndexAlreadyExists(
                    addedColumn: columnName,
                    columnIndex: columnIndex,
                    alreadyContainedName: _columns[columnIndex].FullName);

            columnName = parentName == null
                ? columnName
                : IncludeName(parentName, columnName);

            _columns[columnIndex] = new Column(columnIndex, columnName);

            addedIndex = columnIndex;
        }

        public void IncludeWithOffset(int offset, string includingName, string parentName = null)
        {
            if (parentName == null)
            {
                _includings.Add(new Including(includingName, offset));
            }
            else
            {
                _includings.SingleOrDefault(incl => incl.Name.EndsWith(parentName))
                    ?.Complete(includingName, offset);
            }
        }

        public void Include(string includingName, string parentName = null)
        {
            IncludeWithOffset(0, includingName, parentName);
        }

        public void AddConverter(int columnIndex, Type converterType) =>
            _columns[columnIndex].ConverterType = converterType;

        private string IncludeName(string parentName, string columnName)
        {
            string oldName = _includings
                .SingleOrDefault(including => including.Name.EndsWith(parentName))?.Name;

            return $"{oldName}.{columnName}";
        }

        private int GetColumnIndexWithOffset(string parentName, int columnIndex)
        {
            int? offset = _includings
                .SingleOrDefault(including => including.Name.EndsWith(parentName))?.Offset;

            if (!offset.HasValue) return columnIndex;

            return columnIndex + offset.Value;
        }

        private class Column
        {
            public int Index { get; }
            public string FullName { get; }
            public Type ConverterType { get; set; }

            public Column(int index, string fullName)
            {
                Index = index;
                FullName = fullName;
            }
        }

        private class Including
        {
            public string Name { get; private set; }
            public int Offset { get; private set; }

            public Including(string name, int offset)
            {
                Name = name;
                Offset = offset;
            }

            public void Complete(string name, int offset)
            {
                Name += $".{name}";
                Offset += offset;
            }
        }
    }
}