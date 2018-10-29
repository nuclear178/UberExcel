using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTools.Introspection.Mapping
{
    public class ObjectSchema : IEnumerable<ColumnOptions>
    {
        private readonly IEnumerable<ColumnOptions> _rowObjects;
        private readonly Type _type;

        public ObjectSchema(IEnumerable<ColumnOptions> rowObjects, Type type)
        {
            _rowObjects = rowObjects;
            _type = type;
        }

        public int ColumnMin => _rowObjects.Select(rowObj => rowObj.Index).Min();

        public int ColumnMax => _rowObjects.Select(rowObj => rowObj.Index).Max();

        public object CreateEmpty()
        {
            return _type;
        }

        public IEnumerator<ColumnOptions> GetEnumerator()
        {
            return _rowObjects.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}