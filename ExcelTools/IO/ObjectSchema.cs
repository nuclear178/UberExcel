using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTools.IO
{
    public class ObjectSchema : IEnumerable<ColumnOptions>
    {
        private readonly IEnumerable<ColumnOptions> _rowObjects;

        public ObjectSchema(IEnumerable<ColumnOptions> rowObjects)
        {
            _rowObjects = rowObjects;
        }

        public int ColumnMin => _rowObjects.Select(rowObj => rowObj.Index).Min();

        public int ColumnMax => _rowObjects.Select(rowObj => rowObj.Index).Max();

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