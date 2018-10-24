using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTools.IO
{
    public class ObjectSchema : IEnumerable<RowObject>
    {
        private readonly IEnumerable<RowObject> _rowObjects;

        public ObjectSchema(IEnumerable<RowObject> rowObjects)
        {
            _rowObjects = rowObjects;
        }

        public int ColumnMin => _rowObjects.Select(rowObj => rowObj.ColumnIndex).Min();

        public int ColumnMax => _rowObjects.Select(rowObj => rowObj.ColumnIndex).Max();

        public IEnumerator<RowObject> GetEnumerator()
        {
            return _rowObjects.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}