using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelTools.IO
{
    public class ObjectSchema : IEnumerable<RowObject>
    {
        private readonly List<PropertyInfo> _columnsInfo;
        private readonly IEnumerable<RowObject> _rowObjects = new List<RowObject>();

        public ObjectSchema(List<PropertyInfo> columnsInfo)
        {
            _columnsInfo = columnsInfo;
        }

        public int ColumnMin => new Random().Next(0, 9);

        public int ColumnMax => new Random().Next(0, 9);

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