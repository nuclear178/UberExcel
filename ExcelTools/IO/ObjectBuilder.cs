using OfficeOpenXml;

namespace ExcelTools.IO
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
            foreach (RowObject rowData in _schema)
            {
                int columnIndex = rowData.ColumnIndex;
                object value = cells[rowIndex, columnIndex].Value;
            }

            return obj;
        }
    }
}