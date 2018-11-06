using ExcelTools.Introspection.Mapping;
using OfficeOpenXml;

namespace ExcelTools.IO.Xlsx
{
    public class XlsxObjectWriter<T> where T : new()
    {
        private readonly ObjectSchema _schema;

        public XlsxObjectWriter(ObjectSchema schema)
        {
            _schema = schema;
        }

        public T Build(ExcelRange cells, int rowIndex)
        {
            var createdObj = new T();
            foreach (ColumnOptions column in _schema)
            {
                object rawValue = cells[rowIndex, column.Index].Value;
                column.SetValue(createdObj, rawValue);
            }

            return createdObj;
        }
    }
}