using ExcelTools.Introspection.Mapping;
using OfficeOpenXml;

namespace ExcelTools.IO.Xlsx
{
    public class XlsxObjectReader
    {
        private readonly ObjectSchema _schema;

        public XlsxObjectReader(ObjectSchema schema)
        {
            _schema = schema;
        }

        public void Build(ExcelRange cells, int rowIndex, object obj)
        {
            foreach (ColumnOptions column in _schema)
            {
                cells[rowIndex, column.Index].Value = column.GetValue(obj);
            }
        }
    }
}