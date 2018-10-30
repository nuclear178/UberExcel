using ExcelTools.Introspection.Mapping;
using OfficeOpenXml;

namespace ExcelTools.IO.Xlsx
{
    public class XlsxObjectReader
    {
        private readonly ObjectSchema _schema;
        //private IValueMapper _mapper;

        public XlsxObjectReader(ObjectSchema schema)
        {
            _schema = schema;
        }

        public void Build(ExcelRange cells, int rowIndex, object obj)
        {
            foreach (ColumnOptions column in _schema)
            {
                object rawValue = column.GetValue(obj);
                cells[rowIndex, column.Index].Value = column.MapValueTo(rawValue);
            }
        }
    }
}