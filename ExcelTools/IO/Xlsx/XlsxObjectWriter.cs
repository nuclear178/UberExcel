using ExcelTools.Converters.Mappers;
using ExcelTools.Introspection.Mapping;
using OfficeOpenXml;

namespace ExcelTools.IO.Xlsx
{
    public class XlsxObjectWriter<T> where T : new()
    {
        private readonly ObjectSchema _schema;
        private readonly IValueMapper _mapper;

        public XlsxObjectWriter(ObjectSchema schema, IValueMapper mapper)
        {
            _schema = schema;
            _mapper = mapper;
        }

        public T Build(ExcelRange cells, int rowIndex)
        {
            var createdObj = new T();
            foreach (ColumnOptions column in _schema)
            {
                object rawValue = cells[rowIndex, column.Index].Value;
                column.SetValue(createdObj, column.MapValueFrom(rawValue, _mapper));
            }

            return createdObj;
        }
    }
}