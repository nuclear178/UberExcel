using ExcelTools.Converters;
using ExcelTools.Introspection;
using ExcelTools.Introspection.Mapping;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

namespace ExcelTools.IO.Xlsx
{
    public class XlsxSerializer<T> where T : new()
    {
        public static XlsxSerializer<T> BuildAttributeBased()
        {
            return new XlsxSerializer<T>(
                new AttributeTypeIntrospector(
                    objType: typeof(T),
                    converterFactory: new TypeConverterFactory()
                )
            );
        }

        public static XlsxSerializer<T> BuildJsonBased(JObject mappingJson)
        {
            return new XlsxSerializer<T>(
                new JsonTypeIntrospector(
                    objType: typeof(T),
                    converterFactory: new TypeConverterFactory(),
                    mappingJson: mappingJson["data"]
                )
            );
        }

        private readonly ObjectSchema _objectSchema;

        // ReSharper disable once MemberCanBePrivate.Global
        public XlsxSerializer(ITypeIntrospector typeIntrospector)
        {
            _objectSchema = typeIntrospector.Analyze();
        }

        public void SerializeObject(T obj, ExcelWorksheet worksheet, int rowIndex)
        {
            var reader = new XlsxObjectReader(_objectSchema);
            reader.Build(worksheet.Cells[
                FromRow: rowIndex,
                FromCol: _objectSchema.ColumnMin,
                ToRow: rowIndex + 1,
                ToCol: _objectSchema.ColumnMax
            ], rowIndex, obj);
        }

        public T DeserializeObject(ExcelWorksheet worksheet, int rowIndex)
        {
            var writer = new XlsxObjectWriter<T>(_objectSchema);
            return writer.Build(worksheet.Cells[
                FromRow: rowIndex,
                FromCol: _objectSchema.ColumnMin,
                ToRow: rowIndex + 1,
                ToCol: _objectSchema.ColumnMax
            ], rowIndex);
        }
    }
}