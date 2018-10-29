using System.Collections.Generic;
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
                new AttributeTypeIntrospector(typeof(T))
            );
        }

        public static XlsxSerializer<T> BuildJsonBased(JObject mappingJson)
        {
            return new XlsxSerializer<T>(
                new JsonTypeIntrospector(typeof(T), mappingJson)
            );
        }

        private readonly ObjectSchema _objectSchema;

        private XlsxSerializer(ITypeIntrospector typeIntrospector)
        {
            _objectSchema = typeIntrospector.Analyze();
        }

        public void SerializeObject(List<T> objects, ExcelWorksheet worksheet, int fromRowIndex = 1)
        {
            var reader = new XlsxObjectReader(_objectSchema);

            int currentIndex = fromRowIndex;
            objects.ForEach(rowObj =>
            {
                reader.Build(worksheet.Cells[
                    FromRow: currentIndex,
                    FromCol: _objectSchema.ColumnMin,
                    ToRow: currentIndex + 1,
                    ToCol: _objectSchema.ColumnMax
                ], currentIndex, rowObj);
                currentIndex++;
            });
        }

        public IEnumerable<T> DeserializeObject(ExcelWorksheet worksheet, int fromRowIndex = 1, int toRowIndex = 1)
        {
            var writer = new XlsxObjectWriter<T>(_objectSchema);
            var objects = new List<T>();
            for (int rowIndex = fromRowIndex; rowIndex <= toRowIndex; rowIndex++)
            {
                T builtObj = writer.Build(worksheet.Cells[
                    FromRow: fromRowIndex,
                    FromCol: _objectSchema.ColumnMin,
                    ToRow: fromRowIndex + 1,
                    ToCol:
                    _objectSchema.ColumnMax
                ], rowIndex);

                objects.Add(builtObj);
            }

            return objects;
        }
    }
}