using System.Collections.Generic;
using ExcelTools.Introspection;
using ExcelTools.Introspection.Mapping;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

namespace ExcelTools.IO
{
    public class WorksheetConvert<T> where T : new()
    {
        public static WorksheetConvert<T> BuildAttributeBased()
        {
            return new WorksheetConvert<T>(
                new AttributeTypeIntrospector(typeof(T))
            );
        }

        public static WorksheetConvert<T> BuildJsonBased(JObject mappingJson)
        {
            return new WorksheetConvert<T>(
                new JsonTypeIntrospector(typeof(T), mappingJson)
            );
        }

        private readonly ObjectSchema _objectSchema;

        private WorksheetConvert(ITypeIntrospector typeIntrospector)
        {
            _objectSchema = typeIntrospector.Analyze();
        }

        public void SerializeObject(List<T> objects, ExcelWorksheet worksheet, int fromRowIndex = 1)
        {
            var reader = new ObjectReader(_objectSchema);

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
            var writer = new ObjectWriter<T>(_objectSchema);
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