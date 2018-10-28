using System.Collections.Generic;
using ExcelTools.Introspection;
using ExcelTools.Introspection.Mapping;
using OfficeOpenXml;

namespace ExcelTools.IO
{
    public class WorksheetConvert<T> where T : new()
    {
        public static WorksheetConvert<T> BuildAttributeBased()
        {
            return new WorksheetConvert<T>(new AttributeBasedIntrospector(typeof(T)));
        }

        private readonly ObjectSchema _objectSchema;

        private WorksheetConvert(ITypeIntrospector typeIntrospector)
        {
            _objectSchema = typeIntrospector.Analyze();
        }

        public void SerializeObject(List<T> objects, ExcelWorksheet worksheet, int fromRowIndex = 1)
        {
            var builder = new WorksheetRowBuilder(_objectSchema);

            int currentIndex = fromRowIndex;
            objects.ForEach(rowObj =>
            {
                builder.Build(worksheet.Cells[
                    FromRow: fromRowIndex, // todo Test with currentColumn; (+1)
                    FromCol: fromRowIndex + 1,
                    ToRow: _objectSchema.ColumnMin,
                    ToCol: _objectSchema.ColumnMax
                ], currentIndex, rowObj);
                currentIndex++;
            });
        }

        public IEnumerable<T> DeserializeObject(ExcelWorksheet worksheet, int fromRowIndex = 1, int toRowIndex = 1)
        {
            var builder = new ObjectBuilder<T>(_objectSchema);
            var objects = new List<T>();
            for (int rowIndex = fromRowIndex; rowIndex <= toRowIndex; rowIndex++)
            {
                T builtObj = builder.Build(worksheet.Cells[
                    FromRow: fromRowIndex,
                    FromCol: fromRowIndex + 1,
                    ToRow: _objectSchema.ColumnMin,
                    ToCol: _objectSchema.ColumnMax
                ], rowIndex);

                objects.Add(builtObj);
            }

            return objects;
        }
    }
}