using System.Collections.Generic;
using ExcelTools.Introspection;
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
                    FromRow: fromRowIndex,
                    FromCol: fromRowIndex + 1,
                    ToRow: _objectSchema.ColumnMin,
                    ToCol: _objectSchema.ColumnMax
                ], currentIndex, rowObj);
                currentIndex++;
            });
        }

        public IEnumerable<T> DeserializeObject(ExcelRange worksheet)
        {
            var objectBuilder = new ObjectBuilder<T>(_objectSchema);


            return new List<T>();
        }
    }
}