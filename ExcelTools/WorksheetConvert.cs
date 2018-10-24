using System.Collections.Generic;
using ExcelTools.IO;
using OfficeOpenXml;

namespace ExcelTools
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

        public void SerializeObject(List<T> rows, ExcelWorksheet worksheet)
        {
        }

        public IEnumerable<T> DeserializeObject(ExcelRange worksheet)
        {
            var objectBuilder = new ObjectBuilder<T>(_objectSchema);


            return new List<T>();
        }
    }
}