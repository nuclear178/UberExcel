using System.Collections.Generic;
using ExcelTools.IO;

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

        public string SerializeObject(List<T> rows)
        {
            var objectBuilder = new ObjectBuilder<T>(_objectSchema);

            return string.Empty;
        }

        public IEnumerable<T> DeserializeObject()
        {
            return new List<T>();
        }
    }
}