namespace ExcelTools.IO
{
    public class ObjectBuilder<T> where T : new()
    {
        private readonly ObjectSchema _schema;

        public ObjectBuilder(ObjectSchema schema)
        {
            _schema = schema;
        }

        public T Build()
        {
            return new T();
        }
    }
}