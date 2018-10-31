namespace ExcelTools.Converters
{
    public interface IConverter
    {
        string Write(object value);

        object Read(string value);
    }
}