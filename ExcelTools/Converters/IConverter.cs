namespace ExcelTools.Converters
{
    public interface IConverter
    {
        string Write(object val);

        object Read(string val);
    }
}