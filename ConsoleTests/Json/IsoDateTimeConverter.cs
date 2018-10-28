using ExcelTools.Converters;

namespace ConsoleTests.Json
{
    public class IsoDateTimeConverter : DateTimeConverterBase
    {
        public IsoDateTimeConverter()
        {
            DateTimeFormat = "yyyy-MM-dd--hh-mm-ss";
        }
    }
}