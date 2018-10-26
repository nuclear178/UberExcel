using ExcelTools.Converters;

namespace ConsoleTests.Json
{
    public class IsoDateTimeConverterBase : DateTimeConverterBase
    {
        public IsoDateTimeConverterBase()
        {
            DateTimeFormat = "yyyy-MM-dd--hh-mm-ss";
        }
    }
}