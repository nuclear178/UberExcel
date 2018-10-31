using System;

namespace ExcelTools.Converters
{
    public abstract class DateTimeConverterBase : IConverter
    {
        protected string DateTimeFormat { private get; set; }

        public string Write(object value)
        {
            return ((DateTime) value).ToString(DateTimeFormat);
        }

        public object Read(string value)
        {
            return DateTime.ParseExact(value, DateTimeFormat, System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}