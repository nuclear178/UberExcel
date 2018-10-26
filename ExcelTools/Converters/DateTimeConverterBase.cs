using System;

namespace ExcelTools.Converters
{
    public abstract class DateTimeConverterBase : IConverter
    {
        protected string DateTimeFormat { private get; set; }

        public string Write(object val)
        {
            return ((DateTime) val).ToString(DateTimeFormat);
        }

        public object Read(string val)
        {
            return DateTime.ParseExact(val, DateTimeFormat, System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}