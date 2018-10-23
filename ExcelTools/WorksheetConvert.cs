using System.Collections.Generic;
using ExcelTools.IO;

namespace ExcelTools
{
    public static class WorksheetConvert
    {
        public static string SerializeObject<T>(List<T> value)
        {
            var analyzer = new TypeAnalyzer(typeof(T));
            analyzer.Analyze();

            return string.Empty;
        }

        public static IEnumerable<T> DeserializeObject<T>()
        {
            return new List<T>();
        }
    }
}