using System;
using System.Collections.Generic;
using ExcelTools;
using ExcelTools.IO;

namespace ConsoleTests
{
    internal static class Program
    {
        private static void Main()
        {
            var fuels = new List<Fuel>
            {
                new Fuel {Type = "Wog", Volume = 5, Name = new Name {First = "row1", Last = "col1"}},
                new Fuel {Type = "Shell", Volume = 3, Name = new Name {First = "row2", Last = "col2"}}
            };
            var convert = WorksheetConvert<Fuel>.BuildAttributeBased();
            string fuelsWorksheet = convert.SerializeObject(fuels);

            Console.WriteLine(fuelsWorksheet);

            var analyzer = new AttributeBasedIntrospector(typeof(Fuel));
            Console.WriteLine($"Columns number: {analyzer.Analyze()}");
        }
    }
}