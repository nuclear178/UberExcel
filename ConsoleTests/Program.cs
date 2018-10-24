using System;
using System.Reflection;
using ExcelTools.IO;

namespace ConsoleTests
{
    internal static class Program
    {
        private static void Main()
        {
            /*var fuels = new List<Fuel>
            {
                new Fuel {Type = "Wog", Volume = 5, Name = new Name {First = "row1", Last = "col1"}},
                new Fuel {Type = "Shell", Volume = 3, Name = new Name {First = "row2", Last = "col2"}}
            };*/
            //using (var package = new ExcelPackage())
            //{
            //    var convert = WorksheetConvert<Fuel>.BuildAttributeBased();
            //    ExcelWorksheet fuelsWorksheet = package.Workbook.Worksheets[1];
            //    convert.SerializeObject(fuels, fuelsWorksheet);

             //   Console.WriteLine(fuelsWorksheet);
            //}


            var analyzer = new AttributeBasedIntrospector(typeof(Fuel));
            Console.WriteLine($"Columns number: {analyzer.Analyze()}");

            /*Fuel fuelItem = fuels[0];
            object val = GetPropValue("Name.First", fuelItem);

            Console.WriteLine(val);*/
        }

        private static object GetPropValue(string propAddress, object obj)
        {
            foreach (string part in propAddress.Split('.'))
            {
                if (obj == null) return null;

                Type type = obj.GetType();
                PropertyInfo info = type.GetProperty(part);
                if (info == null) return null;

                obj = info.GetValue(obj, null);
            }

            return obj;
        }
    }
}