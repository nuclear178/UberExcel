using System;
using System.Collections.Generic;
using System.IO;
using ExcelTools;
using OfficeOpenXml;

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


            /*var analyzer = new AttributeBasedIntrospector(typeof(Fuel));
            Console.WriteLine($"Columns number: {analyzer.Analyze()}");*/

            /*Fuel fuelItem = fuels[0];
            object val = GetPropValue("Name.First", fuelItem);

            Console.WriteLine(val);*/

            var fuels = new List<Fuel>
            {
                new Fuel
                {
                    Type = "tp1",
                    Volume = 5,
                    Bucket = new Bucket
                    {
                        Prop1 = 1, Prop2 = new Name
                        {
                            First = "fr1",
                            Last = "ls1"
                        }
                    }
                },
                new Fuel
                {
                    Type = "tp2",
                    Volume = 7,
                    Bucket = new Bucket
                    {
                        Prop1 = 2, Prop2 = new Name
                        {
                            First = "fr5",
                            Last = "ls7"
                        }
                    }
                }
            };

            string fileName = "Example-CRM-" + DateTime.Now.ToString("yyyy-MM-dd--hh-mm-ss") + ".xlsx";

            // Create the file using the FileInfo object
            var file = new FileInfo(fileName);
            using (var package = new ExcelPackage(file))
            {
                var convert = WorksheetConvert<Fuel>.BuildAttributeBased();
                ExcelWorksheet fuelsWorksheet = package.Workbook.Worksheets.Add("sl");
                convert.SerializeObject(fuels, fuelsWorksheet);

                package.Save();
            }
        }

        /*private static object GetPropValue(string propAddress, object obj)
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
        }*/

        /*Console.WriteLine("::");

            foreach (var keyValuePair in _columns)
            {
                Console.WriteLine($"[{keyValuePair.Key}]: {keyValuePair.Value}");
            }

            Console.WriteLine("::");
            _includings.Values.ToList().ForEach(Console.WriteLine);*/

        /*private class Checker<T>
        {
            private readonly Checker<T> _next;
            private readonly Predicate<T> _rule;

            public Checker(Predicate<T> rule, Checker<T> next = null)
            {
                _rule = rule;
                _next = next;
            }

            public bool Check(T entity)
            {
                if (_rule(entity)) return true;
                return _next == null || _next.Check(entity);
            }
        }

        class Entity
        {
            public int Term { get; set; }
        }

        private static class Rules
        {
            public static Predicate<Entity> TermChecker()
            {
                return e => e.Term > 10;
            }
        }

        private class AntiFraudService
        {
            private readonly Checker<Entity> _validator;

            public AntiFraudService()
            {
                _validator = new Checker<Entity>(Rules.TermChecker());
            }
        }*/
    }
}