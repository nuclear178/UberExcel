using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using ExcelTools.Introspection;
using ExcelTools.IO;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

namespace ConsoleTests
{
    [SuppressMessage("ReSharper", "UnusedMember.Local")]
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


            /*var fuel = new Fuel();

            Stopwatch watch = Stopwatch.StartNew();
            ObjectBuilder<Fuel>.SetPropValue("Bucket.Prop2.First", fuel, "f");
            ObjectBuilder<Fuel>.SetPropValue("Bucket.Prop2.Last", fuel, "l");

            ObjectBuilder<Fuel>.SetPropValue("Bucket.Prop2.First", fuel, "v");

            watch.Stop();

            long elapsedMs = watch.ElapsedMilliseconds;

            Console.WriteLine($"val: {fuel.Bucket.Prop2.First} and {fuel.Bucket.Prop2.Last}");

            Console.WriteLine($"elp: {elapsedMs}");

            Console.WriteLine("Opening...");

            watch = Stopwatch.StartNew();
            OpenExcel();

            watch.Stop();
            elapsedMs = watch.ElapsedMilliseconds;

            Console.WriteLine($"Time elapsed opening xlsx: {elapsedMs}");*/

            /*Activator.CreateInstance(typeof(Name));
            Activator.CreateInstance(typeof(Bucket));
            fuel = Activator.CreateInstance<Fuel>();

            Console.WriteLine(fuel);

            Console.WriteLine(fuel.Bucket == null);

            PropertyInfo info = fuel.GetType().GetProperty("Bucket");
            object value = info.GetValue(fuel, null);

            if (value == null)
            {
                object propObj = Activator.CreateInstance(info.PropertyType);
                info.SetValue(fuel, propObj, null);
            }

            Console.WriteLine(fuel.Bucket == null);*/

            //CreateExcel();
            //OpenExcel();

            OpenJson();
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

        /*private void InstantiateWithIncludings(object obj) //todo separate deep instantiation 
        {
        }*/


        private static void CreateExcel()
        {
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
                            CreationDate = DateTime.Now,
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
                            CreationDate = DateTime.Now,
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

        private static void OpenExcel()
        {
            const string fileName = "Example-CRM-2018-10-28--11-04-52.xlsx";

            var file = new FileInfo(fileName);
            using (var package = new ExcelPackage(file))
            {
                var convert = WorksheetConvert<Fuel>.BuildAttributeBased();
                ExcelWorksheet fuelsWorksheet = package.Workbook.Worksheets[1];

                if (fuelsWorksheet == null)
                {
                    Console.WriteLine("returning....");
                    return;
                }

                int totalRows = fuelsWorksheet.Dimension.Rows;


                IEnumerable<Fuel> fuels = convert.DeserializeObject(fuelsWorksheet, 1, totalRows);
                foreach (Fuel fuel in fuels)
                {
                    Console.WriteLine(fuel);
                }
            }
        }

        private static void OpenJson()
        {
            string fileJson = File.ReadAllText("mapping.json");
            JObject mappingJson = JObject.Parse(fileJson);

            //Console.WriteLine(mappingJson["data"]);

            //WorksheetConvert<Fuel>.BuildJsonBased((JObject) mappingJson["data"]);

            var introspector = new JsonTypeIntrospector(typeof(Fuel), (JObject) mappingJson["data"]);
            introspector.Analyze();
        }
    }
}