using System;
using ConsoleTests.Json;
using ExcelTools.Meta.Attributes;

namespace ConsoleTests
{
    public class Name
    {
        [Column(1)] public string First { get; set; }

        [Column(2)]
        [Converter(typeof(IsoDateTimeConverter))]
        public DateTime CreationDate { get; set; }

        [Column(3)] public string Last { get; set; }

        public Name()
        {
        }

        public override string ToString()
        {
            return $"{nameof(First)}: {First}, {nameof(CreationDate)}: {CreationDate}, {nameof(Last)}: {Last}";
        }
    }
}