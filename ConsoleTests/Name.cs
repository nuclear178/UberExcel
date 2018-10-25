using ExcelTools.Meta.Worksheet;

namespace ConsoleTests
{
    public class Name
    {
        [Column(1)] public string First { get; set; } //TODO :: add offset
        [Column(3)] public string Last { get; set; }

        public Name()
        {
        }
    }
}