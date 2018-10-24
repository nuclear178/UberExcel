using ExcelTools.Meta.Worksheet;

namespace ConsoleTests
{
    public class Name
    {
        [Column(3)] public string First { get; set; } //TODO :: add offset
        [Column(4)] public string Last { get; set; }
    }
}