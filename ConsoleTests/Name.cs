using ExcelTools.Meta.Worksheet;

namespace ConsoleTests
{
    public class Name
    {
        [Column(7)] public string First { get; set; } //TODO :: add offset
        [Column(8)] public string Last { get; set; }
    }
}