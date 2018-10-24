using ExcelTools.Meta.Worksheet;

namespace ConsoleTests
{
    public class Name
    {
        [Column(1)] public string First { get; set; } //TODO :: add offset
        [Column(2)] public string Last { get; set; }
    }
}