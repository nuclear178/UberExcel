using ExcelTools.Meta.Worksheet;

namespace ConsoleTests
{
    public class Name
    {
        [Column(3)] public string First { get; set; }
        [Column(4)] public string Last { get; set; }
    }
}