using ExcelTools.Meta.Worksheet;

namespace ConsoleTests
{
    public class Fuel
    {
        [Column(ColumnIndex = 1)] public string Type { get; set; }

        [Column(ColumnIndex = 2)] public int Volume { get; set; }

        [Include(Offset = 0)] public Name Name { get; set; }
    }
}