using ExcelTools.Meta.Worksheet;

namespace ConsoleTests
{
    public class Fuel
    {
        [Column(ColumnIndex = 1)] public string Type { get; set; }

        [Column(ColumnIndex = 2)] public int Volume { get; set; }

        [Include(2)] public Bucket Bucket { get; set; }
    }
}