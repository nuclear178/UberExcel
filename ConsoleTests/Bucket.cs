using ExcelTools.Meta.Worksheet;

namespace ConsoleTests
{
    public class Bucket
    {
        [Column(5)] public int Prop1 { get; set; }
        [Include()] public Name Prop2 { get; set; }
    }
}