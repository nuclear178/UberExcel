using ExcelTools.Meta.Attributes;

namespace ConsoleTests
{
    public class Bucket
    {
        public Bucket()
        {
        }

        [Column(1)] public int Prop1 { get; set; }
        [Include(1)] public Name Prop2 { get; set; }

        public override string ToString()
        {
            return $"{nameof(Prop1)}: {Prop1}, {nameof(Prop2)}: {Prop2.ToString() ?? "null"}";
        }
    }
}