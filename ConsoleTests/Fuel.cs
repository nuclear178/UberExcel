using ExcelTools.Meta.Attributes;

namespace ConsoleTests
{
    public class Fuel
    {
        [Column(1)] public string Type { get; set; }

        [Column(2)] public int Volume { get; set; }

        [Include(2)] public Bucket Bucket { get; set; }

        public Fuel()
        {
        }

        public override string ToString()
        {
            return
                $"{nameof(Type)}: {Type}, {nameof(Volume)}: {Volume}, {nameof(Bucket)}: {Bucket.ToString() ?? "null"}";
        }
    }
}