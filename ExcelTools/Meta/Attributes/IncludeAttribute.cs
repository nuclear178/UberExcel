using System;

namespace ExcelTools.Meta.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class IncludeAttribute : Attribute
    {
        public int Offset { get; }

        public IncludeAttribute(int offset = 0)
        {
            Offset = offset;
        }
    }
}