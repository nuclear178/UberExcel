using System;
using System.Diagnostics.CodeAnalysis;

namespace ExcelTools.Exceptions
{
    public class ExcelWorksheetMapperException : Exception
    {
        [SuppressMessage("ReSharper", "MemberCanBePrivate.Global")]
        public ExcelWorksheetMapperException(string message) : base(message)
        {
        }

        public static ExcelWorksheetMapperException ColumnIndexAlreadyExists(int index, string alreadyContainedName)
        {
            return new ExcelWorksheetMapperException(
                $"Column with specified index [{index}] already exists and contains column with name {alreadyContainedName}");
        }

        public static ExcelWorksheetMapperException ColumnNotFound(int columnIndex)
        {
            return new ExcelWorksheetMapperException($"Column with specified index [{columnIndex}] is not found");
        }
    }
}