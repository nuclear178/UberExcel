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

        public static ExcelWorksheetMapperException ColumnIndexAlreadyExists(
            string addedColumn,
            int index,
            string alreadyContainedName)
        {
            return new ExcelWorksheetMapperException(
                $"Failed to add column with name [{addedColumn}]: column with specified index [{index}] already exists and contains column with name [{alreadyContainedName}]");
        }
    }
}