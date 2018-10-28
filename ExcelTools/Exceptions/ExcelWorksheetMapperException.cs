using System;
using System.Diagnostics.CodeAnalysis;
using ExcelTools.Converters;

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
            int columnIndex,
            string alreadyContainedName)
        {
            return new ExcelWorksheetMapperException(
                $"Failed to add column with name [{addedColumn}]: column with specified index [{columnIndex}] already exists and contains column with name [{alreadyContainedName}].");
        }

        public static Exception UnsupportedColumnType(string typeName)
        {
            return new ExcelWorksheetMapperException($"Column with type [{typeName}] is not supported.");
        }

        public static Exception InvalidConverterType(string converterTypeName)
        {
            return new ExcelWorksheetMapperException(
                $"Invalid converter type [{converterTypeName}]: converter must implement {typeof(IConverter).FullName} interface.");
        }
    }
}