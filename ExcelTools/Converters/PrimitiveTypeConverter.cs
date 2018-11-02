using System;
using System.ComponentModel;

namespace ExcelTools.Converters
{
    public class PrimitiveTypeConverter : IConverter
    {
        private readonly Type _type;
        private readonly bool _nullable;

        public PrimitiveTypeConverter(Type type, bool nullable)
        {
            _type = type;
            _nullable = nullable;
        }

        public string Write(object value)
        {
            return value.ToString();
        }

        public object Read(string value)
        {
            switch (Type.GetTypeCode(_type))
            {
                case TypeCode.Boolean:
                    return _nullable ? ToNullable<bool>(value) : Convert.ToBoolean(value);
                case TypeCode.Byte:
                    return _nullable ? ToNullable<byte>(value) : Convert.ToByte(value);
                case TypeCode.Char:
                    return _nullable ? ToNullable<char>(value) : Convert.ToChar(value);
                case TypeCode.DateTime:
                    break;
                case TypeCode.DBNull:
                    break;
                case TypeCode.Decimal:
                    return _nullable ? ToNullable<decimal>(value) : Convert.ToDecimal(value);
                case TypeCode.Double:
                    return _nullable ? ToNullable<double>(value) : Convert.ToDouble(value);
                case TypeCode.Empty:
                    break;
                case TypeCode.Int16:
                    return _nullable ? ToNullable<short>(value) : Convert.ToInt16(value);
                case TypeCode.Int32:
                    return _nullable ? ToNullable<int>(value) : Convert.ToInt32(value);
                case TypeCode.Int64:
                    return _nullable ? ToNullable<long>(value) : Convert.ToInt64(value);
                case TypeCode.Object:
                    break;
                case TypeCode.SByte:
                    break;
                case TypeCode.Single:
                    break;
                case TypeCode.String:
                    return value;
                case TypeCode.UInt16:
                    return _nullable ? ToNullable<ushort>(value) : Convert.ToUInt16(value);
                case TypeCode.UInt32:
                    return _nullable ? ToNullable<uint>(value) : Convert.ToUInt32(value);
                case TypeCode.UInt64:
                    return _nullable ? ToNullable<ulong>(value) : Convert.ToUInt64(value);
                default:
                    throw new ArgumentOutOfRangeException();
            }

            throw new ArgumentOutOfRangeException();
        }

        // TODO: Make non-static using _type
        private static T? ToNullable<T>(string @string) where T : struct
        {
            if (string.IsNullOrWhiteSpace(@string))
                return new T?();

            TypeConverter converter = TypeDescriptor.GetConverter(typeof(T));
            return (T?) converter.ConvertFrom(@string);
        }
    }
}