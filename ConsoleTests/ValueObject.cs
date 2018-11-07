using System;

// ReSharper disable NonReadonlyMemberInGetHashCode

namespace ConsoleTests
{
    public sealed class ValueObject
    {
        private const int MinNameLength = 5;

        private string _name;

        public ValueObject(string name, int number)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Number = number;
        }

        public string Name
        {
            get => this._name;
            private set
            {
                if (value.Trim().Length < MinNameLength)
                    throw new ArgumentException($"Name must have length not less than {MinNameLength}");

                this._name = value;
            }
        }

        public int Number { get; }

        private bool Equals(ValueObject other)
        {
            return string.Equals(_name, other._name) && Number == other.Number;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            return obj is ValueObject other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((_name != null ? _name.GetHashCode() : 0) * 397) ^ Number;
            }
        }

        public static bool operator ==(ValueObject left, ValueObject right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(ValueObject left, ValueObject right)
        {
            return !Equals(left, right);
        }
    }
}