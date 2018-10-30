using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
// ReSharper disable UnusedMember.Global

namespace ExcelTools.Introspection.Mapping
{
    public class ObjectSchema : IEnumerable<ColumnOptions>
    {
        private readonly IEnumerable<ColumnOptions> _columns;
        private readonly Type _objType;

        public ObjectSchema(IEnumerable<ColumnOptions> columns, Type objType)
        {
            _columns = columns;
            _objType = objType;
        }

        public int ColumnMin => _columns.Select(column => column.Index).Min();

        public int ColumnMax => _columns.Select(column => column.Index).Max();

        public object Instantiate()
        {
            object instance = Activator.CreateInstance(_objType);

            object obj = instance;
            foreach (PropertyInfo info in _objType.GetProperties())
            {
                if (!info.PropertyType.IsClass) continue;
                if (info.GetValue(obj, null) == null) continue;
                
                Console.WriteLine("~~~");
            }

            return instance;
        }

        public IEnumerator<ColumnOptions> GetEnumerator()
        {
            return _columns.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}