using System.Collections.Generic;
using System.Reflection;

namespace ExcelTools.IO
{
    public class ObjectSchema
    {
        private readonly List<PropertyInfo> _columnsInfo;

        public ObjectSchema(List<PropertyInfo> columnsInfo)
        {
            _columnsInfo = columnsInfo;
        }
    }
}