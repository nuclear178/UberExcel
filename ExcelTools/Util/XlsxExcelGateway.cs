using System.Collections.Generic;
using ExcelTools.IO.Xlsx;
using OfficeOpenXml;

namespace ExcelTools.Util
{
    public class XlsxExcelGateway<TEntity> : IExcelGateway<TEntity> where TEntity : new()
    {
        private readonly ExcelPackage _package;
        private readonly ExcelWorksheet _worksheet;
        private readonly XlsxSerializer<TEntity> _serializer;

        public XlsxExcelGateway(ExcelPackage package, ExcelWorksheet worksheet, XlsxSerializer<TEntity> serializer)
        {
            _package = package;
            _worksheet = worksheet;
            _serializer = serializer;
        }

        public void Add(TEntity entity)
        {
            _serializer.SerializeObject(
                obj: entity,
                worksheet: _worksheet,
                rowIndex: _worksheet.Dimension?.Rows + 1 ?? 1
            );
        }

        public void AddRange(IEnumerable<TEntity> entities)
        {
            foreach (TEntity entity in entities)
            {
                Add(entity);
            }
        }

        public List<TEntity> ToList(int fromRow, int toRow)
        {
            var selection = new List<TEntity>();
            for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++)
            {
                selection.Add(_serializer.DeserializeObject(_worksheet, rowIndex));
            }

            return selection;
        }

        public List<TEntity> ToList(int fromRow)
        {
            return ToList(fromRow, _worksheet.Dimension.Rows);
        }

        public List<TEntity> ToList()
        {
            return ToList(1);
        }

        public void SaveChanges()
        {
            _package.Save();
        }
    }
}