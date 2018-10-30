using System.Collections.Generic;

namespace ExcelTools.Util
{
    public interface IExcelGateway<TEntity>
    {
        void Add(TEntity entity);

        void AddRange(IEnumerable<TEntity> entities);

        List<TEntity> ToList(int fromRow, int toRow);

        List<TEntity> ToList(int fromRow);

        List<TEntity> ToList();

        void SaveChanges();
    }
}