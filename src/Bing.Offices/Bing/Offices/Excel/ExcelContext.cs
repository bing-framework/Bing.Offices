using Bing.Offices.Excel.Configurations;

namespace Bing.Offices.Excel
{
    /// <summary>
    /// Excel 上下文
    /// </summary>
    public class ExcelContext
    {
        /// <summary>
        /// 获取实体配置
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        public IExcelConfiguration<TEntity> Entity<TEntity>() => InternalHelper.GetExcelConfigurationMapping<TEntity>();
    }
}
