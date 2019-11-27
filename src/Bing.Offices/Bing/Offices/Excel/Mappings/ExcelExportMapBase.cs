using Bing.Offices.Excel.Configurations;

namespace Bing.Offices.Excel.Mappings
{
    /// <summary>
    /// Excel导出映射配置基类
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    public abstract class ExcelExportMapBase<TEntity> : IExcelExportMap where TEntity : class
    {
        /// <summary>
        /// 配置
        /// </summary>
        protected IExcelConfiguration<TEntity> Configuration { get; private set; }

        /// <summary>
        /// 映射配置
        /// </summary>
        /// <param name="context">上下文</param>
        public void Map(ExcelContext context)
        {
            Configuration = context.Entity<TEntity>();
            MapExport(Configuration);
        }

        /// <summary>
        /// 映射导出配置
        /// </summary>
        /// <param name="configuration">配置</param>
        protected abstract void MapExport(IExcelConfiguration<TEntity> configuration);
    }
}
