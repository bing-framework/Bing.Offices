using Bing.Offices.Abstractions.Configurations;

namespace Bing.Offices.Internal
{
    /// <summary>
    /// 内部上下文
    /// </summary>
    internal static class InternalContext
    {
        /// <summary>
        /// 获取导出设置
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        public static IExcelConfiguration<TEntity> GetExportSetting<TEntity>() =>
            InternalCache.ExcelExportConfigurationDictionary.GetOrAdd(typeof(TEntity),
                t => Helper.GetExcelConfigurationMapping<TEntity>()) as IExcelConfiguration<TEntity>;
    }
}
