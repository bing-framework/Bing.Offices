using Bing.Offices.Configurations;

namespace Bing.Offices.Internal
{
    /// <summary>
    /// 内部帮助类
    /// </summary>
    internal static class Helper
    {
        /// <summary>
        /// 获取Excel配置映射
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        public static ExcelConfiguration<TEntity> GetExcelConfigurationMapping<TEntity>()
        {
            var type = typeof(TEntity);
            var excelConfiguration=new ExcelConfiguration<TEntity>()
            {

            };
            return excelConfiguration;
        }
    }
}
