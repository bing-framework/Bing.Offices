using Bing.Offices.FluentApi;

namespace Bing.Offices
{
    /// <summary>
    /// Fluent 设置
    /// </summary>
    public class FluentSettings
    {
        /// <summary>
        /// Fluent设置
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        public static IExcelFluent<TEntity> For<TEntity>() => InternalHelper.GetExcelFluentMapping<TEntity>();
    }
}
