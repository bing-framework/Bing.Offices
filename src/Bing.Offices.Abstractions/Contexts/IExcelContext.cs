using Bing.Offices.Abstractions.Configurations;

namespace Bing.Offices.Abstractions.Contexts
{
    /// <summary>
    /// Excel 上下文
    /// </summary>
    public interface IExcelContext
    {
        /// <summary>
        /// 初始化
        /// </summary>
        void Init();

        /// <summary>
        /// 获取导出设置
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        IExcelConfiguration<TEntity> GetExportSetting<TEntity>();
    }
}
