using Bing.Offices.Abstractions.Configurations;
using Bing.Offices.Abstractions.Mappings;
using Bing.Offices.Internal;

namespace Bing.Offices.Mappings
{
    /// <summary>
    /// Excel导出映射基类
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    public abstract class ExcelExportMapBase<TEntity> : IExcelExportMap<TEntity> where TEntity : class
    {
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="configuration">Excel配置</param>
        public abstract void Init(IExcelConfiguration<TEntity> configuration);

        /// <summary>
        /// 初始化
        /// </summary>
        public void Init()
        {
            Init(InternalContext.GetExportSetting<TEntity>());
        }
    }
}
