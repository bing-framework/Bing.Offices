using System.Reflection;
using Bing.Offices.Abstractions.Configurations;
using Bing.Offices.Abstractions.Contexts;
using Bing.Offices.Abstractions.Mappings;
using Bing.Offices.Internal;

namespace Bing.Offices.Contexts
{
    /// <summary>
    /// Excel 上下文
    /// </summary>
    public class ExcelContext : IExcelContext
    {
        /// <summary>
        /// 初始化
        /// </summary>
        public void Init()
        {
            var list = Bing.Utils.Helpers.Reflection.GetInstancesByInterface<IExcelExportMap>(
                Assembly.GetAssembly(typeof(IExcelExportMap)));
            foreach (var item in list)
            {
                item.Init();
            }
        }

        /// <summary>
        /// 获取导出设置
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        public IExcelConfiguration<TEntity> GetExportSetting<TEntity>() => InternalCache.ExcelExportConfigurationDictionary.GetOrAdd(typeof(TEntity),
            t => Helper.GetExcelConfigurationMapping<TEntity>()) as IExcelConfiguration<TEntity>;

    }
}
