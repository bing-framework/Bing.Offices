using System;
using Bing.Offices.Abstractions.Configurations;
using Bing.Offices.Internal;

namespace Bing.Offices.Exports
{
    /// <summary>
    /// 导出配置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    internal class ExportConfig<TEntity> where TEntity : class, new()
    {
        /// <summary>
        /// 导出设置
        /// </summary>
        internal IExcelConfiguration Setting => GetSetting();

        /// <summary>
        /// 查询条件
        /// </summary>
        internal object Condition { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ExportConfig{TEntity}"/>类型的实例
        /// </summary>
        public ExportConfig() { }

        /// <summary>
        /// 初始化一个<see cref="ExportConfig{TEntity}"/>类型的实例
        /// </summary>
        /// <param name="condition">查询条件</param>
        public ExportConfig(object condition)
        {
            Condition = condition;
        }

        /// <summary>
        /// 获取导出设置
        /// </summary>
        private IExcelConfiguration GetSetting()
        {
            var config = InternalContext.GetExportSetting<TEntity>();
            if (config == null)
                throw new ArgumentException($"请设置【{nameof(TEntity)}】导出配置");
            return config;
        }
    }
}
