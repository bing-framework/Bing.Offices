using System;
using System.Collections.Concurrent;
using Bing.Offices.Abstractions.Configurations;

namespace Bing.Offices.Internal
{
    /// <summary>
    /// 内部缓存
    /// </summary>
    internal static class InternalCache
    {
        /// <summary>
        /// Excel 导出配置字典
        /// </summary>
        internal static readonly ConcurrentDictionary<Type, IExcelConfiguration> ExcelExportConfigurationDictionary =
            new ConcurrentDictionary<Type, IExcelConfiguration>();
    }
}
