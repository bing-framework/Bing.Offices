using System;
using System.Collections.Concurrent;
using System.Reflection;
using Bing.Offices.Excel.Configurations;

namespace Bing.Offices.Excel
{
    /// <summary>
    /// 内部缓存
    /// </summary>
    public static class InternalCache
    {
        /// <summary>
        /// Excel配置字典
        /// </summary>
        public static readonly ConcurrentDictionary<Type, IExcelConfiguration> TypeExcelConcurrentDictionary = new ConcurrentDictionary<Type, IExcelConfiguration>();

        /// <summary>
        /// 输出格式化函数缓存
        /// </summary>
        public static readonly ConcurrentDictionary<PropertyInfo, object> OutputFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, object>();

        /// <summary>
        /// 输入格式化函数缓存
        /// </summary>
        public static readonly ConcurrentDictionary<PropertyInfo, object> InputFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, object>();
    }
}
