using System;
using System.Collections.Concurrent;
using System.Reflection;
using Bing.Offices.Configurations;

namespace Bing.Offices.Internals
{
    /// <summary>
    /// 内部缓存
    /// </summary>
    internal static class InternalCache
    {
        /// <summary>
        /// 类型 - Excel配置 字典
        /// </summary>
        public static readonly ConcurrentDictionary<Type, IExcelConfiguration> TypeExcelConfigurationDictionary = new ConcurrentDictionary<Type, IExcelConfiguration>();

        /// <summary>
        /// 输出格式化函数 缓存
        /// </summary>
        public static readonly ConcurrentDictionary<PropertyInfo, Delegate> OutputFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, Delegate>();

        /// <summary>
        /// 输入格式化函数 缓存
        /// </summary>
        public static readonly ConcurrentDictionary<PropertyInfo, Delegate> InputFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, Delegate>();

        /// <summary>
        /// 列输入格式化函数 缓存
        /// </summary>
        public static readonly ConcurrentDictionary<PropertyInfo, Delegate> ColumnInputFormatterFunCache = new ConcurrentDictionary<PropertyInfo, Delegate>();
    }
}
