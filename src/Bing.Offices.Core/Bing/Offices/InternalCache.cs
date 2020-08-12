using System;
using System.Collections.Concurrent;
using System.Reflection;
using Bing.Offices.FluentApi;
using Bing.Offices.Metadata;

namespace Bing.Offices
{
    /// <summary>
    /// 内部缓存
    /// </summary>
    internal static class InternalCache
    {
        /// <summary>
        /// 类型 - 属性信息 缓存
        /// </summary>
        public static readonly ConcurrentDictionary<Type, PropertyInfo[]> TypePropertyCache = new ConcurrentDictionary<Type, PropertyInfo[]>();

        /// <summary>
        /// 类型 - Excel设置 缓存
        /// </summary>
        public static readonly ConcurrentDictionary<Type, IExcelFluent> TypeFluentCache = new ConcurrentDictionary<Type, IExcelFluent>();

        /// <summary>
        /// 属性信息 - 委托 输出格式化函数 缓存
        /// </summary>
        public static readonly ConcurrentDictionary<PropertyInfo, Delegate> OutputFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, Delegate>();

        /// <summary>
        /// 属性信息 - 委托 输入格式化函数 缓存
        /// </summary>
        public static readonly ConcurrentDictionary<PropertyInfo, Delegate> InputFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, Delegate>();

        /// <summary>
        /// 属性信息 - 委托 列输入格式化函数 缓存
        /// </summary>
        public static readonly ConcurrentDictionary<PropertyInfo, Delegate> ColumnInputFormatterFuncCache = new ConcurrentDictionary<PropertyInfo, Delegate>();
    }
}
