using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Utils.Extensions;

namespace Bing.Offices.Npoi.Factories
{
    /// <summary>
    /// 导出映射工厂
    /// </summary>
    internal static class ExportMappingFactory
    {
        /// <summary>
        /// 映射字典
        /// </summary>
        private static readonly IDictionary<Type, IDictionary<string, string>> MappingDict =
            new ConcurrentDictionary<Type, IDictionary<string, string>>();

        /// <summary>
        /// 创建映射字典实例
        /// </summary>
        /// <param name="type">导出类型</param>
        public static IDictionary<string, string> CreateInstance(Type type)
        {
            if (MappingDict.ContainsKey(type))
                return MappingDict[type];
            var dict = new Dictionary<string, string>();
            type.GetProperties().ForEach(p =>
            {
                if (p.IsDefined(typeof(ColumnNameAttribute)))
                    dict.Add(p.Name, p.GetCustomAttribute<ColumnNameAttribute>().Name);
                else
                    dict.Add(p.Name, p.Name);
            });
            MappingDict[type] = dict;
            return dict;
        }
    }
}
