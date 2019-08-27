using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.Abstractions.Attributes;
using Bing.Offices.Abstractions.Filters;
using Bing.Offices.Attributes;

namespace Bing.Offices.Factories
{
    /// <summary>
    /// 类型过滤器信息工厂
    /// </summary>
    internal static class TypeFilterInfoFactory
    {
        /// <summary>
        /// 类型过滤器字典
        /// </summary>
        private static readonly IDictionary<Type, TypeFilterInfo> TypeFilterDict = new ConcurrentDictionary<Type, TypeFilterInfo>();

        /// <summary>
        /// 创建类型过滤器信息实例
        /// </summary>
        /// <param name="importType">导入类型</param>
        public static TypeFilterInfo CreateInstance(Type importType)
        {
            if (importType == null)
                throw new ArgumentNullException(nameof(importType));
            if (TypeFilterDict.ContainsKey(importType))
                return TypeFilterDict[importType];
            var typeFilterInfo = new TypeFilterInfo();
            var props = importType.GetProperties().Where(x => x.IsDefined(typeof(ColumnNameAttribute)));
            props.ToList().ForEach(item =>
            {
                typeFilterInfo.PropertyFilterInfos.Add(new PropertyFilterInfo()
                {
                    PropertyName = item.Name,
                    Filters = item.GetCustomAttributes<FilterAttributeBase>().ToList()
                });
            });
            TypeFilterDict[importType] = typeFilterInfo;
            return typeFilterInfo;
        }
    }
}
