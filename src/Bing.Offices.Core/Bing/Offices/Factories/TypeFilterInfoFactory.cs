using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.Filters;

namespace Bing.Offices.Factories;

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
    /// <param name="type">类型</param>
    public static TypeFilterInfo CreateInstance(Type type)
    {
        if (type == null)
            throw new ArgumentNullException(nameof(type));
        if (TypeFilterDict.ContainsKey(type))
            return TypeFilterDict[type];
        var typeFilterInfo = new TypeFilterInfo();
        var props = type.GetProperties().Where(x => x.IsDefined(typeof(ColumnNameAttribute))).ToList();
        props.ForEach(item =>
        {
            typeFilterInfo.PropertyFilterInfos.Add(new PropertyFilterInfo()
            {
                PropertyName = item.Name,
                Filters = item.GetCustomAttributes<FilterAttributeBase>().ToList()
            });
        });
        TypeFilterDict[type] = typeFilterInfo;
        return typeFilterInfo;
    }
}