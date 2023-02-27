using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection;
using Bing.Offices.Attributes;

namespace Bing.Offices.Npoi.Factories;

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
        foreach (var property in type.GetProperties())
        {
            if (HasIgnore(property))
                continue;
            if (property.IsDefined(typeof(ColumnNameAttribute)))
                dict.Add(property.Name, property.GetCustomAttribute<ColumnNameAttribute>().Name);
            else if (property.IsDefined(typeof(DynamicColumnAttribute)))
                dict.Add($"Dynamic:{property.Name}", string.Empty);
            else
                dict.Add(property.Name, property.Name);
        }
        MappingDict[type] = dict;
        return dict;
    }

    /// <summary>
    /// 是否有忽略属性
    /// </summary>
    /// <param name="propertyInfo">属性信息</param>
    private static bool HasIgnore(PropertyInfo propertyInfo)
    {
        if (propertyInfo.IsDefined(typeof(NotMappedAttribute)))
            return true;
        if (propertyInfo.IsDefined(typeof(ExcelIgnoreAttribute)))
            return true;
        return false;
    }
}