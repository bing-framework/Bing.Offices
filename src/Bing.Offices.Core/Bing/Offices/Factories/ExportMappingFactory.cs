using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.Exceptions;
using Bing.Offices.Extensions;
using Bing.Offices.Settings;
using Bing.Reflection;

namespace Bing.Offices.Factories;

/// <summary>
/// 导出映射工厂
/// </summary>
public static class ExportMappingFactory
{
    /// <summary>
    /// 映射字典
    /// </summary>
    private static readonly IDictionary<Type, IDictionary<string, PropertySetting>> MappingDict =
        new ConcurrentDictionary<Type, IDictionary<string, PropertySetting>>();

    /// <summary>
    /// 创建映射字典实例
    /// </summary>
    /// <typeparam name="T">导出类型</typeparam>
    public static IDictionary<string, PropertySetting> CreateInstance<T>() => CreateInstance(typeof(T));

    /// <summary>
    /// 创建映射字典实例
    /// </summary>
    /// <param name="type">导出类型</param>
    public static IDictionary<string, PropertySetting> CreateInstance(Type type)
    {
        if (MappingDict.TryGetValue(type, out var dict))
            return dict;
        dict = new Dictionary<string, PropertySetting>();
        foreach (var propertyInfo in type.GetProperties())
        {
            var setting = new PropertySetting();
            if (propertyInfo.HasIgnore())
                setting.Ignored = true;
            if (HasDynamicColumn(propertyInfo))
                setting.IsDynamicColumn = true;
            SetColumnName(propertyInfo, setting);
            SetFormatter(propertyInfo, setting);
            SetScale(propertyInfo, setting);
            SetValueMapping(propertyInfo, setting);
            dict[propertyInfo.Name] = setting;
        }
        MappingDict[type] = dict;
        return dict;
    }

    /// <summary>
    /// 设置列名
    /// </summary>
    /// <param name="propertyInfo">属性信息</param>
    /// <param name="setting">属性设置</param>
    private static void SetColumnName(PropertyInfo propertyInfo, PropertySetting setting)
    {
        var attribute = propertyInfo.GetCustomAttribute<ColumnNameAttribute>();
        if (attribute != null)
        {
            setting.Title = attribute.Name;
            return;
        }
        setting.Title = propertyInfo.Name;
    }

    /// <summary>
    /// 是否有动态列属性
    /// </summary>
    /// <param name="propertyInfo">属性信息</param>
    private static bool HasDynamicColumn(PropertyInfo propertyInfo)
    {
        if (propertyInfo.IsDefined(typeof(DynamicColumnAttribute)))
            return true;
        return false;
    }

    /// <summary>
    /// 设置格式化字符串
    /// </summary>
    /// <param name="propertyInfo">属性信息</param>
    /// <param name="setting">属性设置</param>
    private static void SetFormatter(PropertyInfo propertyInfo, PropertySetting setting)
    {
        var attribute = propertyInfo.GetCustomAttribute<DataFormatAttribute>();
        if (attribute != null)
            setting.Formatter = attribute.CustomFormat;
        var formatAttribute = propertyInfo.GetCustomAttribute<DisplayFormatAttribute>();
        if (formatAttribute != null)
            setting.Formatter = formatAttribute.DataFormatString;
    }

    /// <summary>
    /// 设置保留小数位数
    /// </summary>
    /// <param name="propertyInfo">属性信息</param>
    /// <param name="setting">属性设置</param>
    private static void SetScale(PropertyInfo propertyInfo, PropertySetting setting)
    {
        if (!Types.IsNumericType(Reflections.GetUnderlyingType(propertyInfo.PropertyType)))
            return;
        var attribute = propertyInfo.GetCustomAttribute<DecimalScaleAttribute>();
        if (attribute != null)
            setting.DecimalScale = attribute.Scale;
    }

    /// <summary>
    /// 设置值映射
    /// </summary>
    /// <param name="propertyInfo">属性信息</param>
    /// <param name="setting">属性设置</param>
    private static void SetValueMapping(PropertyInfo propertyInfo, PropertySetting setting)
    {
        var mappings = propertyInfo.GetCustomAttributes<ValueMappingAttribute>().ToList();
        if (!mappings.Any())
            return;
        if (setting.IsDynamicColumn)
            throw new OfficeException($"【{propertyInfo.Name}】该属性已设置动态列，无法再设置值映射");
        var dictionary = new Dictionary<string, object>();
        foreach (var mappingAttribute in mappings.Where(mappingAttribute => !dictionary.ContainsKey(mappingAttribute.Text)))
            dictionary.Add(mappingAttribute.Text, mappingAttribute.Value);
        // 如果存在自定义映射，则不会生成默认映射
        if (mappings.Any())
        {
            setting.MappingValues = dictionary;
            return;
        }
        // 为bool类型生成默认映射
        switch (propertyInfo.PropertyType.GetCSharpTypeName())
        {
            case "Boolean":
            case "Nullable<Boolean>":
            {
                if (!dictionary.ContainsKey(Resources.Yes))
                    dictionary.Add(Resources.Yes, true);
                if (!dictionary.ContainsKey(Resources.No))
                    dictionary.Add(Resources.No, false);
                setting.MappingValues = dictionary;
                break;
            }
        }

        var type = propertyInfo.PropertyType;
        var isNullable = Nullable.GetUnderlyingType(type) != null;
        if (isNullable)
            type = Nullable.GetUnderlyingType(type);
        // 为枚举类型生成默认映射
        if (type.IsEnum)
        {
            var values = type.GetEnumTextAndValues();
            foreach (var value in values)
                dictionary.Add(value.Key, value.Value);

            if (isNullable)
                if (!dictionary.ContainsKey(string.Empty))
                    dictionary.Add(string.Empty, null);
        }

        setting.MappingValues = dictionary;
    }
}