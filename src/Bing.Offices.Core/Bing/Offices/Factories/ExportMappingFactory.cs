using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.Extensions;
using Bing.Offices.Settings;

namespace Bing.Offices.Factories
{
    /// <summary>
    /// 导出映射工厂
    /// </summary>
    public static class ExportMappingFactory
    {
        /// <summary>
        /// 映射字典
        /// </summary>
        private static readonly IDictionary<Type, IDictionary<string,PropertySetting>> MappingDict =
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
        }
    }
}
