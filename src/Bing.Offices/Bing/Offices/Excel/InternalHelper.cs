using System;
using System.Collections.Generic;
using System.Reflection;
using Bing.Offices.Excel.Configurations;
using Bing.Offices.Excel.Settings;

namespace Bing.Offices.Excel
{
    /// <summary>
    /// 内部帮助类
    /// </summary>
    internal static class InternalHelper
    {
        /// <summary>
        /// 获取Excel映射配置
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        public static ExcelConfiguration<TEntity> GetExcelConfigurationMapping<TEntity>()
        {
            var configuration = (ExcelConfiguration<TEntity>) InternalCache.TypeExcelConcurrentDictionary.GetOrAdd(
                typeof(TEntity),
                type =>
                {
                    var excelConfiguration = new ExcelConfiguration<TEntity>();
                    var dic = new Dictionary<PropertyInfo, PropertyConfiguration>();
                    var propertyInfos = type.GetProperties();
                    foreach (var propertyInfo in propertyInfos)
                    {
                        var propertySettingType =
                            typeof(PropertySetting<,>).MakeGenericType(type, propertyInfo.PropertyType);
                        var propertySetting = Activator.CreateInstance(propertySettingType);

                        var propertyConfigurationType =
                            typeof(PropertyConfiguration<,>).MakeGenericType(type, propertyInfo.PropertyType);
                        var propertyConfiguration =
                            Activator.CreateInstance(propertyConfigurationType, new object[] {propertySetting});
                        dic.Add(propertyInfo, (PropertyConfiguration) propertyConfiguration);
                    }

                    excelConfiguration.PropertyConfigurationDictionary = dic;
                    return excelConfiguration;
                });
            return configuration;
        }
    }
}
