using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Reflection;

namespace Bing.Offices.Internals;

/// <summary>
/// 内部帮助类
/// </summary>
internal static class InternalHelper
{
    /// <summary>
    /// 获取Excel配置映射
    /// </summary>
    /// <param name="entityType">实体类型</param>
    public static IExcelConfiguration GetExcelConfigurationMapping(Type entityType) =>
        InternalCache.TypeExcelConfigurationDictionary.GetOrAdd(entityType, type =>
        {
            var excelConfiguration = CreateExcelConfiguration(type,
                () => (ExcelConfiguration)Activator.CreateInstance(typeof(ExcelConfiguration<>).MakeGenericType(entityType)));
            return excelConfiguration;
        });

    /// <summary>
    /// 获取Excel配置映射
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    public static ExcelConfiguration<TEntity> GetExcelConfigurationMapping<TEntity>() =>
        (ExcelConfiguration<TEntity>) InternalCache.TypeExcelConfigurationDictionary.GetOrAdd(typeof(TEntity),
            type =>
            {
                var excelConfiguration = CreateExcelConfiguration(type,
                    () => new ExcelConfiguration<TEntity>());
                return excelConfiguration;
            });

    /// <summary>
    /// 创建Excel配置
    /// </summary>
    /// <param name="type">实体类型</param>
    /// <param name="newConfigurationFunc">创建Excel配置函数</param>
    private static ExcelConfiguration CreateExcelConfiguration(Type type, Func<ExcelConfiguration> newConfigurationFunc)
    {
        var excelConfiguration = newConfigurationFunc();
        var propertyInfos = TypeReflections.TypeCacheManager.GetTypeProperties(type);
        foreach (var propertyInfo in propertyInfos)
        {
            var propertyConfigurationType = typeof(PropertyConfiguration<,>).MakeGenericType(type, propertyInfo.PropertyType);
            var propertyConfiguration = (PropertyConfiguration)Activator.CreateInstance(propertyConfigurationType, propertyInfo);
            propertyConfiguration.Buid();
            excelConfiguration.PropertyConfigurationDictionary.Add(propertyInfo, propertyConfiguration);
        }
        return excelConfiguration;
    }

    /// <summary>
    /// 获取编码列名
    /// </summary>
    /// <param name="columnName">列名</param>
    public static string GetEncodedColumnName(string columnName) => $"{columnName}{InternalConst.DuplicateColumnMark}{Guid.NewGuid():N}";

    /// <summary>
    /// 获取解码列名
    /// </summary>
    /// <param name="columnName">列名</param>
    public static string GetDecodeColumnName(string columnName)
    {
        var duplicateMarkIndex = columnName.IndexOf(InternalConst.DuplicateColumnMark, StringComparison.OrdinalIgnoreCase);
        if (duplicateMarkIndex > 0)
            return columnName.Substring(0, duplicateMarkIndex);
        return columnName;
    }

    /// <summary>
    /// 是否有忽略特性
    /// </summary>
    /// <param name="property">属性信息</param>
    public static bool HasIgnore(PropertyInfo property)
    {
        if (property.IsDefined(typeof(NotMappedAttribute)))
            return true;
        if (property.IsDefined(typeof(ExcelIgnoreAttribute)))
            return true;
        return false;
    }
}
