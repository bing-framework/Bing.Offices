using System;
using Bing.Offices.Configurations;

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
                () => (ExcelConfiguration)Activator.CreateInstance(
                    typeof(ExcelConfiguration<>).MakeGenericType(entityType)));
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
}