using Bing.Offices.Attributes;
using Bing.Offices.Filters;
using Bing.Offices.Metadata.Excels;

namespace Bing.Offices.Extensions;

/// <summary>
/// 单元格(<see cref="ICell"/>) 扩展
/// </summary>
public static class CellExtensions
{
    /// <summary>
    /// 获取单元格的过滤器特性
    /// </summary>
    /// <typeparam name="T">过滤器特性类型</typeparam>
    /// <param name="cell">单元格</param>
    /// <param name="typeFilterInfo">类型过滤器信息</param>
    public static T GetFilterAttribute<T>(this ICell cell, TypeFilterInfo typeFilterInfo) where T : FilterAttributeBase
    {
        return typeFilterInfo.PropertyFilterInfos
            .SingleOrDefault(x =>
                x.PropertyName.Equals(cell.PropertyName, StringComparison.CurrentCultureIgnoreCase))?.Filters
            ?.SingleOrDefault(x => x.GetType() == typeof(T)) as T;
    }

    /// <summary>
    /// 获取单元格的过滤器特性集合
    /// </summary>
    /// <typeparam name="T">过滤器特性类型</typeparam>
    /// <param name="cell">单元格</param>
    /// <param name="typeFilterInfo">类型过滤器信息</param>
    public static IList<T> GetFilterAttributes<T>(this ICell cell, TypeFilterInfo typeFilterInfo)
        where T : FilterAttributeBase
    {
        return typeFilterInfo.PropertyFilterInfos
            .SingleOrDefault(x =>
                x.PropertyName.Equals(cell.PropertyName, StringComparison.CurrentCultureIgnoreCase))?.Filters
            ?.Where(x => x.GetType() == typeof(T)).Cast<T>().ToList();
    }

    /// <summary>
    /// 是否日期
    /// </summary>
    /// <param name="cell">单元格</param>
    public static bool IsDateTime(this ICell cell) => DateTime.TryParse(cell.Value.ToString(), out var time);

    /// <summary>
    /// 是否在范围内
    /// </summary>
    /// <param name="cell">单元格</param>
    /// <param name="min">最小值</param>
    /// <param name="max">最大值</param>
    public static bool IsInRange(this ICell cell, decimal min, decimal max)
    {
        if (!decimal.TryParse(cell.Value.ToString(), out decimal val))
            return false;
        if (val > max || val < min)
            return false;
        return true;
    }
}
