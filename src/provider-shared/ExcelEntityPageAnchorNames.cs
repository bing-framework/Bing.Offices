using System.Globalization;
using Bing.Offices.Exceptions;

namespace Bing.Offices.Entities;

/// <summary>
/// 生成并识别逐页顺序命名锚点。
/// </summary>
internal static class ExcelEntityPageAnchorNames
{
    /// <summary>
    /// 生成指定逻辑页的名称。
    /// </summary>
    /// <param name="prefix">名称前缀。</param>
    /// <param name="pageIndex">一基逻辑页序号。</param>
    /// <returns>工作簿名称。</returns>
    internal static string Create(string prefix, int pageIndex)
    {
        if (prefix == null)
            return null;
        var name = prefix + "_" + pageIndex.ToString("D3", CultureInfo.InvariantCulture);
        if (pageIndex <= 0 || name.Length > 255)
            throw new BingOfficesConfigurationException("分页小计命名锚点超出 Excel 名称范围。",
                stage: BingOfficesStage.Plan);
        return name;
    }

    /// <summary>
    /// 判断名称是否属于指定前缀的分页序列。
    /// </summary>
    /// <param name="prefix">名称前缀。</param>
    /// <param name="name">待检查名称。</param>
    /// <returns>属于该分页序列时返回 <see langword="true" />。</returns>
    internal static bool IsSeriesName(string prefix, string name)
    {
        if (prefix == null || name == null || !name.StartsWith(prefix + "_",
            StringComparison.OrdinalIgnoreCase))
            return false;
        var suffix = name.Substring(prefix.Length + 1);
        return suffix.Length >= 3 && suffix.All(value => value >= '0' && value <= '9');
    }
}
