namespace Bing.Offices.Exports;

/// <summary>
/// 验证 Excel 工作表名称。
/// </summary>
internal static class ExcelSheetNameValidation
{
    /// <summary>
    /// 验证工作表名称。
    /// </summary>
    /// <param name="name">工作表名称。</param>
    internal static void Validate(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("工作表名称不能为空。", nameof(name));
        if (name.Length > 31 || name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0
            || name[0] == '\'' || name[name.Length - 1] == '\''
            || name.IndexOfAny(new[] { '\0', '\u0003' }) >= 0)
            throw new ArgumentException($"工作表名称无效: {name}", nameof(name));
    }
}
