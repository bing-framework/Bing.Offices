namespace Bing.Offices.Exports;

/// <summary>
/// Workbook 元数据选项；实例只属于一个导出请求。
/// </summary>
public sealed class ExcelWorkbookMetadataOptions
{
    /// <summary>作者。</summary>
    public string Author { get; init; } = "简玄冰";

    /// <summary>公司。</summary>
    public string Company { get; init; } = "简玄冰";

    /// <summary>标题。</summary>
    public string Title { get; init; } = "Bing.Offices";

    /// <summary>主题。</summary>
    public string Subject { get; init; } = "Bing.Offices";

    /// <summary>类别。</summary>
    public string Category { get; init; } = "Bing.Offices";

    /// <summary>备注。</summary>
    public string Description { get; init; } = "Bing.Offices 生成";

    internal ExcelWorkbookMetadataOptions Clone() => new()
    {
        Author = Author,
        Company = Company,
        Title = Title,
        Subject = Subject,
        Category = Category,
        Description = Description
    };
}
