namespace Bing.Offices.Exports;

/// <summary>
/// Workbook 元数据选项；实例只属于一个导出请求。
/// </summary>
public sealed class ExcelWorkbookMetadataOptions
{
    /// <summary>获取或初始化作者名称，默认为“简玄冰”。</summary>
    public string Author { get; init; } = "简玄冰";

    /// <summary>获取或初始化公司名称，默认为“简玄冰”。</summary>
    public string Company { get; init; } = "简玄冰";

    /// <summary>获取或初始化工作簿标题，默认为“Bing.Offices”。</summary>
    public string Title { get; init; } = "Bing.Offices";

    /// <summary>获取或初始化工作簿主题，默认为“Bing.Offices”。</summary>
    public string Subject { get; init; } = "Bing.Offices";

    /// <summary>获取或初始化工作簿类别，默认为“Bing.Offices”。</summary>
    public string Category { get; init; } = "Bing.Offices";

    /// <summary>获取或初始化工作簿描述，默认为“Bing.Offices 生成”。</summary>
    public string Description { get; init; } = "Bing.Offices 生成";

    /// <summary>复制 Workbook 元数据选项。</summary>
    /// <returns>元数据选项的独立副本。</returns>
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
