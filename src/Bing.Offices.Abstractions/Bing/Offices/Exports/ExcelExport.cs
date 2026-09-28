namespace Bing.Offices.Exports;

/// <summary>
/// Workbook 导出请求构建入口。
/// </summary>
public static class ExcelExport
{
    /// <summary>
    /// 创建 Workbook 导出请求。
    /// </summary>
    /// <param name="configure">用于配置 Workbook 导出选项的委托。</param>
    /// <returns>已完成配置的 Workbook 导出请求。</returns>
    public static ExcelWorkbookExportRequest Workbook(Action<ExcelWorkbookExportBuilder> configure)
    {
        if (configure == null)
            throw new ArgumentNullException(nameof(configure));
        var builder = new ExcelWorkbookExportBuilder();
        configure(builder);
        return builder.Build();
    }
}
