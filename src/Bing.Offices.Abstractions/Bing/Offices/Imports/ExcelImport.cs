using System.Linq.Expressions;

namespace Bing.Offices.Imports;

/// <summary>
/// Workbook 级强类型导入请求构建入口。
/// </summary>
public static class ExcelImport
{
    /// <summary>
    /// 创建 Workbook 导入请求。
    /// </summary>
    /// <typeparam name="TWorkbook">根 Workbook 模型类型。</typeparam>
    /// <param name="configure">用于配置 Workbook 导入选项的委托。</param>
    /// <returns>已完成配置的 Workbook 导入请求。</returns>
    public static ExcelWorkbookImportRequest<TWorkbook> Workbook<TWorkbook>(
        Action<ExcelWorkbookImportBuilder<TWorkbook>> configure)
        where TWorkbook : class, new()
    {
        if (configure == null)
            throw new ArgumentNullException(nameof(configure));
        var builder = new ExcelWorkbookImportBuilder<TWorkbook>();
        configure(builder);
        return builder.Build();
    }
}
