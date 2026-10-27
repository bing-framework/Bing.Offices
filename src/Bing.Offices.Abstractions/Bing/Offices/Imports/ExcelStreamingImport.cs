using System.ComponentModel;

namespace Bing.Offices.Imports;

/// <summary>构建没有 Workbook 目标集合的单 Sheet 流式请求。</summary>
public static class ExcelStreamingImport
{
    /// <summary>按名称选择工作表，使用既有 Sheet Fluent 配置。</summary>
    public static ExcelStreamingImportRequest<T> Sheet<T>(string name,
        Action<ExcelSheetImportBuilder<T>> configure = null) where T : class, new()
        => Sheet(ExcelSheetSelector.ByName(name), configure);
    /// <summary>按零基索引选择工作表。</summary>
    public static ExcelStreamingImportRequest<T> Sheet<T>(int index,
        Action<ExcelSheetImportBuilder<T>> configure = null) where T : class, new()
        => Sheet(ExcelSheetSelector.ByIndex(index), configure);
    private static ExcelStreamingImportRequest<T> Sheet<T>(ExcelSheetSelector selector,
        Action<ExcelSheetImportBuilder<T>> configure) where T : class, new()
    {
        var builder = new ExcelSheetImportBuilder<T>(selector, null);
        configure?.Invoke(builder);
        return new ExcelStreamingImportRequest<T>(builder.BuildSnapshot(null));
    }
}

/// <summary>独立的单 Sheet 导入配置快照，不绑定接收全量实体的集合。</summary>
public sealed class ExcelStreamingImportRequest<T> where T : class, new()
{
    internal ExcelStreamingImportRequest(ExcelSheetImportRequest sheet) => Sheet = sheet;
    /// <summary>提供程序使用的配置快照，Target 为 null。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelSheetImportRequest Sheet { get; }
}
