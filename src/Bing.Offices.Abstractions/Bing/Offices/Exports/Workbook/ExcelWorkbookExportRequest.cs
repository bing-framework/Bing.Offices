using System.ComponentModel;
using Bing.Offices.Configurations;

namespace Bing.Offices.Exports;

/// <summary>
/// Workbook 级 Excel 导出请求。
/// </summary>
public sealed class ExcelWorkbookExportRequest
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelWorkbookExportRequest" /> 类型的实例。
    /// </summary>
    /// <param name="sheets">按输出顺序排列的工作表请求。</param>
    /// <param name="template">可选的模板输入流。</param>
    /// <param name="leaveTemplateOpen">是否由调用方继续持有模板流。</param>
    /// <param name="format">目标 Excel 文件格式。</param>
    /// <param name="metadata">工作簿元数据配置。</param>
    /// <param name="metadataSpecified">调用方是否显式设置过元数据。</param>
    internal ExcelWorkbookExportRequest(IReadOnlyList<ExcelSheetExportRequest> sheets, Stream template,
        bool leaveTemplateOpen, ExcelFormat format, ExcelWorkbookMetadataOptions metadata,
        bool metadataSpecified)
    {
        Sheets = sheets;
        Template = template;
        LeaveTemplateOpen = leaveTemplateOpen;
        Format = format;
        Metadata = metadata?.Clone() ?? new ExcelWorkbookMetadataOptions();
        MetadataSpecified = metadataSpecified;
    }

    /// <summary>
    /// 获取请求中的 Sheet 数量。
    /// </summary>
    public int SheetCount => Sheets.Count;

    /// <summary>
    /// 获取不可变 Sheet 执行描述。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelSheetExportRequest> Sheets { get; }

    /// <summary>
    /// 获取模板输入流。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Stream Template { get; }

    /// <summary>
    /// 获取模板流是否由调用方继续持有。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool LeaveTemplateOpen { get; }

    /// <summary>
    /// 获取目标 Excel 格式。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelFormat Format { get; }

    /// <summary>
    /// 获取工作簿元数据。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelWorkbookMetadataOptions Metadata { get; }

    /// <summary>
    /// 获取调用方是否显式设置过元数据。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool MetadataSpecified { get; }
}
