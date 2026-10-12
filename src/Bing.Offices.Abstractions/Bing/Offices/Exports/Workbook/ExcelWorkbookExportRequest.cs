using System.ComponentModel;
using Bing.Offices.Configurations;

namespace Bing.Offices.Exports;

/// <summary>
/// 工作簿级 Excel 导出请求。
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
    /// <param name="lookupLists">自动生成的文本字典。</param>
    /// <param name="cascadeLists">两级联动字典。</param>
    /// <param name="pivotTables">透视表定义。</param>
    internal ExcelWorkbookExportRequest(IReadOnlyList<ExcelSheetExportRequest> sheets, Stream template,
        bool leaveTemplateOpen, ExcelFormat format, ExcelWorkbookMetadataOptions metadata,
        bool metadataSpecified, IReadOnlyList<ExcelLookupListDefinition> lookupLists,
        IReadOnlyList<ExcelCascadeListDefinition> cascadeLists,
        IReadOnlyList<ExcelPivotTableDefinition> pivotTables)
        : this(sheets, template, leaveTemplateOpen, format, metadata, metadataSpecified, lookupLists,
            cascadeLists, pivotTables, false, Array.Empty<ExcelPivotTableDefinition>(), null)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelWorkbookExportRequest" /> 类型的实例。
    /// </summary>
    /// <param name="sheets">按输出顺序排列的工作表请求。</param>
    /// <param name="template">模板输入流。</param>
    /// <param name="leaveTemplateOpen">是否保持模板流打开。</param>
    /// <param name="format">目标文件格式。</param>
    /// <param name="metadata">工作簿元数据。</param>
    /// <param name="metadataSpecified">是否显式设置元数据。</param>
    /// <param name="lookupLists">文本字典。</param>
    /// <param name="cascadeLists">联动字典。</param>
    /// <param name="pivotTables">新建透视定义。</param>
    /// <param name="preserveTemplatePivotTables">是否保留已有透视。</param>
    /// <param name="pivotTableUpdates">已有透视的更新定义。</param>
    internal ExcelWorkbookExportRequest(IReadOnlyList<ExcelSheetExportRequest> sheets, Stream template,
        bool leaveTemplateOpen, ExcelFormat format, ExcelWorkbookMetadataOptions metadata,
        bool metadataSpecified, IReadOnlyList<ExcelLookupListDefinition> lookupLists,
        IReadOnlyList<ExcelCascadeListDefinition> cascadeLists,
        IReadOnlyList<ExcelPivotTableDefinition> pivotTables, bool preserveTemplatePivotTables,
        IReadOnlyList<ExcelPivotTableDefinition> pivotTableUpdates)
        : this(sheets, template, leaveTemplateOpen, format, metadata, metadataSpecified, lookupLists,
            cascadeLists, pivotTables, preserveTemplatePivotTables, pivotTableUpdates, null)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelWorkbookExportRequest" /> 类型的实例。
    /// </summary>
    /// <param name="sheets">按输出顺序排列的工作表请求。</param>
    /// <param name="template">模板输入流。</param>
    /// <param name="leaveTemplateOpen">是否保持模板流打开。</param>
    /// <param name="format">目标文件格式。</param>
    /// <param name="metadata">工作簿元数据。</param>
    /// <param name="metadataSpecified">是否显式设置元数据。</param>
    /// <param name="lookupLists">文本字典。</param>
    /// <param name="cascadeLists">联动字典。</param>
    /// <param name="pivotTables">新建透视定义。</param>
    /// <param name="preserveTemplatePivotTables">是否保留已有透视。</param>
    /// <param name="pivotTableUpdates">已有透视的更新定义。</param>
    /// <param name="resourceLimits">模板预检使用的资源限制。</param>
    internal ExcelWorkbookExportRequest(IReadOnlyList<ExcelSheetExportRequest> sheets, Stream template,
        bool leaveTemplateOpen, ExcelFormat format, ExcelWorkbookMetadataOptions metadata,
        bool metadataSpecified, IReadOnlyList<ExcelLookupListDefinition> lookupLists,
        IReadOnlyList<ExcelCascadeListDefinition> cascadeLists,
        IReadOnlyList<ExcelPivotTableDefinition> pivotTables, bool preserveTemplatePivotTables,
        IReadOnlyList<ExcelPivotTableDefinition> pivotTableUpdates,
        Bing.Offices.Imports.ExcelResourceLimits resourceLimits)
    {
        Sheets = Array.AsReadOnly(sheets.ToArray());
        Template = template;
        LeaveTemplateOpen = leaveTemplateOpen;
        Format = format;
        Metadata = metadata?.Clone() ?? new ExcelWorkbookMetadataOptions();
        MetadataSpecified = metadataSpecified;
        LookupLists = Array.AsReadOnly((lookupLists ?? Array.Empty<ExcelLookupListDefinition>()).ToArray());
        CascadeLists = Array.AsReadOnly((cascadeLists ?? Array.Empty<ExcelCascadeListDefinition>()).ToArray());
        PivotTables = Array.AsReadOnly((pivotTables ?? Array.Empty<ExcelPivotTableDefinition>())
            .Select(pivot => pivot.Clone()).ToArray());
        PreserveTemplatePivotTables = preserveTemplatePivotTables;
        PivotTableUpdates = Array.AsReadOnly((pivotTableUpdates ?? Array.Empty<ExcelPivotTableDefinition>())
            .Select(pivot => pivot.Clone()).ToArray());
        ResourceLimits = resourceLimits;
    }

    /// <summary>
    /// 获取请求中的工作表数量。
    /// </summary>
    public int SheetCount => Sheets.Count;

    /// <summary>
    /// 获取不可变工作表执行描述。
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

    /// <summary>
    /// 获取自动生成的文本字典。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelLookupListDefinition> LookupLists { get; }

    /// <summary>
    /// 获取两级联动字典。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelCascadeListDefinition> CascadeLists { get; }

    /// <summary>获取透视表定义。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelPivotTableDefinition> PivotTables { get; }

    /// <summary>
    /// 获取是否保留模板中的已有透视表。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool PreserveTemplatePivotTables { get; }

    /// <summary>
    /// 获取已有透视表的更新定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelPivotTableDefinition> PivotTableUpdates { get; }

    /// <summary>
    /// 获取模板预检使用的资源限制。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Bing.Offices.Imports.ExcelResourceLimits ResourceLimits { get; }
}
