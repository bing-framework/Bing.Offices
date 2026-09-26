using System.ComponentModel;

namespace Bing.Offices.Providers;

/// <summary>
/// Provider 声明的静态能力集合。
/// </summary>
[Flags]
[EditorBrowsable(EditorBrowsableState.Never)]
public enum ExcelProviderCapabilities
{
    /// <summary>
    /// 不声明任何 Excel Provider 能力。
    /// </summary>
    None = 0,
    /// <summary>
    /// 支持列表导入或导出操作。
    /// </summary>
    List = 1,
    /// <summary>
    /// 支持 Workbook 级多工作表操作。
    /// </summary>
    Workbook = 2,
    /// <summary>
    /// 支持 Entity 固定单元格和列表区域布局。
    /// </summary>
    Entity = 4,
    /// <summary>
    /// 支持基于既有工作簿模板的操作。
    /// </summary>
    Template = 8,
    /// <summary>
    /// 支持读取或写入合并区域。
    /// </summary>
    Merge = 16,
    /// <summary>
    /// 支持可取消的异步外围 IO。
    /// </summary>
    Async = 32,
    /// <summary>
    /// 支持 XLS/HSSF 工作簿格式。
    /// </summary>
    Xls = 64,
    /// <summary>
    /// 支持 XLSX/XSSF 工作簿格式。
    /// </summary>
    Xlsx = 128,
    /// <summary>
    /// 支持 XLSB 工作簿格式。
    /// </summary>
    Xlsb = 256,
    /// <summary>
    /// 支持 XLSM 工作簿格式。
    /// </summary>
    Xlsm = 512,
    /// <summary>
    /// 支持 ODS 工作簿格式。
    /// </summary>
    Ods = 1024
}

/// <summary>
/// Provider 的只读能力描述。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelProviderCapabilities
{
    /// <summary>
    /// 获取用于诊断和错误上下文的 Provider 稳定名称。
    /// </summary>
    string ProviderName { get; }
    /// <summary>
    /// 获取 Provider 静态声明的能力集合。
    /// </summary>
    ExcelProviderCapabilities Capabilities { get; }
    /// <summary>
    /// 判断 Provider 是否同时支持指定能力组合。
    /// </summary>
    /// <param name="capabilities">要检查的一个或多个能力标志。</param>
    /// <returns>当全部能力均受支持时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    bool Supports(ExcelProviderCapabilities capabilities);
}

/// <summary>
/// Provider 的方向化能力描述。
/// </summary>
/// <remarks>
/// 该接口是对历史位标志的可选补充，用于描述只读 Provider、分批导入和格式方向。
/// </remarks>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelProviderCapabilityDescriptor : IExcelProviderCapabilities
{
    /// <summary>
    /// 获取可读取的工作簿格式。
    /// </summary>
    IReadOnlyList<ExcelFormat> ReadFormats { get; }

    /// <summary>
    /// 获取可写入的工作簿格式。
    /// </summary>
    IReadOnlyList<ExcelFormat> WriteFormats { get; }

    /// <summary>
    /// 获取是否支持完整 Workbook 导入。
    /// </summary>
    bool SupportsCompleteWorkbookImport { get; }

    /// <summary>
    /// 获取是否支持按批次导入。
    /// </summary>
    bool SupportsBatchImport { get; }

    /// <summary>
    /// 获取是否支持完整 Workbook 导出。
    /// </summary>
    bool SupportsCompleteWorkbookExport { get; }

    /// <summary>
    /// 获取是否覆盖真实异步外围 IO。
    /// </summary>
    bool SupportsTrueAsyncIo { get; }

    /// <summary>
    /// 获取 Provider 的明确限制说明。
    /// </summary>
    IReadOnlyList<string> Limitations { get; }
}

/// <summary>
/// Provider 的可组合细粒度能力项。
/// </summary>
[Flags]
[EditorBrowsable(EditorBrowsableState.Never)]
public enum ExcelProviderFeatures
{
    /// <summary>
    /// 未声明细粒度能力。
    /// </summary>
    None = 0,
    /// <summary>
    /// 支持前向流式创建工作簿。
    /// </summary>
    StreamingWorkbookCreation = 1 << 0,
    /// <summary>
    /// 支持编辑模板工作簿。
    /// </summary>
    TemplateEditing = 1 << 1,
    /// <summary>
    /// 支持完整工作簿编辑。
    /// </summary>
    WorkbookEditing = 1 << 2,
    /// <summary>
    /// 支持 Excel 表格。
    /// </summary>
    Tables = 1 << 3,
    /// <summary>
    /// 支持自动筛选。
    /// </summary>
    AutoFilter = 1 << 4,
    /// <summary>
    /// 支持冻结窗格。
    /// </summary>
    FreezePanes = 1 << 5,
    /// <summary>
    /// 支持条件格式。
    /// </summary>
    ConditionalFormatting = 1 << 6,
    /// <summary>
    /// 支持名称范围。
    /// </summary>
    NamedRanges = 1 << 7,
    /// <summary>
    /// 支持打印布局。
    /// </summary>
    PrintLayout = 1 << 8,
    /// <summary>
    /// 支持读取公式文本。
    /// </summary>
    FormulaText = 1 << 9,
    /// <summary>
    /// 支持读取公式缓存值。
    /// </summary>
    FormulaCachedValues = 1 << 10,
    /// <summary>
    /// 支持重新计算公式。
    /// </summary>
    FormulaRecalculation = 1 << 11,
    /// <summary>
    /// 支持标记打开工作簿时重新计算。
    /// </summary>
    CalculateOnOpen = 1 << 12,
    /// <summary>
    /// 支持 PDF 渲染。
    /// </summary>
    PdfRendering = 1 << 13,
    /// <summary>
    /// 支持页面图片渲染。
    /// </summary>
    PageImageRendering = 1 << 14,
    /// <summary>
    /// 支持保留宏项目。
    /// </summary>
    MacroPreservation = 1 << 15,
    /// <summary>
    /// 支持 ODS 格式。
    /// </summary>
    Ods = 1 << 16,
    /// <summary>
    /// 支持读取加密工作簿。
    /// </summary>
    EncryptionRead = 1 << 17,
    /// <summary>
    /// 支持写入加密工作簿。
    /// </summary>
    EncryptionWrite = 1 << 18
}

/// <summary>
/// 可选的细粒度 Provider 能力描述。
/// </summary>
/// <remarks>未实现此接口的 Provider 继续使用历史能力契约。</remarks>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelProviderFeatureDescriptor : IExcelProviderCapabilityDescriptor
{
    /// <summary>
    /// 获取 Provider 声明的细粒度能力组合。
    /// </summary>
    ExcelProviderFeatures Features { get; }

    /// <summary>
    /// 判断是否同时支持指定的细粒度能力。
    /// </summary>
    /// <param name="features">要检查的能力组合。</param>
    /// <returns>全部支持时返回 <see langword="true"/>，否则返回 <see langword="false"/>。</returns>
    bool Supports(ExcelProviderFeatures features);
}
