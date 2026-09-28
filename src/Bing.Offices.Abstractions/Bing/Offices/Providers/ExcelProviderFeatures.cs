using System.ComponentModel;

namespace Bing.Offices.Providers;

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
