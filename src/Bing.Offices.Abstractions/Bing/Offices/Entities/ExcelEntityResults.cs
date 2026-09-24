using System.ComponentModel;

namespace Bing.Offices.Entities;

/// <summary>
/// 单个聚合对象的导入结果。
/// </summary>
/// <typeparam name="TEntity">聚合对象类型。</typeparam>
public sealed class ExcelEntityImportResult<TEntity> where TEntity : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityImportResult{TEntity}" /> 类型的实例。
    /// </summary>
    /// <param name="entity">导入得到的实体。</param>
    /// <param name="errors">结构化导入错误。</param>
    /// <param name="sheets">各列表区域的来源结果。</param>
    /// <param name="errorsTruncated">是否因错误数量上限而截断。</param>
    /// <param name="maxErrors">生效的最大错误数。</param>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelEntityImportResult(TEntity entity, IReadOnlyList<Imports.ExcelImportError> errors,
        IReadOnlyList<Imports.ExcelSheetImportResult> sheets, bool errorsTruncated = false,
        int? maxErrors = null)
    {
        Entity = entity;
        Errors = errors ?? Array.Empty<Imports.ExcelImportError>();
        Sheets = sheets ?? Array.Empty<Imports.ExcelSheetImportResult>();
        ErrorsTruncated = errorsTruncated;
        MaxErrors = maxErrors;
    }

    /// <summary>
    /// 获取由固定单元格和列表区域组装的聚合对象。
    /// </summary>
    public TEntity Entity { get; }
    /// <summary>
    /// 获取固定单元格和列表区域产生的结构化错误。
    /// </summary>
    public IReadOnlyList<Imports.ExcelImportError> Errors { get; }
    /// <summary>
    /// 获取各列表区域的导入结果。
    /// </summary>
    public IReadOnlyList<Imports.ExcelSheetImportResult> Sheets { get; }
    /// <summary>
    /// 获取错误集合是否因达到数量上限而被截断。
    /// </summary>
    /// <remarks>
    /// 达到上限时为 <see langword="true" />，否则为 <see langword="false" />。
    /// </remarks>
    public bool ErrorsTruncated { get; }
    /// <summary>
    /// 获取生效的最大错误数；未限制时为 <see langword="null" />。
    /// </summary>
    public int? MaxErrors { get; }
    /// <summary>
    /// 获取本次导入是否成功。
    /// </summary>
    /// <remarks>
    /// 没有结构化错误时为 <see langword="true" />，否则为 <see langword="false" />。
    /// </remarks>
    public bool IsSuccess => Errors.Count == 0;
}

/// <summary>
/// 实体模板输入或输出选项。
/// </summary>
/// <remarks>
/// 导出时模板是输出工作簿的来源；导入时模板仅用于校验工作表和合并结构，实际数据仍从操作的 source 流读取。
/// 当 <see cref="LeaveOpen" /> 为 false 时，Provider 在操作结束后释放模板流。
/// </remarks>
public sealed class ExcelEntityTemplateOptions
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityTemplateOptions" /> 类型的实例。
    /// </summary>
    /// <param name="template">可读取的模板流。</param>
    /// <param name="leaveOpen">完成操作后是否保持模板流打开。</param>
    public ExcelEntityTemplateOptions(Stream template, bool leaveOpen = false)
    {
        if (template == null)
            throw new ArgumentNullException(nameof(template));
        if (!template.CanRead)
            throw new ArgumentException("模板流不可读取。", nameof(template));
        Template = template;
        LeaveOpen = leaveOpen;
    }

    /// <summary>
    /// 获取包含既有工作簿模板的可读流。
    /// </summary>
    public Stream Template { get; }
    /// <summary>
    /// 获取操作完成后是否由调用方继续持有模板流。
    /// </summary>
    public bool LeaveOpen { get; }
}
