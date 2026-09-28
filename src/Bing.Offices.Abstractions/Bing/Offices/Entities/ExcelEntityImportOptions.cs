using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 实体导入的资源限制选项。
/// </summary>
/// <remarks>
/// 该选项只影响实体导入，不改变实体布局的可复用性，也不影响实体导出。
/// 未提供限制时使用 <see cref="ExcelResourceLimits" /> 的默认值。
/// </remarks>
public sealed class ExcelEntityImportOptions
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityImportOptions" /> 类型的实例。
    /// </summary>
    /// <param name="resourceLimits">实体导入所使用的资源限制；为空时使用默认限制。</param>
    public ExcelEntityImportOptions(ExcelResourceLimits resourceLimits = null)
        : this(resourceLimits, ExcelValidationFailureMode.StopOnFirstFailure)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityImportOptions" /> 类型的实例。
    /// </summary>
    /// <param name="resourceLimits">实体导入所使用的资源限制；为空时使用默认限制。</param>
    /// <param name="validationFailureMode">列表行校验失败后的继续策略。</param>
    public ExcelEntityImportOptions(ExcelResourceLimits resourceLimits,
        ExcelValidationFailureMode validationFailureMode)
    {
        ResourceLimits = resourceLimits ?? new ExcelResourceLimits();
        ResourceLimits.Validate();
        if (!Enum.IsDefined(typeof(ExcelValidationFailureMode), validationFailureMode))
            throw new ArgumentOutOfRangeException(nameof(validationFailureMode));
        ValidationFailureMode = validationFailureMode;
    }

    /// <summary>
    /// 获取实体导入所使用的资源限制。
    /// </summary>
    public ExcelResourceLimits ResourceLimits { get; }

    /// <summary>
    /// 获取列表行校验失败后的处理策略。
    /// </summary>
    /// <remarks>
    /// 默认在当前行首次校验失败后停止继续处理该行；选择 <see cref="ExcelValidationFailureMode.Continue" />
    /// 时继续检查当前行的其余单元格。
    /// </remarks>
    public ExcelValidationFailureMode ValidationFailureMode { get; }
}
