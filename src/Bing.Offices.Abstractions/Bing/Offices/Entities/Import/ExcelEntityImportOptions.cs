using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 实体导入的资源限制选项。
/// </summary>
/// <remarks>
/// 该选项只影响实体导入，不改变实体布局的可复用性，也不影响实体导出。
/// 未提供限制时使用 <see cref="ExcelResourceLimits" /> 的默认值。
/// 配置 <see cref="FailureOptions" /> 后，NPOI 和 ClosedXML 会在完整实体导入或 Buffered 批次完成后生成失败工作簿；
/// XLSX/HSSF Forward 批次支持 ErrorRowsOnly 和 AnnotatedOriginal，并在解析和回调完成后额外打开一次原始工作簿生成失败产物。
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
        : this(resourceLimits, validationFailureMode, ExcelImportValidationMode.ConfiguredRules,
            ExcelUnsupportedFeaturePolicy.Fail)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityImportOptions" /> 类型的实例。
    /// </summary>
    /// <param name="resourceLimits">实体导入所使用的资源限制；为空时使用默认限制。</param>
    /// <param name="validationFailureMode">列表行校验失败后的继续策略。</param>
    /// <param name="validationMode">配置规则与工作簿原生规则的选择。</param>
    /// <param name="unsupportedFeaturePolicy">原生规则不受支持时的处理策略。</param>
    public ExcelEntityImportOptions(ExcelResourceLimits resourceLimits,
        ExcelValidationFailureMode validationFailureMode, ExcelImportValidationMode validationMode,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy)
        : this(resourceLimits, validationFailureMode, validationMode, unsupportedFeaturePolicy, null)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityImportOptions" /> 类型的实例。
    /// </summary>
    /// <param name="resourceLimits">实体导入所使用的资源限制；为空时使用默认限制。</param>
    /// <param name="validationFailureMode">列表行校验失败后的继续策略。</param>
    /// <param name="validationMode">配置规则与工作簿原生规则的选择。</param>
    /// <param name="unsupportedFeaturePolicy">原生规则不受支持时的处理策略。</param>
    /// <param name="failureOptions">失败工作簿输出选项；为空或模式为 None 时不输出。</param>
    public ExcelEntityImportOptions(ExcelResourceLimits resourceLimits,
        ExcelValidationFailureMode validationFailureMode, ExcelImportValidationMode validationMode,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy,
        ExcelImportFailureOptions failureOptions)
    {
        ResourceLimits = resourceLimits ?? new ExcelResourceLimits();
        ResourceLimits.Validate();
        if (!Enum.IsDefined(typeof(ExcelValidationFailureMode), validationFailureMode))
            throw new ArgumentOutOfRangeException(nameof(validationFailureMode));
        if (!Enum.IsDefined(typeof(ExcelImportValidationMode), validationMode))
            throw new ArgumentOutOfRangeException(nameof(validationMode));
        if (!Enum.IsDefined(typeof(ExcelUnsupportedFeaturePolicy), unsupportedFeaturePolicy))
            throw new ArgumentOutOfRangeException(nameof(unsupportedFeaturePolicy));
        ValidationFailureMode = validationFailureMode;
        ValidationMode = validationMode;
        UnsupportedFeaturePolicy = unsupportedFeaturePolicy;
        FailureOptions = failureOptions;
        FailureOptions?.Validate();
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

    /// <summary>
    /// 获取导入规则来源策略。
    /// </summary>
    public ExcelImportValidationMode ValidationMode { get; }

    /// <summary>
    /// 获取不支持的工作簿原生规则处理策略。
    /// </summary>
    public ExcelUnsupportedFeaturePolicy UnsupportedFeaturePolicy { get; }

    /// <summary>
    /// 获取失败工作簿输出选项。
    /// </summary>
    public ExcelImportFailureOptions FailureOptions { get; }
}
