namespace Bing.Offices.Exceptions;

/// <summary>
/// 映射、Profile 或请求配置无效异常。
/// </summary>
public sealed class BingOfficesConfigurationException : BingOfficesException
{
    /// <summary>
    /// 初始化一个 <see cref="BingOfficesConfigurationException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述配置错误的消息。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    /// <param name="stage">发生配置错误的处理阶段。</param>
    public BingOfficesConfigurationException(string message, Exception innerException = null,
        BingOfficesStage stage = BingOfficesStage.Plan)
        : base(BingOfficesErrorCode.ConfigurationInvalid, BingOfficesOperation.Configuration,
            "Core", stage, message, innerException)
    {
    }
}

/// <summary>
/// 导入公共边界不可恢复失败异常。
/// </summary>
public sealed class BingOfficesImportException : BingOfficesException
{
    /// <summary>
    /// 初始化一个 <see cref="BingOfficesImportException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述导入失败的消息。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    /// <param name="provider">报告异常的提供程序名称。</param>
    /// <param name="stage">发生导入错误的处理阶段。</param>
    /// <param name="sheetName">相关工作表名称；未知时为 <see langword="null" />。</param>
    /// <param name="rowIndex">相关的一基行号；未知时为 <see langword="null" />。</param>
    /// <param name="columnIndex">相关的一基列号；未知时为 <see langword="null" />。</param>
    /// <param name="propertyName">相关模型属性名称；未知时为 <see langword="null" />。</param>
    /// <param name="code">导入失败的稳定错误码。</param>
    public BingOfficesImportException(string message, Exception innerException = null,
        string provider = "Core", BingOfficesStage stage = BingOfficesStage.Read,
        string sheetName = null, int? rowIndex = null, int? columnIndex = null,
        string propertyName = null, BingOfficesErrorCode code = BingOfficesErrorCode.ImportFailed)
        : base(code, BingOfficesOperation.Import, provider, stage, message, innerException,
            sheetName, rowIndex, columnIndex, propertyName)
    {
    }
}

/// <summary>
/// 导出公共边界不可恢复失败异常。
/// </summary>
public sealed class BingOfficesExportException : BingOfficesException
{
    /// <summary>
    /// 初始化一个 <see cref="BingOfficesExportException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述导出失败的消息。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    /// <param name="provider">报告异常的提供程序名称。</param>
    /// <param name="stage">发生导出错误的处理阶段。</param>
    /// <param name="sheetName">相关工作表名称；未知时为 <see langword="null" />。</param>
    /// <param name="rowIndex">相关的一基行号；未知时为 <see langword="null" />。</param>
    /// <param name="columnIndex">相关的一基列号；未知时为 <see langword="null" />。</param>
    /// <param name="propertyName">相关模型属性名称；未知时为 <see langword="null" />。</param>
    /// <param name="code">导出失败的稳定错误码。</param>
    public BingOfficesExportException(string message, Exception innerException = null,
        string provider = "Core", BingOfficesStage stage = BingOfficesStage.Write,
        string sheetName = null, int? rowIndex = null, int? columnIndex = null,
        string propertyName = null, BingOfficesErrorCode code = BingOfficesErrorCode.ExportFailed)
        : base(code, BingOfficesOperation.Export, provider, stage, message, innerException,
            sheetName, rowIndex, columnIndex, propertyName)
    {
    }
}

/// <summary>
/// 输入或输出资源预算超出异常。
/// </summary>
public sealed class BingOfficesResourceLimitException : BingOfficesException
{
    /// <summary>
    /// 初始化一个 <see cref="BingOfficesResourceLimitException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述资源限制失败的消息。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    /// <param name="provider">报告异常的提供程序名称。</param>
    /// <param name="operation">触发资源限制的操作类型。</param>
    /// <param name="stage">触发资源限制的处理阶段。</param>
    public BingOfficesResourceLimitException(string message, Exception innerException = null,
        string provider = "Core", BingOfficesOperation operation = BingOfficesOperation.Import,
        BingOfficesStage stage = BingOfficesStage.Preflight)
        : base(BingOfficesErrorCode.ResourceLimitExceeded, operation, provider, stage,
            message, innerException)
    {
    }
}

/// <summary>
/// 原子文件提交异常。
/// </summary>
public sealed class BingOfficesFileCommitException : BingOfficesException
{
    /// <summary>
    /// 初始化一个 <see cref="BingOfficesFileCommitException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述文件提交失败的消息。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    /// <param name="provider">报告异常的提供程序名称。</param>
    /// <param name="stage">发生提交错误的处理阶段。</param>
    public BingOfficesFileCommitException(string message, Exception innerException = null,
        string provider = "Core", BingOfficesStage stage = BingOfficesStage.Commit)
        : base(BingOfficesErrorCode.FileCommitFailed, BingOfficesOperation.FileCommit,
            provider, stage, message, innerException)
    {
    }
}

/// <summary>
/// 当前提供程序不支持请求功能异常。
/// </summary>
public sealed class BingOfficesUnsupportedFeatureException : BingOfficesException
{
    /// <summary>
    /// 初始化一个 <see cref="BingOfficesUnsupportedFeatureException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述不支持功能的消息。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    /// <param name="provider">报告异常的提供程序名称。</param>
    /// <param name="operation">触发不支持功能错误的操作类型。</param>
    /// <param name="stage">发生不支持功能错误的处理阶段。</param>
    public BingOfficesUnsupportedFeatureException(string message, Exception innerException = null,
        string provider = "Core", BingOfficesOperation operation = BingOfficesOperation.Import,
        BingOfficesStage stage = BingOfficesStage.Read)
        : base(BingOfficesErrorCode.UnsupportedFeature, operation, provider, stage,
            message, innerException)
    {
    }
}
