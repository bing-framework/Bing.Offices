namespace Bing.Offices.Exceptions;

/// <summary>
/// 错误码。
/// </summary>
public enum BingOfficesErrorCode
{
    /// <summary>
    /// 配置无效。
    /// </summary>
    ConfigurationInvalid,
    /// <summary>
    /// 导入失败。
    /// </summary>
    ImportFailed,
    /// <summary>
    /// 导出失败。
    /// </summary>
    ExportFailed,
    /// <summary>
    /// 资源限制超出。
    /// </summary>
    ResourceLimitExceeded,
    /// <summary>
    /// 文件提交失败。
    /// </summary>
    FileCommitFailed,
    /// <summary>
    /// 不支持的功能。
    /// </summary>
    UnsupportedFeature,
    /// <summary>
    /// 用户扩展执行失败。
    /// </summary>
    UserExtensionFailed
}

/// <summary>
/// 业务操作类型。
/// </summary>
public enum BingOfficesOperation
{
    /// <summary>
    /// 配置加载或解析。
    /// </summary>
    Configuration,
    /// <summary>
    /// 导入。
    /// </summary>
    Import,
    /// <summary>
    /// 导出。
    /// </summary>
    Export,
    /// <summary>
    /// 文件提交。
    /// </summary>
    FileCommit
}

/// <summary>
/// 业务操作阶段。
/// </summary>
public enum BingOfficesStage
{
    /// <summary>
    /// 打开输入。
    /// </summary>
    Open,
    /// <summary>
    /// 资源预检。
    /// </summary>
    Preflight,
    /// <summary>
    /// 配置解析或映射计划。
    /// </summary>
    Plan,
    /// <summary>
    /// 读取。
    /// </summary>
    Read,
    /// <summary>
    /// 转换或校验。
    /// </summary>
    Validate,
    /// <summary>
    /// 写入。
    /// </summary>
    Write,
    /// <summary>
    /// 序列化。
    /// </summary>
    Serialize,
    /// <summary>
    /// 提交。
    /// </summary>
    Commit,
    /// <summary>
    /// 清理。
    /// </summary>
    Cleanup
}

/// <summary>
/// 接收 Bing.Offices 公共运行异常的观察器。
/// </summary>
public interface IBingOfficesExceptionObserver
{
    /// <summary>
    /// 观察一个已经完成分类的公共异常。
    /// </summary>
    /// <param name="exception">公共异常。</param>
    void Observe(BingOfficesException exception);
}

/// <summary>
/// Bing.Offices 公共运行异常基类。
/// </summary>
public abstract class BingOfficesException : InvalidOperationException
{
    /// <summary>
    /// 初始化一个 <see cref="BingOfficesException" /> 类型的实例。
    /// </summary>
    /// <param name="code">稳定错误码。</param>
    /// <param name="operation">发生异常的操作类型。</param>
    /// <param name="provider">报告异常的提供程序名称。</param>
    /// <param name="stage">发生异常的处理阶段。</param>
    /// <param name="message">异常消息。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    /// <param name="sheetName">相关工作表名称；未知时为 <see langword="null" />。</param>
    /// <param name="rowIndex">相关的一基行号；未知时为 <see langword="null" />。</param>
    /// <param name="columnIndex">相关的一基列号；未知时为 <see langword="null" />。</param>
    /// <param name="propertyName">相关模型属性名称；未知时为 <see langword="null" />。</param>
    public BingOfficesException(BingOfficesErrorCode code, BingOfficesOperation operation,
        string provider, BingOfficesStage stage, string message, Exception innerException = null,
        string sheetName = null, int? rowIndex = null, int? columnIndex = null, string propertyName = null)
        : base(message, innerException)
    {
        Code = code;
        Operation = operation;
        Provider = provider;
        Stage = stage;
        SheetName = sheetName;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        PropertyName = propertyName;
    }

    /// <summary>
    /// 获取稳定错误码。
    /// </summary>
    public BingOfficesErrorCode Code { get; }

    /// <summary>
    /// 获取操作类型。
    /// </summary>
    public BingOfficesOperation Operation { get; }

    /// <summary>
    /// 获取提供程序名称。
    /// </summary>
    public string Provider { get; }

    /// <summary>
    /// 获取操作阶段。
    /// </summary>
    public BingOfficesStage Stage { get; }

    /// <summary>
    /// 获取工作表名称。
    /// </summary>
    public string SheetName { get; }

    /// <summary>
    /// 获取一基行号。
    /// </summary>
    public int? RowIndex { get; }

    /// <summary>
    /// 获取一基列号。
    /// </summary>
    public int? ColumnIndex { get; }

    /// <summary>
    /// 获取属性名称。
    /// </summary>
    public string PropertyName { get; }
}
