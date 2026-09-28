namespace Bing.Offices.Exceptions;

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
