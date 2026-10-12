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

    /// <summary>
    /// 初始化一个 <see cref="BingOfficesConfigurationException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述配置错误的消息。</param>
    /// <param name="stage">发生配置错误的处理阶段。</param>
    /// <param name="sheetName">相关工作表名称。</param>
    /// <param name="rowIndex">相关单元格的一基行号。</param>
    /// <param name="columnIndex">相关单元格的一基列号。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    public BingOfficesConfigurationException(string message, BingOfficesStage stage,
        string sheetName, int rowIndex, int columnIndex, Exception innerException = null)
        : base(BingOfficesErrorCode.ConfigurationInvalid, BingOfficesOperation.Configuration,
            "Core", stage, message, innerException, sheetName, rowIndex, columnIndex)
    {
    }
}
