namespace Bing.Offices.Exceptions;

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
