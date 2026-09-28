using System.ComponentModel;

namespace Bing.Offices.Providers;

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
