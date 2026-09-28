using System.ComponentModel;

namespace Bing.Offices.Providers;

/// <summary>
/// 可选的细粒度 Provider 能力描述。
/// </summary>
/// <remarks>未实现此接口的 Provider 继续使用历史能力契约。</remarks>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelProviderFeatureDescriptor : IExcelProviderCapabilityDescriptor
{
    /// <summary>
    /// 获取 Provider 声明的细粒度能力组合。
    /// </summary>
    ExcelProviderFeatures Features { get; }

    /// <summary>
    /// 判断是否同时支持指定的细粒度能力。
    /// </summary>
    /// <param name="features">要检查的能力组合。</param>
    /// <returns>全部支持时返回 <see langword="true"/>，否则返回 <see langword="false"/>。</returns>
    bool Supports(ExcelProviderFeatures features);
}
