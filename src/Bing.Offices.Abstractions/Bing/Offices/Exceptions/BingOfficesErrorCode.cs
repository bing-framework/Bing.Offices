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
