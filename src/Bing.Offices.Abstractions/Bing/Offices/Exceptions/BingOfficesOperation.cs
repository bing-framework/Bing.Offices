namespace Bing.Offices.Exceptions;

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
