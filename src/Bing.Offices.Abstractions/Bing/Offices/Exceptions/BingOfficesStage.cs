namespace Bing.Offices.Exceptions;

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
