namespace Bing.Offices.Configurations;

/// <summary>
/// 映射配置加载期间产生的非阻断诊断。
/// </summary>
public sealed class ExcelMappingDiagnostic
{
    /// <summary>
    /// 初始化诊断信息。
    /// </summary>
    public ExcelMappingDiagnostic(string code, string path, string message)
    {
        Code = code;
        Path = path;
        Message = message;
    }

    /// <summary>获取诊断代码。</summary>
    public string Code { get; }

    /// <summary>获取配置路径。</summary>
    public string Path { get; }

    /// <summary>获取诊断消息。</summary>
    public string Message { get; }
}
