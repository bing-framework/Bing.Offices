namespace Bing.Offices.Conversions;

/// <summary>
/// 打开工作簿的安全选项。
/// </summary>
public sealed class ExcelWorkbookOpenOptions
{
    /// <summary>
    /// 获取或初始化解密密码。
    /// </summary>
    /// <remarks>不得将密码写入日志。</remarks>
    public string Password { get; init; }
    /// <summary>
    /// 获取或初始化宏处理策略，默认拒绝。
    /// </summary>
    public ExcelMacroPolicy MacroPolicy { get; init; } = ExcelMacroPolicy.Reject;
}
