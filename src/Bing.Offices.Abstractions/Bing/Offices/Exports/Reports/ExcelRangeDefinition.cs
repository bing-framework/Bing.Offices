using System.ComponentModel;
using System.Text.RegularExpressions;

namespace Bing.Offices.Exports;

/// <summary>
/// 工作表区域定义。
/// </summary>
public sealed class ExcelRangeDefinition
{
    /// <summary>
    /// 获取或初始化零基起始行，包含。
    /// </summary>
    public int StartRow { get; init; }
    /// <summary>
    /// 获取或初始化零基起始列，包含。
    /// </summary>
    public int StartColumn { get; init; }
    /// <summary>
    /// 获取或初始化零基结束行，包含。
    /// </summary>
    public int EndRow { get; init; }
    /// <summary>
    /// 获取或初始化零基结束列，包含。
    /// </summary>
    public int EndColumn { get; init; }

    /// <summary>
    /// 验证区域边界。
    /// </summary>
    /// <param name="parameterName">验证失败时使用的参数名称。</param>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public void Validate(string parameterName)
    {
        if (StartRow < 0 || StartColumn < 0 || EndRow < StartRow || EndColumn < StartColumn)
            throw new ArgumentException("区域必须是非空的正向区域。", parameterName);
    }
}
