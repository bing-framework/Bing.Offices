namespace Bing.Offices.Exports;

/// <summary>
/// 与 Excel 提供程序无关的传统批注描述。
/// </summary>
public sealed class ExcelComment
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelComment" /> 类型的实例。
    /// </summary>
    /// <param name="text">批注文本。</param>
    /// <param name="author">批注作者；未指定时为空字符串。</param>
    /// <param name="visible">是否在工作表中显示批注。</param>
    public ExcelComment(string text, string author = null, bool visible = false)
    {
        Text = text ?? string.Empty;
        Author = author ?? string.Empty;
        Visible = visible;
    }

    /// <summary>
    /// 获取批注文本。
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// 获取批注作者。
    /// </summary>
    public string Author { get; }

    /// <summary>
    /// 获取批注是否可见。
    /// </summary>
    public bool Visible { get; }
}
