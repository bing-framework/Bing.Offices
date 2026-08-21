namespace Bing.Offices.Exports;

/// <summary>
/// 与 Excel 提供程序无关的传统批注描述。
/// </summary>
public sealed class ExcelComment
{
    /// <summary>创建批注描述。</summary>
    public ExcelComment(string text, string author = null, bool visible = false)
    {
        Text = text ?? string.Empty;
        Author = author ?? string.Empty;
        Visible = visible;
    }

    /// <summary>获取批注文本。</summary>
    public string Text { get; }

    /// <summary>获取批注作者。</summary>
    public string Author { get; }

    /// <summary>获取批注是否可见。</summary>
    public bool Visible { get; }
}
