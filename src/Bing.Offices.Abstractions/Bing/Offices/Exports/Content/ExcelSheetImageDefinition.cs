namespace Bing.Offices.Exports;

/// <summary>
/// 工作表嵌入图片定义。
/// </summary>
/// <remarks>支持 PNG/JPEG，采用零基单元格锚点，尺寸和偏移单位为像素。</remarks>
public sealed class ExcelSheetImageDefinition
{
    /// <summary>
    /// 获取或初始化图片编码内容。
    /// </summary>
    /// <remarks>构建请求时复制内容，不读取外部文件或网络资源。</remarks>
    public byte[] Content { get; init; }
    /// <summary>
    /// 获取或初始化零基行。
    /// </summary>
    public int Row { get; init; }
    /// <summary>
    /// 获取或初始化零基列。
    /// </summary>
    public int Column { get; init; }
    /// <summary>
    /// 获取或初始化输出宽度。
    /// </summary>
    public int Width { get; init; }
    /// <summary>
    /// 获取或初始化输出高度。
    /// </summary>
    public int Height { get; init; }
    /// <summary>
    /// 获取或初始化水平偏移。
    /// </summary>
    public int OffsetX { get; init; }
    /// <summary>
    /// 获取或初始化垂直偏移。
    /// </summary>
    public int OffsetY { get; init; }

    /// <summary>
    /// 验证图片内容签名和坐标。
    /// </summary>
    /// <remarks>图片编码完整性由图片读取器验证。</remarks>
    public void Validate()
    {
        if (Row < 0 || Column < 0 || Width <= 0 || Height <= 0 || OffsetX < 0 || OffsetY < 0)
            throw new ArgumentOutOfRangeException(nameof(Row), "图片坐标和偏移非负，尺寸为正数。");
        var png = Content != null && Content.Length >= 24 && Content[0] == 137 && Content[1] == 80
            && Content[2] == 78 && Content[3] == 71 && Content[4] == 13 && Content[5] == 10 && Content[6] == 26 && Content[7] == 10;
        var jpeg = Content != null && Content.Length >= 4 && Content[0] == 255 && Content[1] == 216 && Content[2] == 255;
        if (!png && !jpeg) throw new ArgumentException("图片必须是 PNG 或 JPEG 编码。", nameof(Content));
    }
}
