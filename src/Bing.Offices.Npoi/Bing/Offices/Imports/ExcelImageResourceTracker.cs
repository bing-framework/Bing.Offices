namespace Bing.Offices.Imports;

/// <summary>
/// 跟踪工作簿图片资源的数量和字节配额。
/// </summary>
internal sealed class ExcelImageResourceTracker
{
    /// <summary>
    /// 记录工作簿允许读取的最大图片数量；为 null 时不限制数量。
    /// </summary>
    private readonly int? _maxPictures;
    /// <summary>
    /// 记录单张图片允许占用的最大字节数；为 null 时不限制单张大小。
    /// </summary>
    private readonly long? _maxPictureBytes;
    /// <summary>
    /// 记录所有图片合计允许占用的最大字节数；为 null 时不限制总大小。
    /// </summary>
    private readonly long? _maxTotalPictureBytes;
    /// <summary>
    /// 记录当前工作簿已接纳的图片数量。
    /// </summary>
    private int _count;
    /// <summary>
    /// 记录当前工作簿已接纳图片的累计字节数。
    /// </summary>
    private long _totalBytes;

    /// <summary>
    /// 初始化一个 <see cref="ExcelImageResourceTracker" /> 类型的实例。
    /// </summary>
    /// <param name="limits">导入请求配置的资源限制。</param>
    internal ExcelImageResourceTracker(ExcelResourceLimits limits)
    {
        _maxPictures = limits?.MaxPictures;
        _maxPictureBytes = limits?.MaxPictureBytes;
        _maxTotalPictureBytes = limits?.MaxTotalPictureBytes;
    }

    /// <summary>
    /// 验证并记录一张图片对工作簿资源配额的消耗。
    /// </summary>
    /// <param name="bytes">待接纳图片的字节数。</param>
    internal void Consume(long bytes)
    {
        if (_maxPictureBytes.HasValue && bytes > _maxPictureBytes.Value)
            throw new ImageResourceLimitException($"单张图片超过最大字节数: {_maxPictureBytes.Value}");
        if (_maxPictures.HasValue && _count >= _maxPictures.Value)
            throw new ImageResourceLimitException($"图片数量超过限制: {_maxPictures.Value}");
        if (_maxTotalPictureBytes.HasValue && bytes > _maxTotalPictureBytes.Value - _totalBytes)
            throw new ImageResourceLimitException($"图片总字节数超过限制: {_maxTotalPictureBytes.Value}");
        _count++;
        _totalBytes += bytes;
    }
}
