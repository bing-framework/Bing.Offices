namespace Bing.Offices.Imports;

/// <summary>
/// 基于 NPOI 的 Excel 导入器；输入会先复制并由 NPOI 建立内存中的 Workbook DOM。
/// </summary>
internal sealed class ExcelImportRuntime
{
    /// <summary>当前工作簿已尝试处理的数据行数量。</summary>
    private int _rowCount;
    /// <summary>指示行数上限错误是否已加入结果，避免重复报告。</summary>
    private bool _rowLimitReported;

    /// <summary>使用请求资源限制初始化工作簿级运行时状态。</summary>
    /// <param name="limits">导入请求配置的资源限制。</param>
    internal ExcelImportRuntime(ExcelResourceLimits limits)
    {
        MaxRows = limits?.MaxRows;
        ImageResources = new ExcelImageResourceTracker(limits);
    }

    /// <summary>获取允许处理的最大数据行数。</summary>
    internal int? MaxRows { get; }

    /// <summary>获取是否已达到数据行数上限。</summary>
    internal bool RowLimitReached => MaxRows.HasValue && _rowCount >= MaxRows.Value;

    /// <summary>获取跟踪工作簿图片数量和字节数的资源限制器。</summary>
    internal ExcelImageResourceTracker ImageResources { get; }

    /// <summary>尝试为一行数据消耗全局行数配额。</summary>
    /// <returns>成功消耗配额时为 true；已达到上限时为 false。</returns>
    internal bool TryConsumeRow()
    {
        if (MaxRows.HasValue && _rowCount >= MaxRows.Value)
            return false;
        _rowCount++;
        return true;
    }

    /// <summary>确保行数上限错误在单个工作簿中仅报告一次。</summary>
    /// <returns>本次调用首次标记上限错误时为 true。</returns>
    internal bool TryMarkRowLimitReported()
    {
        if (_rowLimitReported)
            return false;
        _rowLimitReported = true;
        return true;
    }
}

internal sealed class ExcelImageResourceTracker
{
    /// <summary>工作簿允许读取的最大图片数量。</summary>
    private readonly int? _maxPictures;
    /// <summary>单张图片允许占用的最大字节数。</summary>
    private readonly long? _maxPictureBytes;
    /// <summary>所有图片合计允许占用的最大字节数。</summary>
    private readonly long? _maxTotalPictureBytes;
    /// <summary>已接纳的图片数量。</summary>
    private int _count;
    /// <summary>已接纳图片的累计字节数。</summary>
    private long _totalBytes;

    /// <summary>从导入请求资源限制初始化图片配额跟踪器。</summary>
    /// <param name="limits">导入请求配置的资源限制。</param>
    internal ExcelImageResourceTracker(ExcelResourceLimits limits)
    {
        _maxPictures = limits?.MaxPictures;
        _maxPictureBytes = limits?.MaxPictureBytes;
        _maxTotalPictureBytes = limits?.MaxTotalPictureBytes;
    }

    /// <summary>验证并记录一张图片对工作簿资源配额的消耗。</summary>
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

internal sealed class ImageResourceLimitException : InvalidOperationException
{
    /// <summary>使用图片资源限制错误消息初始化异常。</summary>
    /// <param name="message">描述超出图片资源限制的消息。</param>
    internal ImageResourceLimitException(string message) : base(message)
    {
    }
}

internal sealed class SourceLocation
{
    /// <summary>使用导入实体的来源工作表和行号初始化位置。</summary>
    /// <param name="sheetName">实体所属的工作表名称。</param>
    /// <param name="rowIndex">实体所在的一基数据行号。</param>
    internal SourceLocation(string sheetName, int rowIndex)
    {
        SheetName = sheetName;
        RowIndex = rowIndex;
    }

    /// <summary>获取实体所属的工作表名称。</summary>
    internal string SheetName { get; }
    /// <summary>获取实体所在的一基数据行号。</summary>
    internal int RowIndex { get; }
}

/// <summary>按照对象引用而非对象值比较关联实体的比较器。</summary>
internal sealed class ReferenceObjectComparer : IEqualityComparer<object>
{
    /// <summary>获取可复用的引用比较器实例。</summary>
    internal static readonly ReferenceObjectComparer Instance = new();

    /// <inheritdoc />
    public new bool Equals(object x, object y) => ReferenceEquals(x, y);

    /// <inheritdoc />
    public int GetHashCode(object obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
}

#pragma warning restore CS0618
