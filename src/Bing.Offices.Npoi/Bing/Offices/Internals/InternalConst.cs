namespace Bing.Offices.Internals;

/// <summary>
/// 内部常量
/// </summary>
internal static class InternalConst
{
    /// <summary>
    /// Xls 工作簿允许的最大工作表数量（256 个）。
    /// </summary>
    public const int MaxSheetCountXls = 256;

    /// <summary>
    /// Xlsx 工作簿允许的最大工作表数量（16,384 个）。
    /// </summary>
    public const int MaxSheetCountXlsx = 16_384;

    /// <summary>
    /// Xls 工作表允许的最大行数（65,536 行）。
    /// </summary>
    public const int MaxRowCountXls = 65_536;

    /// <summary>
    /// Xlsx 工作表允许的最大行数（1,048,576 行）。
    /// </summary>
    public const int MaxRowCountXlsx = 1_048_576;

    /// <summary>
    /// NPOI 写入文档元数据的应用程序名称（Bing.Offices.Npoi）。
    /// </summary>
    public const string ApplicationName = "Bing.Offices.Npoi";
}
