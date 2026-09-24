namespace Bing.Offices;

/// <summary>
/// 内部常量。
/// </summary>
internal static class InternalConst
{
    /// <summary>
    /// Core 映射允许处理的最大工作表数量（256 个）。
    /// </summary>
    internal const int MaxSheetNum = 256;

    /// <summary>
    /// 基本类型映射使用的默认属性名称（Value）。
    /// </summary>
    internal const string DefaultPropertyNameForBasicType = "Value";

    /// <summary>
    /// Core 程序集写入文档元数据的应用名称（Bing.Offices）。
    /// </summary>
    internal const string ApplicationName = "Bing.Offices";

    /// <summary>
    /// 编码重复列名时使用的内部标记（__dup_mark__）。
    /// </summary>
    public const string DuplicateColumnMark = "__dup_mark__";
}
