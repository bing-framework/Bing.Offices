namespace Bing.Offices.Exports;

/// <summary>
/// 动态列相对于固定列的请求级位置。
/// </summary>
public sealed class ExcelColumnPlacement
{
    /// <summary>初始化一个 <see cref="ExcelColumnPlacement" /> 类型的实例。</summary>
    /// <param name="beforeKey">要插入到其前方的固定列键。</param>
    /// <param name="afterKey">要插入到其后方的固定列键。</param>
    /// <param name="physicalColumnIndex">显式物理列索引；未指定时为 <see langword="null" />。</param>
    private ExcelColumnPlacement(string beforeKey, string afterKey, int? physicalColumnIndex)
    {
        BeforeKey = beforeKey;
        AfterKey = afterKey;
        PhysicalColumnIndex = physicalColumnIndex;
    }

    /// <summary>
    /// 获取位于该固定列之前的列键。
    /// </summary>
    public string BeforeKey { get; }

    /// <summary>
    /// 获取位于该固定列之后的列键。
    /// </summary>
    public string AfterKey { get; }

    /// <summary>
    /// 获取显式物理列索引。
    /// </summary>
    public int? PhysicalColumnIndex { get; }

    /// <summary>
    /// 创建相对于固定列之前的位置。
    /// </summary>
    /// <param name="key">作为插入参照的固定列键。</param>
    /// <returns>位于指定固定列之前的列位置。</returns>
    public static ExcelColumnPlacement Before(string key) => Create(key, null, null);

    /// <summary>
    /// 创建相对于固定列之后的位置。
    /// </summary>
    /// <param name="key">作为插入参照的固定列键。</param>
    /// <returns>位于指定固定列之后的列位置。</returns>
    public static ExcelColumnPlacement After(string key) => Create(null, key, null);

    /// <summary>
    /// 创建显式物理列索引位置。
    /// </summary>
    /// <param name="columnIndex">零基物理列索引。</param>
    /// <returns>位于指定物理列索引的列位置。</returns>
    public static ExcelColumnPlacement At(int columnIndex) => Create(null, null, columnIndex);

    /// <summary>根据相对键或物理索引创建列位置。</summary>
    /// <param name="beforeKey">要插入到其前方的固定列键。</param>
    /// <param name="afterKey">要插入到其后方的固定列键。</param>
    /// <param name="physicalColumnIndex">显式物理列索引。</param>
    /// <returns>验证通过的列位置定义。</returns>
    internal static ExcelColumnPlacement Create(string beforeKey, string afterKey, int? physicalColumnIndex)
    {
        if (physicalColumnIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(physicalColumnIndex));
        if (beforeKey != null && afterKey != null)
            throw new ArgumentException("动态列位置不能同时指定 Before 和 After。", nameof(beforeKey));
        if (beforeKey == null && afterKey == null && physicalColumnIndex == null)
            throw new ArgumentException("动态列位置必须指定 Before、After 或物理列索引。", nameof(beforeKey));
        if ((beforeKey != null || afterKey != null) && physicalColumnIndex != null)
            throw new ArgumentException("动态列位置不能同时指定相对位置和物理列索引。", nameof(physicalColumnIndex));
        return new ExcelColumnPlacement(beforeKey, afterKey, physicalColumnIndex);
    }
}
