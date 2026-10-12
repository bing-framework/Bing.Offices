using System.ComponentModel;

namespace Bing.Offices.Entities;

/// <summary>
/// 描述实体列表的实际重复页眉。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityPageHeaderDefinition
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityPageHeaderDefinition" /> 类型的实例。
    /// </summary>
    internal ExcelEntityPageHeaderDefinition(ExcelEntityCellRange sourceRange,
        ExcelEntityCellReference marker, string markerText,
        IReadOnlyList<ExcelEntityPageHeaderCell> cells,
        IReadOnlyList<ExcelEntityCellRange> merges)
    {
        SourceRange = sourceRange;
        Marker = marker;
        MarkerText = markerText;
        Cells = Array.AsReadOnly(cells.ToArray());
        Merges = Array.AsReadOnly(merges.ToArray());
    }

    /// <summary>获取首页的源区域。</summary>
    public ExcelEntityCellRange SourceRange { get; }

    /// <summary>获取首页标记的绝对坐标。</summary>
    public ExcelEntityCellReference Marker { get; }

    /// <summary>获取规范化的 marker 文本。</summary>
    public string MarkerText { get; }

    /// <summary>获取相对覆盖单元格。</summary>
    public IReadOnlyList<ExcelEntityPageHeaderCell> Cells { get; }

    /// <summary>获取相对合并区域。</summary>
    public IReadOnlyList<ExcelEntityCellRange> Merges { get; }

    /// <summary>获取页眉占用的行数。</summary>
    public int RowCount => SourceRange.Last.Row - SourceRange.First.Row + 1;
}
