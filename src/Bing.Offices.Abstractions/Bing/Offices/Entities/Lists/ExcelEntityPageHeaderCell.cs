using System.ComponentModel;
using Bing.Offices.Styles;

namespace Bing.Offices.Entities;

/// <summary>
/// 描述重复页眉的覆盖单元格。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityPageHeaderCell
{
    /// <summary>保存覆盖值工厂。</summary>
    private readonly Func<int, int, IReadOnlyList<object>, object> _valueFactory;

    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityPageHeaderCell" /> 类型的实例。
    /// </summary>
    internal ExcelEntityPageHeaderCell(ExcelEntityCellReference reference,
        Func<int, int, IReadOnlyList<object>, object> valueFactory, ExcelCellStyle style,
        string numberFormat)
    {
        Reference = reference;
        _valueFactory = valueFactory;
        Style = style;
        NumberFormat = numberFormat;
    }

    /// <summary>获取相对于源区域左上角的单元格坐标。</summary>
    public ExcelEntityCellReference Reference { get; }

    /// <summary>获取单元格样式。</summary>
    public ExcelCellStyle Style { get; }

    /// <summary>获取数字格式。</summary>
    public string NumberFormat { get; }

    /// <summary>获取页眉单元格的值。</summary>
    public object GetValue(int pageNumber, int pageCount, IReadOnlyList<object> items) =>
        _valueFactory(pageNumber, pageCount, items);
}
