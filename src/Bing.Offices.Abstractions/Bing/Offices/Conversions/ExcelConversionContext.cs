namespace Bing.Offices.Conversions;

using System.Globalization;

/// <summary>
/// Excel 属性值转换上下文。
/// </summary>
public sealed class ExcelConversionContext
{
    /// <summary>
    /// 初始化一个<see cref="ExcelConversionContext"/>类型的实例。
    /// </summary>
    /// <param name="value">当前原始值。</param>
    /// <param name="propertyName">目标属性名称。</param>
    /// <param name="propertyType">目标属性类型。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">从一开始的行号。</param>
    /// <param name="columnIndex">从一开始的列号。</param>
    /// <param name="culture">值转换使用的区域性。</param>
    /// <param name="cell">与提供程序无关的单元格值描述。</param>
    public ExcelConversionContext(object value, string propertyName, Type propertyType, string sheetName, int rowIndex,
        int columnIndex, CultureInfo culture = null, ExcelCellValue cell = null)
    {
        Value = value;
        PropertyName = propertyName ?? throw new ArgumentNullException(nameof(propertyName));
        PropertyType = propertyType ?? throw new ArgumentNullException(nameof(propertyType));
        SheetName = sheetName;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Culture = culture ?? CultureInfo.InvariantCulture;
        Cell = cell;
    }

    /// <summary>
    /// 获取当前原始值。
    /// </summary>
    public object Value { get; }

    /// <summary>
    /// 获取目标属性名称。
    /// </summary>
    public string PropertyName { get; }

    /// <summary>
    /// 获取目标属性类型。
    /// </summary>
    public Type PropertyType { get; }

    /// <summary>
    /// 获取工作表名称。
    /// </summary>
    public string SheetName { get; }

    /// <summary>
    /// 获取从一开始的行号。
    /// </summary>
    public int RowIndex { get; }

    /// <summary>
    /// 获取从一开始的列号。
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// 获取值转换使用的区域性。
    /// </summary>
    public CultureInfo Culture { get; }

    /// <summary>
    /// 获取与提供程序无关的单元格值描述。CSV 等纯文本来源可能为 <see langword="null"/>。
    /// </summary>
    public ExcelCellValue Cell { get; }
}
