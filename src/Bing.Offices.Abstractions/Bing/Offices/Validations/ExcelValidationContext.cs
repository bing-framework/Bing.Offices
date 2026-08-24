namespace Bing.Offices.Validations;

using System.Globalization;
using Bing.Offices.Conversions;

/// <summary>
/// Excel 单元格校验上下文。
/// </summary>
public sealed class ExcelValidationContext
{
    /// <summary>
    /// 初始化一个<see cref="ExcelValidationContext"/>类型的实例。
    /// </summary>
    /// <param name="value">原始单元格文本。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">从一开始的行号。</param>
    /// <param name="columnIndex">从一开始的列号。</param>
    /// <param name="propertyName">目标属性名称。</param>
    /// <param name="convertedValue">转换后的属性值。</param>
    /// <param name="propertyType">目标属性类型。</param>
    /// <param name="cell">提供程序无关的原始单元格描述。</param>
    /// <param name="culture">当前请求的转换与校验区域性。</param>
    public ExcelValidationContext(string value, string sheetName, int rowIndex, int columnIndex, string propertyName,
        object convertedValue = null, Type propertyType = null, ExcelCellValue cell = null,
        CultureInfo culture = null)
    {
        Value = value ?? string.Empty;
        SheetName = sheetName;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        PropertyName = propertyName;
        ConvertedValue = convertedValue;
        PropertyType = propertyType;
        Cell = cell;
        Culture = culture ?? CultureInfo.InvariantCulture;
    }

    /// <summary>
    /// 获取原始单元格文本。
    /// </summary>
    public string Value { get; }

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
    /// 获取目标属性名称。
    /// </summary>
    public string PropertyName { get; }

    /// <summary>
    /// 获取转换后的属性值。
    /// </summary>
    public object ConvertedValue { get; }

    /// <summary>
    /// 获取目标属性类型。
    /// </summary>
    public Type PropertyType { get; }

    /// <summary>
    /// 获取提供程序无关的原始单元格描述；CSV 输入为文本单元格。
    /// </summary>
    public ExcelCellValue Cell { get; }

    /// <summary>
    /// 获取当前请求的转换与校验区域性。
    /// </summary>
    public CultureInfo Culture { get; }

}
