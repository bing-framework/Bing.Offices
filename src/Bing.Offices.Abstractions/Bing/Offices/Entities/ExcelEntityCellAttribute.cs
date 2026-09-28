using System;

namespace Bing.Offices.Entities;

/// <summary>
/// 将实体公开属性绑定到实体布局中的固定单元格。
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class ExcelEntityCellAttribute : Attribute
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityCellAttribute" /> 类型的实例。
    /// </summary>
    /// <param name="sheetName">目标工作表名称。</param>
    /// <param name="address">目标单元格的 A1 地址。</param>
    public ExcelEntityCellAttribute(string sheetName, string address)
    {
        SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
        Address = address ?? throw new ArgumentNullException(nameof(address));
    }

    /// <summary>
    /// 获取目标工作表名称。
    /// </summary>
    public string SheetName { get; }

    /// <summary>
    /// 获取目标单元格的 A1 地址。
    /// </summary>
    public string Address { get; }

    /// <summary>
    /// 获取或初始化可选的命名转换器。
    /// </summary>
    public string ConverterName { get; init; }
}
