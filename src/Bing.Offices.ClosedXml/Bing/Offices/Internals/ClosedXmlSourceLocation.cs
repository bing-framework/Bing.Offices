using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace Bing.Offices.ClosedXml.Internals;

/// <summary>
/// 记录 ClosedXML 前向导入实体的源工作表和一基行号。
/// </summary>
internal sealed class ClosedXmlSourceLocation
{
    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlSourceLocation" /> 类型的实例。
    /// </summary>
    /// <param name="sheetName">实体所属的工作表名称。</param>
    /// <param name="rowIndex">实体所在的一基行号。</param>
    internal ClosedXmlSourceLocation(string sheetName, int rowIndex)
    {
        SheetName = sheetName;
        RowIndex = rowIndex;
    }

    /// <summary>
    /// 获取实体所属的工作表名称。
    /// </summary>
    internal string SheetName { get; }

    /// <summary>
    /// 获取实体所在的一基行号。
    /// </summary>
    internal int RowIndex { get; }
}

/// <summary>
/// 按对象引用比较 ClosedXML 前向导入实体。
/// </summary>
internal sealed class ClosedXmlReferenceObjectComparer : IEqualityComparer<object>
{
    /// <summary>
    /// 获取可复用的引用比较器实例。
    /// </summary>
    internal static readonly ClosedXmlReferenceObjectComparer Instance = new();

    /// <inheritdoc />
    public new bool Equals(object x, object y) => ReferenceEquals(x, y);

    /// <inheritdoc />
    public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
}
