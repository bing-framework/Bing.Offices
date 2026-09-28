using System.Diagnostics;
using System.Globalization;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Imports;

/// <summary>
/// 按 Sheet 名称（忽略大小写）和物理行号标识失败工作簿候选行。
/// </summary>
internal sealed class ErrorRowComparer : IEqualityComparer<(string SheetName, int RowIndex)>
{
    /// <inheritdoc />
    public bool Equals((string SheetName, int RowIndex) left, (string SheetName, int RowIndex) right) =>
        left.RowIndex == right.RowIndex
        && string.Equals(left.SheetName, right.SheetName, StringComparison.OrdinalIgnoreCase);

    /// <inheritdoc />
    public int GetHashCode((string SheetName, int RowIndex) value) =>
        HashCode.Combine(StringComparer.OrdinalIgnoreCase.GetHashCode(value.SheetName ?? string.Empty), value.RowIndex);
}
