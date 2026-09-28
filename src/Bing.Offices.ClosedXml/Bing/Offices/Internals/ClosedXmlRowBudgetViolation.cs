using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Xml;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.ClosedXml.Internals;

/// <summary>
/// ClosedXML 行预算预检命中的结构化描述。
/// </summary>
internal sealed class ClosedXmlRowBudgetViolation
{
    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlRowBudgetViolation" /> 类型的实例。
    /// </summary>
    /// <param name="sheetName">超限工作表名称。</param>
    /// <param name="itemType">工作表项目类型。</param>
    /// <param name="rowIndex">首次可能超限的行号。</param>
    /// <param name="maximumRows">允许的最大行数。</param>
    internal ClosedXmlRowBudgetViolation(string sheetName, Type itemType, int rowIndex,
        int maximumRows)
    {
        SheetName = sheetName;
        ItemType = itemType;
        RowIndex = rowIndex;
        MaximumRows = maximumRows;
    }

    /// <summary>
    /// 获取超限工作表名称。
    /// </summary>
    internal string SheetName { get; }

    /// <summary>
    /// 获取工作表项目类型。
    /// </summary>
    internal Type ItemType { get; }

    /// <summary>
    /// 获取首次可能超限的行号。
    /// </summary>
    internal int RowIndex { get; }

    /// <summary>
    /// 获取允许的最大行数。
    /// </summary>
    internal int MaximumRows { get; }
}
