using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using Bing.Offices.Exports;
using Bing.Offices.Styles;

namespace Bing.Offices.Entities;

/// <summary>
/// 最终尾部跨小计明细求和的 Provider SPI 值。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityFooterDetailSumValue
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityFooterDetailSumValue" /> 类型的实例。
    /// </summary>
    /// <param name="sourceColumn">待求和的工作表列字母。</param>
    internal ExcelEntityFooterDetailSumValue(string sourceColumn)
    {
        SourceColumn = sourceColumn ?? throw new ArgumentNullException(nameof(sourceColumn));
    }

    /// <summary>
    /// 获取待求和的工作表列字母。
    /// </summary>
    public string SourceColumn { get; }

    /// <summary>
    /// 根据一基明细行段生成不含等号的公式。
    /// </summary>
    /// <param name="detailRows">按写出顺序排列的一基闭区间明细行段。</param>
    /// <returns>不含等号的 Excel A1 公式。</returns>
    /// <remarks>
    /// 每段只包含真实明细行，间隔、小计和尾部行应由 Provider 排除。空明细生成 <c>0</c>。
    /// </remarks>
    public string ToFormulaA1(IReadOnlyList<(int FirstRow, int LastRow)> detailRows)
    {
        if (detailRows == null)
            throw new ArgumentNullException(nameof(detailRows));
        if (detailRows.Count == 0)
            return "0";
        var formula = new StringBuilder();
        var previousLastRow = 0;
        for (var index = 0; index < detailRows.Count; index++)
        {
            var (firstRow, lastRow) = detailRows[index];
            if (firstRow <= previousLastRow || lastRow < firstRow)
                throw new ArgumentException("明细行段必须是递增且不重叠的一基闭区间。", nameof(detailRows));
            if (index % 255 == 0)
            {
                if (index > 0)
                    formula.Append("+");
                formula.Append("SUM(");
            }
            else
                formula.Append(",");
            formula.Append(SourceColumn).Append(firstRow).Append(":")
                .Append(SourceColumn).Append(lastRow);
            if (index % 255 == 254 || index == detailRows.Count - 1)
                formula.Append(")");
            if (formula.Length > 8192)
                throw new ArgumentException("明细行段生成的公式超出 Excel 长度限制。", nameof(detailRows));
            previousLastRow = lastRow;
        }
        return formula.ToString();
    }
}
