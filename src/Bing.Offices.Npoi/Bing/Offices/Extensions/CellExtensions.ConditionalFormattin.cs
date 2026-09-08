using NPOI.SS.Util;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// NPOI单元格(<see cref="NPOI.SS.UserModel.ICell"/>) 扩展
/// </summary>
public static partial class CellExtensions
{
    #region AddConditionalFormattingRules(添加条件格式规则)

    /// <summary>
    /// 将条件格式规则应用到当前单元格对应的单元格区域。
    /// </summary>
    /// <param name="cell">单元格</param>
    /// <param name="cfrs">条件格式规则</param>
    public static void AddConditionalFormattingRules(this NPOI.SS.UserModel.ICell cell, NPOI.SS.UserModel.IConditionalFormattingRule[] cfrs)
    {
        var regions = new[]
        {
            new CellRangeAddress(cell.RowIndex, cell.RowIndex, cell.ColumnIndex, cell.ColumnIndex),
        };
        cell.Sheet.SheetConditionalFormatting.AddConditionalFormatting(regions, cfrs);
    }

    #endregion

    #region GetConditionalFormattingRules(获取条件格式规则)

    /// <summary>
    /// 获取覆盖当前单元格的所有条件格式规则。
    /// </summary>
    /// <param name="cell">单元格</param>
    /// <returns>按工作表条件格式顺序返回规则；没有匹配项时返回空数组。</returns>
    public static NPOI.SS.UserModel.IConditionalFormattingRule[] GetConditionalFormattingRules(this NPOI.SS.UserModel.ICell cell)
    {
        var cfrs = new List<NPOI.SS.UserModel.IConditionalFormattingRule>();
        var scf = cell.Sheet.SheetConditionalFormatting;
        for (var i = 0; i < scf.NumConditionalFormattings; i++)
        {
            var cf = scf.GetConditionalFormattingAt(i);
            if (cell.ExistConditionalFormatting(cf))
            {
                for (var j = 0; j < cf.NumberOfRules; j++)
                {
                    cfrs.Add(cf.GetRule(j));
                }
            }
        }

        return cfrs.ToArray();
    }

    /// <summary>
    /// 判断条件格式区域是否覆盖当前单元格。
    /// </summary>
    /// <param name="cell">单元格</param>
    /// <param name="cf">条件格式</param>
    /// <returns>条件格式区域覆盖当前单元格时为 true。</returns>
    private static bool ExistConditionalFormatting(this NPOI.SS.UserModel.ICell cell, NPOI.SS.UserModel.IConditionalFormatting cf)
    {
        return cf.GetFormattingRanges().Any(cra =>
            cell.RowIndex >= cra.FirstRow && cell.RowIndex <= cra.LastRow && cell.ColumnIndex >= cra.FirstColumn &&
            cell.ColumnIndex <= cra.LastColumn);
    }

    #endregion
}
