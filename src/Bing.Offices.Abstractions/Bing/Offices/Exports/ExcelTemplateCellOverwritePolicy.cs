namespace Bing.Offices.Exports;

using System;

/// <summary>
/// 模板单元格写入策略。
/// </summary>
public enum ExcelTemplateCellOverwritePolicy
{
    /// <summary>
    /// 写入值时保留模板样式和批注；已有公式会被导出值替换。
    /// </summary>
    PreserveTemplate,
    /// <summary>
    /// 写入值前清除模板样式和批注；已有公式会被导出值替换。
    /// </summary>
    ReplaceTemplate
}
