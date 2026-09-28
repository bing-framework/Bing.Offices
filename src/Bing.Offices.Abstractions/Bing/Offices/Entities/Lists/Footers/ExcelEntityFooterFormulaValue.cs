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
/// 实体列表尾部的显式公式值。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityFooterFormulaValue
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityFooterFormulaValue" /> 类型的实例。
    /// </summary>
    /// <param name="formulaA1">不含等号的 Excel A1 公式文本。</param>
    internal ExcelEntityFooterFormulaValue(string formulaA1)
    {
        FormulaA1 = formulaA1 ?? throw new ArgumentNullException(nameof(formulaA1));
    }

    /// <summary>
    /// 获取不含等号的 Excel A1 公式文本。
    /// </summary>
    public string FormulaA1 { get; }
}
