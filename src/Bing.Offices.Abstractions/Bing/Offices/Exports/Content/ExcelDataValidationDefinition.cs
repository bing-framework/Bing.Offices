namespace Bing.Offices.Exports;

/// <summary>
/// 有界区域的 Excel 原生数据校验定义。
/// </summary>
/// <remarks>
/// 调用方不能提供任意公式或外部引用；级联列表公式由 Provider 生成。
/// </remarks>
public sealed class ExcelDataValidationDefinition
{
    /// <summary>
    /// 获取或初始化应用区域。
    /// </summary>
    /// <remarks>区域坐标为零基，且包含结束位置。</remarks>
    public ExcelRangeDefinition Range { get; init; }
    /// <summary>
    /// 获取或初始化校验类型。
    /// </summary>
    public ExcelDataValidationType Type { get; init; }
    /// <summary>
    /// 获取或初始化比较运算符。
    /// </summary>
    /// <remarks>列表校验不使用比较运算符。</remarks>
    public ExcelConditionalComparisonOperator Operator { get; init; }
    /// <summary>
    /// 获取或初始化数字或长度的第一个边界。
    /// </summary>
    public double Value1 { get; init; }
    /// <summary>
    /// 获取或初始化第二个边界。
    /// </summary>
    /// <remarks>仅 Between 和 NotBetween 运算使用该边界。</remarks>
    public double Value2 { get; init; }
    /// <summary>
    /// 获取或初始化日期第一个边界。
    /// </summary>
    public DateTime? Date1 { get; init; }
    /// <summary>
    /// 获取或初始化日期第二个边界。
    /// </summary>
    public DateTime? Date2 { get; init; }
    /// <summary>
    /// 获取或初始化显式列表项。
    /// </summary>
    /// <remarks>列表项不能包含逗号、引号或换行。</remarks>
    public IReadOnlyList<string> Values { get; init; } = Array.Empty<string>();
    /// <summary>
    /// 获取或初始化列表校验使用的工作簿名称。
    /// </summary>
    public string ListSourceName { get; init; }
    /// <summary>
    /// 获取或初始化级联列表名称。
    /// </summary>
    public string CascadeSourceName { get; init; }
    /// <summary>
    /// 获取或初始化同一行上级单元格的零基列号。
    /// </summary>
    public int? ParentColumnIndex { get; init; }
    /// <summary>
    /// 获取或初始化是否允许空单元格。
    /// </summary>
    public bool IgnoreBlanks { get; init; } = true;
    /// <summary>
    /// 获取或初始化是否显示停止型错误提示。
    /// </summary>
    public bool ShowErrorMessage { get; init; } = true;
    /// <summary>
    /// 获取或初始化输入提示标题。
    /// </summary>
    public string InputTitle { get; init; }
    /// <summary>
    /// 获取或初始化输入提示正文。
    /// </summary>
    public string InputMessage { get; init; }
    /// <summary>
    /// 获取或初始化错误提示标题。
    /// </summary>
    public string ErrorTitle { get; init; }
    /// <summary>
    /// 获取或初始化错误提示正文。
    /// </summary>
    public string ErrorMessage { get; init; }

    /// <summary>
    /// 验证跨 Provider 共用的有限边界合同。
    /// </summary>
    public void Validate()
    {
        if (Range == null) throw new ArgumentException("数据校验必须指定区域。", nameof(Range));
        Range.Validate(nameof(Range));
        if (!Enum.IsDefined(typeof(ExcelDataValidationType), Type)) throw new ArgumentOutOfRangeException(nameof(Type));
        if (!Enum.IsDefined(typeof(ExcelConditionalComparisonOperator), Operator)) throw new ArgumentOutOfRangeException(nameof(Operator));
        if ((InputTitle?.Length ?? 0) > 32 || (ErrorTitle?.Length ?? 0) > 32
            || (InputMessage?.Length ?? 0) > 255 || (ErrorMessage?.Length ?? 0) > 255)
            throw new ArgumentException("校验标题最多 32 字符，正文最多 255 字符。");
        var between = Operator == ExcelConditionalComparisonOperator.Between || Operator == ExcelConditionalComparisonOperator.NotBetween;
        if (Type == ExcelDataValidationType.List)
        {
            var hasName = ListSourceName != null;
            var hasValues = Values?.Count > 0;
            var hasCascade = CascadeSourceName != null;
            if ((hasName ? 1 : 0) + (hasValues ? 1 : 0) + (hasCascade ? 1 : 0) != 1)
                throw new ArgumentException("列表必须且只能指定一种值来源。");
            if (hasName)
                ExcelLookupListDefinition.ValidateName(ListSourceName);
            else if (hasCascade)
            {
                ExcelLookupListDefinition.ValidateName(CascadeSourceName);
                if (!ParentColumnIndex.HasValue || ParentColumnIndex.Value < 0
                    || Range.StartColumn != Range.EndColumn
                    || ParentColumnIndex.Value == Range.StartColumn)
                    throw new ArgumentException("级联列表必须指定区域外的同一行上级列。", nameof(ParentColumnIndex));
            }
            else if (Values.Any(value => string.IsNullOrEmpty(value)
                || value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0) || string.Join(",", Values).Length > 253)
                throw new ArgumentException("显式列表不能包含分隔符，编码后含引号最多 255 字符。", nameof(Values));
            if (!hasCascade && ParentColumnIndex.HasValue)
                throw new ArgumentException("只有级联列表可以指定上级列。", nameof(ParentColumnIndex));
            return;
        }
        if (ListSourceName != null || CascadeSourceName != null || ParentColumnIndex.HasValue)
            throw new ArgumentException("只有列表校验可以指定列表来源和上级列。");
        if (Type == ExcelDataValidationType.Date)
        {
            if (!Date1.HasValue || between && !Date2.HasValue) throw new ArgumentException("日期校验缺少边界。");
            foreach (var date in new[] { Date1, between ? Date2 : null }.Where(value => value.HasValue))
                if (date.Value < new DateTime(1900, 1, 1) || date.Value.TimeOfDay != TimeSpan.Zero)
                    throw new ArgumentException("日期边界必须是 1900 年起的日期且不含时间。");
            if (between && Date1 > Date2) throw new ArgumentException("日期边界顺序错误。");
            return;
        }
        foreach (var value in between ? new[] { Value1, Value2 } : new[] { Value1 })
        {
            if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentException("校验边界必须有限。");
            if ((Type == ExcelDataValidationType.Integer || Type == ExcelDataValidationType.TextLength)
                && (value != Math.Truncate(value) || value < int.MinValue || value > int.MaxValue))
                throw new ArgumentException("整数和长度边界必须是 Int32 整数。");
            if (Type == ExcelDataValidationType.TextLength && value < 0) throw new ArgumentException("长度不能为负数。");
        }
        if (between && Value1 > Value2) throw new ArgumentException("校验边界顺序错误。");
    }

    /// <summary>
    /// 复制数据校验定义。
    /// </summary>
    /// <returns>与原定义相互独立的快照。</returns>
    internal ExcelDataValidationDefinition Clone() => new()
    {
        Range = Range == null ? null : new ExcelRangeDefinition
        {
            StartRow = Range.StartRow, EndRow = Range.EndRow,
            StartColumn = Range.StartColumn, EndColumn = Range.EndColumn
        },
        Type = Type, Operator = Operator, Value1 = Value1, Value2 = Value2,
        Date1 = Date1, Date2 = Date2, Values = Values?.ToArray(), ListSourceName = ListSourceName,
        CascadeSourceName = CascadeSourceName, ParentColumnIndex = ParentColumnIndex,
        IgnoreBlanks = IgnoreBlanks, ShowErrorMessage = ShowErrorMessage,
        InputTitle = InputTitle, InputMessage = InputMessage,
        ErrorTitle = ErrorTitle, ErrorMessage = ErrorMessage
    };
}
