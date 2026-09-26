namespace Bing.Offices.Exports;

/// <summary>
/// 工作表原生数据校验类型。
/// </summary>
/// <remarks>用于生成工作表规则，不执行导入校验。</remarks>
public enum ExcelDataValidationType
{
    /// <summary>
    /// 整数。
    /// </summary>
    Integer,
    /// <summary>
    /// 小数。
    /// </summary>
    Decimal,
    /// <summary>
    /// 日期。
    /// </summary>
    Date,
    /// <summary>
    /// 文本长度。
    /// </summary>
    TextLength,
    /// <summary>
    /// 显式列表。
    /// </summary>
    List
}

/// <summary>
/// 有界区域的 Excel 原生数据校验定义。
/// </summary>
/// <remarks>不接受公式或外部引用。</remarks>
public sealed class ExcelDataValidationDefinition
{
    /// <summary>
    /// 获取或初始化应用区域，零基且包含结束位置。
    /// </summary>
    public ExcelRangeDefinition Range { get; init; }
    /// <summary>
    /// 获取或初始化校验类型。
    /// </summary>
    public ExcelDataValidationType Type { get; init; }
    /// <summary>
    /// 获取或初始化比较运算符，列表不使用。
    /// </summary>
    public ExcelConditionalComparisonOperator Operator { get; init; }
    /// <summary>
    /// 获取或初始化数字或长度的第一个边界。
    /// </summary>
    public double Value1 { get; init; }
    /// <summary>
    /// 获取或初始化Between/NotBetween 的第二个边界。
    /// </summary>
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
    /// 获取或初始化显式列表项，不支持逗号、引号及换行。
    /// </summary>
    public IReadOnlyList<string> Values { get; init; } = Array.Empty<string>();
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
            if (Values == null || Values.Count == 0 || Values.Any(value => string.IsNullOrEmpty(value)
                || value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0) || string.Join(",", Values).Length > 253)
                throw new ArgumentException("显式列表不能为空，不能包含分隔符，编码后含引号最多 255 字符。", nameof(Values));
            return;
        }
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
}

/// <summary>
/// 工作表嵌入图片定义。
/// </summary>
/// <remarks>支持 PNG/JPEG，采用零基单元格锚点，尺寸和偏移单位为像素。</remarks>
public sealed class ExcelSheetImageDefinition
{
    /// <summary>
    /// 获取或初始化图片编码内容。
    /// </summary>
    /// <remarks>构建请求时复制内容，不读取外部文件或网络资源。</remarks>
    public byte[] Content { get; init; }
    /// <summary>
    /// 获取或初始化零基行。
    /// </summary>
    public int Row { get; init; }
    /// <summary>
    /// 获取或初始化零基列。
    /// </summary>
    public int Column { get; init; }
    /// <summary>
    /// 获取或初始化输出宽度。
    /// </summary>
    public int Width { get; init; }
    /// <summary>
    /// 获取或初始化输出高度。
    /// </summary>
    public int Height { get; init; }
    /// <summary>
    /// 获取或初始化水平偏移。
    /// </summary>
    public int OffsetX { get; init; }
    /// <summary>
    /// 获取或初始化垂直偏移。
    /// </summary>
    public int OffsetY { get; init; }

    /// <summary>
    /// 验证图片内容签名和坐标。
    /// </summary>
    /// <remarks>图片编码完整性由图片读取器验证。</remarks>
    public void Validate()
    {
        if (Row < 0 || Column < 0 || Width <= 0 || Height <= 0 || OffsetX < 0 || OffsetY < 0)
            throw new ArgumentOutOfRangeException(nameof(Row), "图片坐标和偏移非负，尺寸为正数。");
        var png = Content != null && Content.Length >= 24 && Content[0] == 137 && Content[1] == 80
            && Content[2] == 78 && Content[3] == 71 && Content[4] == 13 && Content[5] == 10 && Content[6] == 26 && Content[7] == 10;
        var jpeg = Content != null && Content.Length >= 4 && Content[0] == 255 && Content[1] == 216 && Content[2] == 255;
        if (!png && !jpeg) throw new ArgumentException("图片必须是 PNG 或 JPEG 编码。", nameof(Content));
    }
}
