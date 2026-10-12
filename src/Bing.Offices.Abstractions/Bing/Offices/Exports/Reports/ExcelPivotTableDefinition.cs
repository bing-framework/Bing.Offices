using System;
using System.Collections.Generic;
using System.Linq;

namespace Bing.Offices.Exports;

/// <summary>
/// 定义工作簿透视表。
/// </summary>
/// <remarks>
/// 默认布局包含一个行字段和一个求和字段，也可描述多维轴、多个度量及报表筛选。
/// </remarks>
public sealed class ExcelPivotTableDefinition
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTableDefinition"/> 类型的实例。
    /// </summary>
    public ExcelPivotTableDefinition()
    {
    }

    /// <summary>
    /// 获取或初始化透视表名称。
    /// </summary>
    public string Name { get; init; }

    /// <summary>
    /// 获取或初始化来源工作表名称。
    /// </summary>
    public string SourceSheet { get; init; }

    /// <summary>
    /// 获取或初始化包含表头的来源区域。
    /// </summary>
    public ExcelRangeDefinition SourceRange { get; init; }

    /// <summary>
    /// 获取或初始化目标工作表名称。
    /// </summary>
    /// <remarks>新建时指定新工作表；更新时指定已有透视所在工作表。</remarks>
    public string TargetSheet { get; init; }

    /// <summary>
    /// 获取或初始化零基目标行。
    /// </summary>
    public int TargetRow { get; init; }

    /// <summary>
    /// 获取或初始化零基目标列。
    /// </summary>
    public int TargetColumn { get; init; }

    /// <summary>
    /// 获取或初始化行字段的来源表头。
    /// </summary>
    public string RowField { get; init; }

    /// <summary>
    /// 获取或初始化求和字段的来源表头。
    /// </summary>
    public string ValueField { get; init; }

    /// <summary>
    /// 获取或初始化按层级顺序排列的行轴字段。
    /// </summary>
    public IReadOnlyList<ExcelPivotAxisFieldDefinition> RowFields { get; init; } =
        Array.Empty<ExcelPivotAxisFieldDefinition>();

    /// <summary>
    /// 获取或初始化按层级顺序排列的列轴字段。
    /// </summary>
    public IReadOnlyList<ExcelPivotAxisFieldDefinition> ColumnFields { get; init; } =
        Array.Empty<ExcelPivotAxisFieldDefinition>();

    /// <summary>
    /// 获取或初始化按显示顺序排列的度量字段。
    /// </summary>
    public IReadOnlyList<ExcelPivotValueDefinition> ValueFields { get; init; } =
        Array.Empty<ExcelPivotValueDefinition>();

    /// <summary>
    /// 获取或初始化透视表的报表筛选。
    /// </summary>
    public IReadOnlyList<ExcelPivotFilterDefinition> Filters { get; init; } =
        Array.Empty<ExcelPivotFilterDefinition>();

    /// <summary>
    /// 获取或初始化透视表保守展开单元格上限。
    /// </summary>
    /// <remarks>
    /// 未设置时仅执行 Excel 格式边界检查；设置后必须大于零。
    /// </remarks>
    public long? MaxExpandedCells { get; init; }

    /// <summary>
    /// 验证透视表定义。
    /// </summary>
    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(Name) || string.IsNullOrWhiteSpace(SourceSheet)
            || string.IsNullOrWhiteSpace(TargetSheet))
            throw new ArgumentException("透视表名称和工作表名称不能为空。");
        ExcelSheetNameValidation.Validate(SourceSheet);
        ExcelSheetNameValidation.Validate(TargetSheet);
        if (TargetRow < 0 || TargetColumn < 0)
            throw new ArgumentOutOfRangeException(nameof(TargetRow), "透视表目标坐标不能为负数。");
        if (SourceRange == null)
            throw new ArgumentException("透视表必须指定来源区域。", nameof(SourceRange));
        SourceRange.Validate(nameof(SourceRange));
        if (SourceRange.EndRow == SourceRange.StartRow)
            throw new ArgumentException("透视表来源区域必须包含表头和至少一行数据。", nameof(SourceRange));
        if (MaxExpandedCells.HasValue && MaxExpandedCells.Value <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxExpandedCells),
                "透视表保守展开单元格上限必须大于零。");

        if (RowFields == null || ColumnFields == null || ValueFields == null || Filters == null)
            throw new ArgumentException("透视表高级字段集合不能为 null。");

        var advanced = RowFields.Count != 0 || ColumnFields.Count != 0
            || ValueFields.Count != 0 || Filters.Count != 0;
        if (!advanced)
        {
            ValidateLegacyFields();
            return;
        }

        if (RowField != null || ValueField != null)
            throw new ArgumentException("旧版行字段和求和字段不能与高级字段集合混用。");
        ValidateAdvancedFields();
    }

    /// <summary>
    /// 生成独立的透视表定义。
    /// </summary>
    /// <returns>包含独立来源区域的定义快照。</returns>
    internal ExcelPivotTableDefinition Clone() => new()
    {
        Name = Name,
        SourceSheet = SourceSheet,
        SourceRange = new ExcelRangeDefinition
        {
            StartRow = SourceRange.StartRow,
            StartColumn = SourceRange.StartColumn,
            EndRow = SourceRange.EndRow,
            EndColumn = SourceRange.EndColumn
        },
        TargetSheet = TargetSheet,
        TargetRow = TargetRow,
        TargetColumn = TargetColumn,
        RowField = RowField,
        ValueField = ValueField,
        RowFields = CloneAxes(RowFields),
        ColumnFields = CloneAxes(ColumnFields),
        ValueFields = Array.AsReadOnly(ValueFields.Select(value => new ExcelPivotValueDefinition
        {
            Field = value.Field,
            Caption = value.Caption,
            Aggregation = value.Aggregation,
            NumberFormat = value.NumberFormat
        }).ToArray()),
        Filters = Array.AsReadOnly(Filters.Select(filter => new ExcelPivotFilterDefinition
        {
            Field = filter.Field,
            Values = CloneFilterValues(filter.Values)
        }).ToArray()),
        MaxExpandedCells = MaxExpandedCells
    };

    /// <summary>
    /// 验证兼容旧版的单行 Sum 字段。
    /// </summary>
    private void ValidateLegacyFields()
    {
        if (string.IsNullOrWhiteSpace(RowField) || string.IsNullOrWhiteSpace(ValueField))
            throw new ArgumentException("旧版透视表行字段和求和字段不能为空。");
        if (string.Equals(RowField, ValueField, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("透视表行字段与求和字段不能相同。");
    }

    /// <summary>
    /// 验证高级透视字段、角色冲突和筛选值。
    /// </summary>
    private void ValidateAdvancedFields()
    {
        if (RowFields.Count == 0 && ColumnFields.Count == 0)
            throw new ArgumentException("高级透视表至少需要一个行轴或列轴字段。");
        if (ValueFields.Count == 0)
            throw new ArgumentException("高级透视表至少需要一个度量字段。");

        var fieldRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        ValidateAxes(RowFields, fieldRoles);
        ValidateAxes(ColumnFields, fieldRoles);

        foreach (var filter in Filters)
        {
            if (filter == null)
                throw new ArgumentException("透视表筛选定义不能为 null。", nameof(Filters));
            ValidateField(filter.Field, nameof(Filters));
            if (!fieldRoles.Add(filter.Field))
                throw new ArgumentException($"透视表字段不能重复声明不同角色: {filter.Field}");
            if (filter.Values == null)
                throw new ArgumentException("透视表筛选值集合不能为 null。", nameof(Filters));
            foreach (var value in filter.Values)
                if (!IsAllowedFilterValue(value))
                    throw new ArgumentException("透视表筛选值必须是 Excel 字面值。", nameof(Filters));
        }

        foreach (var value in ValueFields)
        {
            if (value == null)
                throw new ArgumentException("透视表度量定义不能为 null。", nameof(ValueFields));
            ValidateField(value.Field, nameof(ValueFields));
            if (!Enum.IsDefined(typeof(ExcelPivotAggregation), value.Aggregation))
                throw new ArgumentOutOfRangeException(nameof(ValueFields), "透视表聚合类型无效。");
            if (fieldRoles.Contains(value.Field))
                throw new ArgumentException($"透视表字段不能重复声明不同角色: {value.Field}");
        }

        var captions = new HashSet<string>(fieldRoles, StringComparer.OrdinalIgnoreCase);
        foreach (var value in ValueFields)
        {
            if (string.IsNullOrWhiteSpace(value.Caption))
                throw new ArgumentException("透视表度量标题不能为空。", nameof(ValueFields));
            if (!captions.Add(value.Caption))
                throw new ArgumentException($"透视表标题与字段或其它度量标题冲突: {value.Caption}");
        }
    }

    /// <summary>
    /// 验证行轴或列轴中的字段并登记字段角色。
    /// </summary>
    /// <param name="fields">待验证的轴字段。</param>
    /// <param name="fieldRoles">已登记字段名称的集合。</param>
    private static void ValidateAxes(IReadOnlyList<ExcelPivotAxisFieldDefinition> fields,
        ISet<string> fieldRoles)
    {
        foreach (var field in fields)
        {
            if (field == null)
                throw new ArgumentException("透视表轴字段定义不能为 null。");
            ValidateField(field.Field, nameof(fields));
            if (!Enum.IsDefined(typeof(ExcelPivotSortOrder), field.SortOrder))
                throw new ArgumentOutOfRangeException(nameof(fields), "透视表排序方式无效。");
            if (!Enum.IsDefined(typeof(ExcelPivotSubtotalMode), field.SubtotalMode))
                throw new ArgumentOutOfRangeException(nameof(fields), "透视表小计方式无效。");
            if (!fieldRoles.Add(field.Field))
                throw new ArgumentException($"透视表轴字段不能重复: {field.Field}");
        }
    }

    /// <summary>
    /// 验证字段名称不为空白。
    /// </summary>
    /// <param name="field">来源表头名称。</param>
    /// <param name="parameterName">触发验证的参数名称。</param>
    private static void ValidateField(string field, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(field))
            throw new ArgumentException("透视表字段名称不能为空。", parameterName);
    }

    /// <summary>
    /// 判断筛选值是否属于允许的 Excel 字面值类型。
    /// </summary>
    /// <param name="value">待验证的筛选值；null 表示空项。</param>
    /// <returns>true 表示该值可作为筛选字面值；false 表示类型不受支持。</returns>
    private static bool IsAllowedFilterValue(object value)
    {
        if (value == null || value is string || value is bool || value is DateTime)
            return true;
        if (value.GetType().IsEnum)
            return false;
        if (value is double doubleValue)
            return !double.IsNaN(doubleValue) && !double.IsInfinity(doubleValue);
        if (value is float floatValue)
            return !float.IsNaN(floatValue) && !float.IsInfinity(floatValue);

        switch (Type.GetTypeCode(value.GetType()))
        {
            case TypeCode.Byte:
            case TypeCode.SByte:
            case TypeCode.Int16:
            case TypeCode.UInt16:
            case TypeCode.Int32:
            case TypeCode.UInt32:
            case TypeCode.Int64:
            case TypeCode.UInt64:
            case TypeCode.Decimal:
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// 按 Excel 筛选比较语义保留首次出现的唯一值。
    /// </summary>
    /// <param name="values">待复制的筛选值。</param>
    /// <returns>只读且保持首次出现顺序的筛选值集合。</returns>
    private static IReadOnlyList<object> CloneFilterValues(IReadOnlyList<object> values)
    {
        var distinct = new List<object>();
        foreach (var value in values)
            if (!distinct.Any(existing => FilterValuesEqual(existing, value)))
                distinct.Add(value);
        return Array.AsReadOnly(distinct.ToArray());
    }

    /// <summary>
    /// 比较两个筛选值在 Excel 中的匹配语义。
    /// </summary>
    /// <param name="left">第一个值。</param>
    /// <param name="right">第二个值。</param>
    /// <returns>true 表示两个值在透视筛选中相同。</returns>
    private static bool FilterValuesEqual(object left, object right)
    {
        if (left == null || right == null)
            return left == right;
        if (left is string leftText && right is string rightText)
            return string.Equals(leftText, rightText, StringComparison.OrdinalIgnoreCase);
        if (IsNumeric(left) && IsNumeric(right))
            return Convert.ToDouble(left) == Convert.ToDouble(right);
        return left.GetType() == right.GetType() && left.Equals(right);
    }

    /// <summary>
    /// 判断值是否为受支持的数值字面类型。
    /// </summary>
    /// <param name="value">待判断的值。</param>
    /// <returns>true 表示值为整数、十进制数或浮点数。</returns>
    private static bool IsNumeric(object value)
    {
        if (value is double || value is float)
            return true;
        switch (Type.GetTypeCode(value.GetType()))
        {
            case TypeCode.Byte:
            case TypeCode.SByte:
            case TypeCode.Int16:
            case TypeCode.UInt16:
            case TypeCode.Int32:
            case TypeCode.UInt32:
            case TypeCode.Int64:
            case TypeCode.UInt64:
            case TypeCode.Decimal:
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// 深克隆行轴或列轴字段集合。
    /// </summary>
    /// <param name="fields">待克隆的轴字段集合。</param>
    /// <returns>只读的独立轴字段集合。</returns>
    private static IReadOnlyList<ExcelPivotAxisFieldDefinition> CloneAxes(
        IReadOnlyList<ExcelPivotAxisFieldDefinition> fields) =>
        Array.AsReadOnly(fields.Select(field => new ExcelPivotAxisFieldDefinition
        {
            Field = field.Field,
            SortOrder = field.SortOrder,
            SubtotalMode = field.SubtotalMode
        }).ToArray());
}
