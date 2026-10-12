using System.ComponentModel;

namespace Bing.Offices.Exports;

/// <summary>
/// 定义工作表保护及可填写区域。
/// </summary>
public sealed class ExcelSheetProtectionDefinition
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelSheetProtectionDefinition" /> 类型的实例。
    /// </summary>
    /// <param name="editableRanges">可填写的零基闭区间区域。</param>
    /// <param name="password">可选的工作表保护密码。</param>
    public ExcelSheetProtectionDefinition(IReadOnlyList<ExcelRangeDefinition> editableRanges,
        string password = null)
    {
        if (editableRanges == null)
            throw new ArgumentNullException(nameof(editableRanges));
        var ranges = new ExcelRangeDefinition[editableRanges.Count];
        for (var index = 0; index < ranges.Length; index++)
        {
            var range = editableRanges[index] ?? throw new ArgumentException("可填写区域不能为 null。",
                nameof(editableRanges));
            range.Validate(nameof(editableRanges));
            ranges[index] = new ExcelRangeDefinition
            {
                StartRow = range.StartRow, EndRow = range.EndRow,
                StartColumn = range.StartColumn, EndColumn = range.EndColumn
            };
        }
        EditableRanges = Array.AsReadOnly(ranges);
        Password = password;
    }

    /// <summary>
    /// 获取可填写区域快照。
    /// </summary>
    public IReadOnlyList<ExcelRangeDefinition> EditableRanges { get; }

    /// <summary>
    /// 获取工作表保护密码。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public string Password { get; }

    /// <summary>
    /// 复制保护定义。
    /// </summary>
    internal ExcelSheetProtectionDefinition Clone() => new(EditableRanges, Password);
}
