using System.ComponentModel;

namespace Bing.Offices.Imports;

/// <summary>
/// 单个 Sheet 导入结果。
/// </summary>
public sealed class ExcelSheetImportResult
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelSheetImportResult" /> 类型的实例。
    /// </summary>
    /// <param name="name">工作表名称。</param>
    /// <param name="itemType">工作表实体类型。</param>
    /// <param name="sourceRows">每个导入实体对应的源行索引。</param>
    /// <param name="errors">当前工作表的导入错误。</param>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelSheetImportResult(string name, Type itemType,
        IReadOnlyList<int> sourceRows, IReadOnlyList<ExcelImportError> errors)
    {
        Name = name;
        ItemType = itemType;
        SourceRows = sourceRows;
        Errors = errors;
    }

    /// <summary>
    /// 获取Sheet 名称。
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// 获取实体类型。
    /// </summary>
    public Type ItemType { get; }

    /// <summary>
    /// 获取实体对应的 zero-based 原始行索引。
    /// </summary>
    public IReadOnlyList<int> SourceRows { get; }

    /// <summary>
    /// 获取该 Sheet 错误。
    /// </summary>
    public IReadOnlyList<ExcelImportError> Errors { get; }
}
