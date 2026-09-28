namespace Bing.Offices.Imports;

/// <summary>
/// 记录导入实体来源的工作表和行号。
/// </summary>
internal sealed class SourceLocation
{
    /// <summary>
    /// 初始化一个 <see cref="SourceLocation" /> 类型的实例。
    /// </summary>
    /// <param name="sheetName">实体所属的工作表名称。</param>
    /// <param name="rowIndex">实体所在的一基数据行号。</param>
    internal SourceLocation(string sheetName, int rowIndex)
    {
        SheetName = sheetName;
        RowIndex = rowIndex;
    }

    /// <summary>
    /// 获取实体所属的工作表名称。
    /// </summary>
    internal string SheetName { get; }
    /// <summary>
    /// 获取实体所在的一基数据行号。
    /// </summary>
    internal int RowIndex { get; }
}
