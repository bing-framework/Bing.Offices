using System;
using System.Collections.Generic;

namespace Bing.Offices.Imports;

/// <summary>
/// Workbook 导入结果。
/// </summary>
public sealed class ExcelWorkbookImportResult<TWorkbook> where TWorkbook : class, new()
{
    internal ExcelWorkbookImportResult(TWorkbook workbook, IReadOnlyList<ExcelSheetImportResult> sheets,
        IReadOnlyList<ExcelImportError> errors, bool errorsTruncated, int? maxErrors)
    {
        Workbook = workbook;
        Sheets = sheets;
        Errors = errors;
        ErrorsTruncated = errorsTruncated;
        MaxErrors = maxErrors;
    }

    /// <summary>
    /// 获取根 Workbook 模型。
    /// </summary>
    public TWorkbook Workbook { get; }

    /// <summary>
    /// 获取各 Sheet 结果。
    /// </summary>
    public IReadOnlyList<ExcelSheetImportResult> Sheets { get; }

    /// <summary>
    /// 获取关系和结构化导入错误。
    /// </summary>
    public IReadOnlyList<ExcelImportError> Errors { get; }

    /// <summary>
    /// 获取是否因错误数量上限而停止收集后续错误。
    /// </summary>
    public bool ErrorsTruncated { get; }

    /// <summary>
    /// 获取生效的最大错误数。
    /// </summary>
    public int? MaxErrors { get; }

    /// <summary>
    /// 获取是否没有错误。
    /// </summary>
    public bool IsSuccess => Errors.Count == 0;
}

/// <summary>
/// 单个 Sheet 导入结果。
/// </summary>
public sealed class ExcelSheetImportResult
{
    internal ExcelSheetImportResult(string name, Type itemType,
        IReadOnlyList<int> sourceRows, IReadOnlyList<ExcelImportError> errors)
    {
        Name = name;
        ItemType = itemType;
        SourceRows = sourceRows;
        Errors = errors;
    }

    /// <summary>
    /// 获取 Sheet 名称。
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
