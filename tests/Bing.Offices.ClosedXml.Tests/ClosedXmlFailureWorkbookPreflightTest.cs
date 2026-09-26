using System.Collections.Generic;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 ClosedXML Failure Workbook 的候选错误行预算语义。
/// </summary>
public sealed class ClosedXmlFailureWorkbookPreflightTest
{
    /// <summary>
    /// 同一 Sheet 的同一物理行多个错误只消耗一个候选行。
    /// </summary>
    [Fact]
    public void ValidatePreflight_ShouldDeduplicateErrorsOnSameSheetRow()
    {
        using var workbook = CreateWorkbook();
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
            MaxCandidateErrorRows = 1
        };

        ClosedXmlFailureWorkbookWriter.ValidatePreflight(workbook, options, new[]
        {
            Error("Data", 2, 1),
            Error("Data", 2, 2)
        }, new Dictionary<string, ExcelSheetImportRequest>());
    }

    /// <summary>
    /// 相同 RowIndex 但不同 Sheet 必须消耗两个候选行。
    /// </summary>
    [Fact]
    public void ValidatePreflight_ShouldCountSameRowIndexOnDifferentSheetsSeparately()
    {
        using var workbook = CreateWorkbook();
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
            MaxCandidateErrorRows = 1
        };

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            ClosedXmlFailureWorkbookWriter.ValidatePreflight(workbook, options, new[]
            {
                Error("Data", 2, 1),
                Error("Other", 2, 1)
            }, new Dictionary<string, ExcelSheetImportRequest>()));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    /// <summary>
    /// AnnotatedOriginal 不应用仅错误行候选预算。
    /// </summary>
    [Fact]
    public void ValidatePreflight_ShouldIgnoreCandidateBudgetForAnnotatedOriginal()
    {
        using var workbook = CreateWorkbook();
        var options = new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            MaxCandidateErrorRows = 1
        };

        ClosedXmlFailureWorkbookWriter.ValidatePreflight(workbook, options, new[]
        {
            Error("Data", 2, 1),
            Error("Other", 2, 1)
        }, new Dictionary<string, ExcelSheetImportRequest>());
    }

    /// <summary>
    /// 创建带两个数据工作表的最小源工作簿。
    /// </summary>
    /// <returns>供当前测试使用的工作簿。</returns>
    private static XLWorkbook CreateWorkbook()
    {
        var workbook = new XLWorkbook();
        workbook.Worksheets.Add("Data").Cell(1, 1).Value = "Code";
        workbook.Worksheets.Add("Other").Cell(1, 1).Value = "Code";
        return workbook;
    }

    /// <summary>
    /// 创建用于预检计数的结构化错误。
    /// </summary>
    /// <param name="sheet">错误所属工作表名称。</param>
    /// <param name="row">错误所在物理行号，从 1 开始。</param>
    /// <param name="column">错误所在物理列号，从 1 开始。</param>
    /// <returns>包含指定物理位置的结构化错误。</returns>
    private static ExcelImportError Error(string sheet, int row, int column) =>
        new(ExcelImportErrorCode.InvalidInput, "invalid", sheet, row, column, "Code");
}
