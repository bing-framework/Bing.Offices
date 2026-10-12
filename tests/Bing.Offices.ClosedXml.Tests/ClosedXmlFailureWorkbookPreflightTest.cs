using System.Collections.Generic;
using System;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using System.Xml.Linq;
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
    /// 验证失败工作簿输入预检应用调用方 XML 预算且不修改目标流。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导入入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Import_FailureWorkbookXmlLimit_ShouldRejectBeforeWriting(bool asynchronous)
    {
        var bytes = CreateWorkbookWithLargeRootRelationship();
        using var source = new MemoryStream(bytes);
        using var destination = new MemoryStream();
        var sentinel = new byte[] { 3, 7, 11 };
        destination.Write(sentinel);
        destination.Position = 0;
        var request = ExcelImport.Workbook<ImportRoot>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxXmlCharacters = 32768 })
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination
            })
            .Sheet<ImportRow>("Data", root => root.Rows));
        var importer = new ClosedXmlExcelImporter();

        BingOfficesResourceLimitException exception;
        if (asynchronous)
        {
            exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
                importer.ImportAsync(source, request));
        }
        else
        {
            exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
                importer.Import(source, request));
        }

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal("ClosedXML", exception.Provider);
        Assert.Contains("_rels/.rels", exception.Message);
        Assert.Equal(sentinel, destination.ToArray());
        Assert.True(source.CanRead);
        Assert.True(destination.CanWrite);
    }

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
    /// 创建根关系 XML 超出预算但其它 XLSX XML 未超限的工作簿。
    /// </summary>
    /// <returns>包含长根关系元数据的 XLSX 文件。</returns>
    private static byte[] CreateWorkbookWithLargeRootRelationship()
    {
        using var source = new MemoryStream();
        using (var workbook = CreateWorkbook())
            workbook.SaveAs(source);
        source.Position = 0;

        using var input = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                var target = archive.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                using var inputStream = entry.Open();
                using var outputStream = target.Open();
                if (!string.Equals(entry.FullName, "_rels/.rels", StringComparison.Ordinal))
                {
                    inputStream.CopyTo(outputStream);
                    continue;
                }

                var relationships = XDocument.Load(inputStream);
                XNamespace namespaceName = "http://schemas.openxmlformats.org/package/2006/relationships";
                relationships.Root.Add(new XElement(namespaceName + "Relationship",
                    new XAttribute("Id", "rIdResourceBudget"),
                    new XAttribute("Type", "urn:bing-offices:resource-budget"),
                    new XAttribute("Target", new string('x', 65536))));
                relationships.Save(outputStream, System.Xml.Linq.SaveOptions.DisableFormatting);
            }
        }

        return output.ToArray();
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

    private sealed class ImportRoot
    {
        public List<ImportRow> Rows { get; } = new();
    }

    private sealed class ImportRow
    {
        public string Code { get; set; }
    }
}
