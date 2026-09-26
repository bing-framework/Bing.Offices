using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.ExcelDataReader.Tests;

/// <summary>
/// ExcelDataReader 分批、资源边界和异常契约测试。
/// </summary>
public sealed class ExcelDataReaderBoundaryTest
{
    /// <summary>
    /// 验证缺少表头时分批导入仅交付结构化错误。
    /// </summary>
    [Fact]
    public void ImportBatches_MissingHeaders_ShouldDeliverOnlyStructuredErrors()
    {
        using var source = CreateSheet("People", new[] { "Unknown" }, includeSpacer: false);
        var batches = new List<ExcelImportBatch<Person>>();

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source,
            new ExcelBatchImportRequest<Person>("People"), batches.Add);

        var batch = Assert.Single(batches);
        Assert.Empty(batch.Items);
        Assert.Contains(batch.Errors, error => error.Code == ExcelImportErrorCode.InvalidHeader);
        Assert.False(summary.IsSuccess);
        Assert.True(summary.ErrorCount >= 2);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证缺少表头时异步分批导入仅交付结构化错误。
    /// </summary>
    [Fact]
    public async Task ImportBatchesAsync_MissingHeaders_ShouldDeliverOnlyStructuredErrors()
    {
        using var source = CreateSheet("People", new[] { "Unknown" }, includeSpacer: false);
        var batches = new List<ExcelImportBatch<Person>>();

        var summary = await new ExcelDataReaderExcelImporter().ImportBatchesAsync(source,
            new ExcelBatchImportRequest<Person>("People"),
            (batch, _) =>
            {
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.Empty(Assert.Single(batches).Items);
        Assert.False(summary.IsSuccess);
        Assert.Contains(batches[0].Errors, error => error.Code == ExcelImportErrorCode.InvalidHeader);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证完整导入和分批导入遵循空行后的零基数据起始位置。
    /// </summary>
    [Fact]
    public void ImportAndBatch_ShouldHonorZeroBasedDataStartAfterSpacer()
    {
        using var source = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: true);
        var importer = new ExcelDataReaderExcelImporter();
        var workbookRequest = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .Sheet<Person>("People", root => root.People,
                sheet => sheet.DataRowStartIndex(2)));

        var result = importer.Import(source, workbookRequest);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Alice", Assert.Single(result.Workbook.People).Name);
        source.Position = 0;
        var batches = new List<ExcelImportBatch<Person>>();
        var summary = importer.ImportBatches(source, new ExcelBatchImportRequest<Person>("People")
        {
            DataRowStartIndex = 2,
            BatchSize = 1
        }, batches.Add);

        Assert.True(summary.IsSuccess);
        Assert.Equal(3, Assert.Single(batches).FirstRowIndex);
        Assert.Equal("Alice", Assert.Single(batches[0].Items).Name);
    }

    /// <summary>
    /// 验证异步分批导入遵循空行后的零基数据起始位置。
    /// </summary>
    [Fact]
    public async Task ImportBatchesAsync_ShouldHonorZeroBasedDataStartAfterSpacer()
    {
        using var source = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: true);
        var batches = new List<ExcelImportBatch<Person>>();

        var summary = await new ExcelDataReaderExcelImporter().ImportBatchesAsync(source,
            new ExcelBatchImportRequest<Person>("People")
            {
                DataRowStartIndex = 2,
                BatchSize = 1
            },
            (batch, _) =>
            {
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess);
        Assert.Equal(3, Assert.Single(batches).FirstRowIndex);
        Assert.Equal("Alice", Assert.Single(batches[0].Items).Name);
    }

    /// <summary>
    /// 验证分批导入在回调前拒绝动态列和唯一值映射。
    /// </summary>
    [Fact]
    public void ImportBatches_DynamicAndUniqueMappings_ShouldBeRejectedBeforeCallback()
    {
        using var dynamicSource = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        var dynamicBatches = new List<ExcelImportBatch<Person>>();
        var dynamicException = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ExcelDataReaderExcelImporter().ImportBatches(dynamicSource,
                new ExcelBatchImportRequest<Person>("People")
                {
                    MappingConfiguration = new ExcelMappingConfiguration
                    {
                        DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
                        {
                            new ExcelMappingDynamicColumnConfiguration { Key = "region", Title = "Region" }
                        }
                    }
                }, dynamicBatches.Add));

        Assert.Equal("ExcelDataReader", dynamicException.Provider);
        Assert.Equal(BingOfficesOperation.Import, dynamicException.Operation);
        Assert.Equal(BingOfficesStage.Preflight, dynamicException.Stage);
        Assert.Empty(dynamicBatches);
        Assert.True(dynamicSource.CanRead);

        using var uniqueSource = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        var uniqueBatches = new List<ExcelImportBatch<UniquePerson>>();
        var uniqueException = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ExcelDataReaderExcelImporter().ImportBatches(uniqueSource,
                new ExcelBatchImportRequest<UniquePerson>("People"), uniqueBatches.Add));

        Assert.Equal("ExcelDataReader", uniqueException.Provider);
        Assert.Equal(BingOfficesStage.Preflight, uniqueException.Stage);
        Assert.Empty(uniqueBatches);
        Assert.True(uniqueSource.CanRead);
    }

    /// <summary>
    /// 验证工作表、列和单元格超限时不返回部分实体。
    /// </summary>
    [Fact]
    public void Import_ShouldApplySheetColumnAndCellBudgetsWithoutPartialEntities()
    {
        using var source = CreateWorkbookWithTwoSheets();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxSheets = 1 })
            .Sheet<Person>("People", root => root.People));

        var sheetResult = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(sheetResult.IsSuccess);
        Assert.Empty(sheetResult.Workbook.People);
        Assert.Contains(sheetResult.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);

        using var columnSource = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        var columnResult = new ExcelDataReaderExcelImporter().Import(columnSource,
            ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .ResourceLimits(new ExcelResourceLimits { MaxColumnsPerSheet = 1 })
                .Sheet<Person>("People", root => root.People)));
        Assert.False(columnResult.IsSuccess);
        Assert.Empty(columnResult.Workbook.People);
        Assert.Contains(columnResult.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);

        using var cellSource = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        var cellResult = new ExcelDataReaderExcelImporter().Import(cellSource,
            ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .ResourceLimits(new ExcelResourceLimits { MaxCells = 2 })
                .Sheet<Person>("People", root => root.People)));
        Assert.False(cellResult.IsSuccess);
        Assert.Empty(cellResult.Workbook.People);
        Assert.Contains(cellResult.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);
    }

    /// <summary>
    /// 验证单元格预算覆盖未选工作表及表头之前的单元格。
    /// </summary>
    [Fact]
    public void Import_ShouldCountWorkbookCellsOnUnselectedSheetsAndBeforeHeader()
    {
        using var workbookSource = CreateWorkbookWithTwoSheets();
        var workbookResult = new ExcelDataReaderExcelImporter().Import(workbookSource,
            ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .ResourceLimits(new ExcelResourceLimits { MaxCells = 4 })
                .Sheet<Person>("People", root => root.People)));

        Assert.False(workbookResult.IsSuccess);
        Assert.Empty(workbookResult.Workbook.People);
        Assert.Contains(workbookResult.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);

        using var preambleSource = CreateSheetWithPreamble();
        var preambleResult = new ExcelDataReaderExcelImporter().Import(preambleSource,
            ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .ResourceLimits(new ExcelResourceLimits { MaxCells = 4 })
                .Sheet<Person>("People", root => root.People,
                    sheet => sheet.HeaderRowIndex(1).DataRowStartIndex(2))));

        Assert.False(preambleResult.IsSuccess);
        Assert.Empty(preambleResult.Workbook.People);
        Assert.Contains(preambleResult.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);

        using var boundarySource = CreateSheetWithPreamble();
        var boundaryResult = new ExcelDataReaderExcelImporter().Import(boundarySource,
            ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .ResourceLimits(new ExcelResourceLimits { MaxCells = 5 })
                .Sheet<Person>("People", root => root.People,
                    sheet => sheet.HeaderRowIndex(1).DataRowStartIndex(2))));

        Assert.True(boundaryResult.IsSuccess, string.Join(";", boundaryResult.Errors.Select(error => error.Message)));
        Assert.Single(boundaryResult.Workbook.People);
    }

    /// <summary>
    /// 验证异步导入的单元格预算覆盖未选工作表。
    /// </summary>
    [Fact]
    public async Task ImportAsync_ShouldCountWorkbookCellsOnUnselectedSheets()
    {
        using var source = CreateWorkbookWithTwoSheets();
        var result = await new ExcelDataReaderExcelImporter().ImportAsync(source,
            ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .ResourceLimits(new ExcelResourceLimits { MaxCells = 4 })
                .Sheet<Person>("People", root => root.People)));

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.People);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);
    }

    /// <summary>
    /// 验证唯一值跟踪超限时返回资源错误且不返回部分实体。
    /// </summary>
    [Fact]
    public void Import_UniqueLimit_ShouldReturnResourceErrorWithoutPartialEntities()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People", new[]
        {
            new UniquePerson { Name = "A", Age = 1 },
            new UniquePerson { Name = "B", Age = 2 }
        }));
        using var source = new MemoryStream();
        new Bing.Offices.Exports.NpoiExcelExporter().Export(export, source);
        source.Position = 0;
        var request = ExcelImport.Workbook<UniqueWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxTrackedUniqueValues = 1 })
            .Sheet<UniquePerson>("People", root => root.Rows));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);
    }

    /// <summary>
    /// 验证损坏工作簿在分批导入时产生结构化异常。
    /// </summary>
    [Fact]
    public void ImportBatches_MalformedWorkbook_ShouldUseStructuredProviderException()
    {
        using var source = new MemoryStream(new byte[] { 1, 2, 3, 4, 5 });
        var exception = Assert.Throws<BingOfficesImportException>(() =>
            new ExcelDataReaderExcelImporter().ImportBatches(source,
                new ExcelBatchImportRequest<Person>("People"), _ => { }));

        Assert.Equal("ExcelDataReader", exception.Provider);
        Assert.Equal(BingOfficesErrorCode.ImportFailed, exception.Code);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(BingOfficesStage.Read, exception.Stage);
        Assert.NotNull(exception.InnerException);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证分批回调异常原样传播。
    /// </summary>
    [Fact]
    public void ImportBatches_CallbackException_ShouldPropagateOriginalException()
    {
        using var source = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        var expected = new InvalidOperationException("callback failure");

        var actual = Assert.Throws<InvalidOperationException>(() =>
            new ExcelDataReaderExcelImporter().ImportBatches(source,
                new ExcelBatchImportRequest<Person>("People") { BatchSize = 1 },
                _ => throw expected));

        Assert.Same(expected, actual);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证异步批次取消传播并清理临时文件。
    /// </summary>
    [Fact]
    public async Task ImportBatchesAsync_CallbackCancellation_ShouldPropagateAndCleanStaging()
    {
        using var source = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        using var cancellation = new CancellationTokenSource();
        var stagingDirectory = Path.Combine(Path.GetTempPath(),
            "bing-offices-excel-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(stagingDirectory);

        try
        {
            var importer = new ExcelDataReaderExcelImporter(() => Path.Combine(
                stagingDirectory, Guid.NewGuid().ToString("N") + ".tmp"));

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                importer.ImportBatchesAsync(source,
                    new ExcelBatchImportRequest<Person>("People") { BatchSize = 1 },
                    (batch, token) =>
                    {
                        cancellation.Cancel();
                        token.ThrowIfCancellationRequested();
                        return Task.CompletedTask;
                    }, cancellation.Token));

            Assert.True(source.CanRead);
            Assert.Empty(Directory.GetFiles(stagingDirectory));
        }
        finally
        {
            if (Directory.Exists(stagingDirectory))
                Directory.Delete(stagingDirectory, recursive: true);
        }
    }

    /// <summary>
    /// 创建包含指定表头和一行数据的测试工作簿。
    /// </summary>
    /// <param name="name">测试工作表名称。</param>
    /// <param name="headers">按列排列的表头文本。</param>
    /// <param name="includeSpacer">是否在表头与数据行之间保留空行。</param>
    /// <returns>包含测试工作簿内容且位于起始位置的流。</returns>
    private static MemoryStream CreateSheet(string name, string[] headers, bool includeSpacer)
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet(name);
        var header = sheet.CreateRow(0);
        for (var index = 0; index < headers.Length; index++)
            header.CreateCell(index).SetCellValue(headers[index]);
        var dataRow = sheet.CreateRow(includeSpacer ? 2 : 1);
        for (var index = 0; index < headers.Length; index++)
        {
            if (string.Equals(headers[index], "Age", StringComparison.OrdinalIgnoreCase))
                dataRow.CreateCell(index).SetCellValue(30);
            else if (string.Equals(headers[index], "Name", StringComparison.OrdinalIgnoreCase))
                dataRow.CreateCell(index).SetCellValue("Alice");
            else
                dataRow.CreateCell(index).SetCellValue("value");
        }
        var stream = new MemoryStream();
        workbook.Write(stream, false);
        var bytes = stream.ToArray();
        workbook.Dispose();
        stream.Dispose();
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>
    /// 创建包含两个工作表的测试工作簿。
    /// </summary>
    /// <returns>包含测试工作簿内容且位于起始位置的流。</returns>
    private static MemoryStream CreateWorkbookWithTwoSheets()
    {
        var first = CreateSheet("People", new[] { "Name", "Age" }, includeSpacer: false);
        using var firstWorkbook = WorkbookFactory.Create(first);
        var second = firstWorkbook.CreateSheet("Other");
        var header = second.CreateRow(0);
        header.CreateCell(0).SetCellValue("Name");
        header.CreateCell(1).SetCellValue("Age");
        var row = second.CreateRow(1);
        row.CreateCell(0).SetCellValue("Bob");
        row.CreateCell(1).SetCellValue(31);
        var output = new MemoryStream();
        firstWorkbook.Write(output, false);
        var bytes = output.ToArray();
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>
    /// 创建表头前含说明行的测试工作簿。
    /// </summary>
    /// <returns>包含测试工作簿内容且位于起始位置的流。</returns>
    private static MemoryStream CreateSheetWithPreamble()
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("People");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Report");
        var header = sheet.CreateRow(1);
        header.CreateCell(0).SetCellValue("Name");
        header.CreateCell(1).SetCellValue("Age");
        var row = sheet.CreateRow(2);
        row.CreateCell(0).SetCellValue("Alice");
        row.CreateCell(1).SetCellValue(30);
        var output = new MemoryStream();
        workbook.Write(output, false);
        var bytes = output.ToArray();
        workbook.Dispose();
        output.Dispose();
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>
    /// 承载人员导入结果的测试工作簿。
    /// </summary>
    private sealed class ImportWorkbook
    {
        /// <summary>
        /// 获取主要工作表导入的人员集合。
        /// </summary>
        public List<Person> People { get; } = new List<Person>();
    }

    /// <summary>
    /// 用于必填姓名映射的人员测试模型。
    /// </summary>
    private sealed class Person
    {
        /// <summary>
        /// 获取或设置测试人员姓名。
        /// </summary>
        [ExcelRequired]
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置测试人员年龄。
        /// </summary>
        public int Age { get; set; }
    }

    /// <summary>
    /// 用于姓名唯一性校验的人员测试模型。
    /// </summary>
    private sealed class UniquePerson
    {
        /// <summary>
        /// 获取或设置测试人员姓名。
        /// </summary>
        [ExcelUnique]
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置测试人员年龄。
        /// </summary>
        public int Age { get; set; }
    }

    /// <summary>
    /// 承载姓名唯一性校验结果的测试工作簿。
    /// </summary>
    private sealed class UniqueWorkbook
    {
        /// <summary>
        /// 获取工作表导入的数据行集合。
        /// </summary>
        public List<UniquePerson> Rows { get; } = new List<UniquePerson>();
    }
}
