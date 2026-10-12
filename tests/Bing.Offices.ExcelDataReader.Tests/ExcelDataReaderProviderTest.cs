using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.ExcelDataReader;
using Bing.Offices.ExcelDataReader.Extensions;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Providers;
using Microsoft.Extensions.DependencyInjection;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.ExcelDataReader.Tests;

/// <summary>
/// ExcelDataReader Provider 的职责级测试。
/// </summary>
public sealed class ExcelDataReaderProviderTest
{
    /// <summary>
    /// 验证低保留失败捕获不会重复保存全局和逐行错误引用。
    /// </summary>
    [Fact]
    public void FailureCapture_LowRetention_ShouldAvoidDuplicateErrorReferences()
    {
        var capture = new ExcelDataReaderFailureWorkbookCapture(
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly
            },
            captureErrors: false);
        var error = new ExcelImportError(ExcelImportErrorCode.ValueConversion,
            "转换失败", "明细", 2, 2, "Quantity", rawValue: "wrong");

        capture.CaptureSheet("明细", new object[] { "Code", "Quantity" }, 0);
        capture.CaptureErrorRow("明细", 2, new object[] { "B", "wrong" }, new[] { error });

        Assert.Empty(capture.Errors);
        Assert.Single(capture.Sheets);
        Assert.Empty(capture.Sheets[0].Rows[0].Errors);
    }

    /// <summary>
    /// 验证低保留失败捕获在追加失败行前执行目标对象预算。
    /// </summary>
    [Fact]
    public void FailureCapture_LowRetention_ShouldHonorObjectBudgetBeforeAppendingRow()
    {
        var capture = new ExcelDataReaderFailureWorkbookCapture(
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                MaxEstimatedTargetObjects = 3
            },
            providerName: "CaptureTest", captureErrors: false);
        var error = new ExcelImportError(ExcelImportErrorCode.ValueConversion,
            "转换失败", "明细", 2, 2, "Quantity", rawValue: "wrong");

        capture.CaptureSheet("明细", new object[] { "Code" }, 0);
        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            capture.CaptureErrorRow("明细", 2, new object[] { "B", "wrong" }, new[] { error }));

        Assert.Equal("CaptureTest", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Empty(capture.Sheets.Single().Rows);
    }

    /// <summary>
    /// 验证低保留失败捕获将命名列表上下文行计入目标对象预算。
    /// </summary>
    [Fact]
    public void FailureCapture_LowRetention_ShouldCountNamedListContextObjects()
    {
        var capture = new ExcelDataReaderFailureWorkbookCapture(
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                MaxEstimatedTargetObjects = 3
            },
            providerName: "CaptureTest", captureErrors: false);

        capture.CaptureSheet("字典", new object[] { "Value" }, 0);
        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            capture.CaptureContextRow("字典", 1, new object[] { "采购" }));

        Assert.Equal("CaptureTest", exception.Provider);
        Assert.Empty(capture.Sheets.Single().ContextRows);
    }

    /// <summary>
    /// 验证低保留失败捕获将全局错误暂存计入目标对象预算。
    /// </summary>
    [Fact]
    public void FailureCapture_LowRetention_ShouldCountStagedErrors()
    {
        var capture = new ExcelDataReaderFailureWorkbookCapture(
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                MaxEstimatedTargetObjects = 1
            },
            providerName: "CaptureTest");
        var error = new ExcelImportError(ExcelImportErrorCode.ValueConversion,
            "转换失败", "明细", 2, 2, "Quantity", rawValue: "wrong");
        var secondError = new ExcelImportError(ExcelImportErrorCode.Validation,
            "不能为空", "明细", 3, 1, "Code", rawValue: null);

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            capture.CaptureErrors(new[] { error, secondError }));

        Assert.Equal("CaptureTest", exception.Provider);
        Assert.Single(capture.Errors);
    }

    /// <summary>
    /// 验证 Provider 声明 XLSB 只读和分批导入能力。
    /// </summary>
    [Fact]
    public void Capabilities_ShouldDescribeReadOnlyXlsbAndBatchSupport()
    {
        var importer = new ExcelDataReaderExcelImporter();

        Assert.Contains(ExcelFormat.Xlsb, importer.ReadFormats);
        Assert.Empty(importer.WriteFormats);
        Assert.True(importer.SupportsBatchImport);
        Assert.True(importer.Supports(ExcelProviderFeatures.MultiSheetBatchImport));
        Assert.True(importer.Supports(ExcelProviderFeatures.NamedListValidation));
        Assert.False(importer.SupportsCompleteWorkbookExport);
    }

    /// <summary>
    /// 验证 XLS 和 XLSX 多工作表分批导入按物理顺序交付批次。
    /// </summary>
    /// <param name="format">输入工作簿格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void ImportWorkbookBatches_ShouldDeliverSheetsInOrder(ExcelFormat format)
    {
        using var source = CreateWorkbook(format, includeSecondSheet: true);
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .Sheet<Person>("People", root => root.People, sheet => sheet.BatchSize(1))
            .Sheet<Person>("Other", root => root.Other, sheet => sheet.BatchSize(1)));

        var summary = new ExcelDataReaderExcelImporter().ImportWorkbookBatches(source, request, batches.Add);

        Assert.True(summary.IsSuccess);
        Assert.Equal(new[] { "People", "Other" }, summary.Sheets.Select(sheet => sheet.SheetName));
        Assert.Equal(new[] { 1, 2, 3, 4 }, batches.Select(batch => batch.BatchNumber));
        Assert.Equal(new[] { "People", "People", "Other", "Other" },
            batches.Select(batch => batch.SheetName));
        Assert.Equal(new[] { "Alice", "Bob", "Carol", "Dan" },
            batches.SelectMany(batch => batch.Items).Cast<Person>().Select(item => item.Name));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 异构工作表应通过实体类型和对象批次交付完整数据。
    /// </summary>
    [Fact]
    public void ImportWorkbookBatches_ShouldSupportHeterogeneousSheetTypes()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("People", new[] { new Person { Name = "Alice", Age = 30 } })
            .AddSheet("Notes", new[]
            {
                new RichPerson { Name = "Memo", RegisteredAt = new DateTime(2026, 9, 25), Note = "文本" }
            }));
        using var source = new MemoryStream();
        new Bing.Offices.Exports.NpoiExcelExporter().Export(request, source);
        source.Position = 0;
        var batches = new List<ExcelWorkbookImportBatch>();
        var importRequest = ExcelImport.Workbook<HeterogeneousWorkbook>(workbook => workbook
            .Sheet<Person>("People", root => root.People)
            .Sheet<RichPerson>("Notes", root => root.Notes));

        var summary = new ExcelDataReaderExcelImporter().ImportWorkbookBatches(source, importRequest, batches.Add);

        Assert.True(summary.IsSuccess);
        Assert.Equal(new[] { typeof(Person), typeof(RichPerson) }, batches.Select(batch => batch.ItemType));
        Assert.Equal("Alice", Assert.IsType<Person>(Assert.Single(batches[0].Items)).Name);
        Assert.Equal("文本", Assert.IsType<RichPerson>(Assert.Single(batches[1].Items)).Note);
    }

    /// <summary>
    /// 多工作表分批导入应共享行数预算并交付资源错误批次。
    /// </summary>
    [Fact]
    public void ImportWorkbookBatches_ShouldShareResourceBudgetAcrossSheets()
    {
        using var source = CreateWorkbook(ExcelFormat.Xlsx, includeSecondSheet: true);
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 2 })
            .Sheet<Person>("People", root => root.People)
            .Sheet<Person>("Other", root => root.Other));

        var summary = new ExcelDataReaderExcelImporter().ImportWorkbookBatches(source, request, batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(2, summary.RowsRead);
        Assert.True(summary.ResourceLimitExceeded);
        Assert.Contains(batches, batch => batch.SheetName == "Other"
            && batch.Errors.Any(error => error.Code == ExcelImportErrorCode.ResourceLimit));
    }

    /// <summary>
    /// 异步多工作表导入应等待每个批次回调并保持调用方流可用。
    /// </summary>
    [Fact]
    public async Task ImportWorkbookBatchesAsync_ShouldAwaitEachBatchAndPreserveSource()
    {
        using var source = CreateWorkbook(ExcelFormat.Xlsx, includeSecondSheet: true);
        var observed = new List<string>();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .Sheet<Person>("People", root => root.People, sheet => sheet.BatchSize(1))
            .Sheet<Person>("Other", root => root.Other, sheet => sheet.BatchSize(1)));

        var summary = await new ExcelDataReaderExcelImporter().ImportWorkbookBatchesAsync(source, request,
            async (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                observed.Add($"{batch.BatchNumber}:{batch.SheetName}");
                await Task.Yield();
            });

        Assert.True(summary.IsSuccess);
        Assert.Equal(new[] { "1:People", "2:People", "3:Other", "4:Other" }, observed);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// Workbook 分批 ErrorRowsOnly 应保留失败行并生成最小 XLSX。
    /// </summary>
    /// <param name="format">输入工作簿格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void ImportWorkbookBatches_ErrorRowsOnly_ShouldWriteFailureRows(ExcelFormat format)
    {
        using var source = CreateWorkbook(format, people: new[]
        {
            new Person { Name = string.Empty, Age = 31 },
            new Person { Name = "Bob", Age = 32 }
        });
        using var destination = new MemoryStream();
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People, sheet => sheet
                .BatchSize(1)
                .Validate(ExcelValidationFailureMode.Continue)));

        var summary = new ExcelDataReaderExcelImporter().ImportWorkbookBatches(source, request, batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Contains(batches, batch => batch.Errors.Count > 0);
        Assert.True(destination.Length > 0);
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var errors = workbook.GetSheet("__BingErrors");
        Assert.NotNull(errors);
        Assert.Equal("People", errors.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal(2d, errors.GetRow(1).GetCell(1).NumericCellValue);
        var people = workbook.GetSheet("People");
        Assert.NotNull(people);
        Assert.Equal("Name", people.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal(string.Empty, people.GetRow(1).GetCell(0).StringCellValue);
        Assert.Contains("必填", people.GetRow(1).GetCell(2).StringCellValue);
    }

    /// <summary>
    /// Workbook 分批异步 ErrorRowsOnly 应等待回调后生成失败工作簿并保持目标流可用。
    /// </summary>
    [Fact]
    public async Task ImportWorkbookBatchesAsync_ErrorRowsOnly_ShouldWriteFailureRowsAndKeepDestinationOpen()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var result = await new ExcelDataReaderExcelImporter().ImportWorkbookBatchesAsync(source, request,
            async (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                await Task.Yield();
            });

        Assert.False(result.IsSuccess);
        Assert.True(destination.CanRead);
        Assert.True(destination.Length > 0);
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        Assert.NotNull(workbook.GetSheet("__BingErrors"));
    }

    /// <summary>
    /// Workbook 分批目标路径应使用宿主注册的文件提交器。
    /// </summary>
    [Fact]
    public void ImportWorkbookBatches_ErrorRowsOnly_DestinationPath_ShouldUseRegisteredCommitter()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        var committer = new RecordingCommitter();
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-edr-batch-{Guid.NewGuid():N}.xlsx");
        try
        {
            using var provider = new ServiceCollection()
                .AddSingleton<IFileExportCommitter>(committer)
                .AddBingOfficesExcelDataReader()
                .BuildServiceProvider();
            var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .FailureWorkbook(new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                    DestinationPath = path
                })
                .Sheet<Person>("People", root => root.People));

            var summary = provider.GetRequiredService<IExcelWorkbookBatchImporter>()
                .ImportWorkbookBatches(source, request, _ => { });

            Assert.False(summary.IsSuccess);
            Assert.Equal("xlsx", committer.Format);
            Assert.True(File.Exists(path));
            Assert.True(new FileInfo(path).Length > 0);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 真实 XLSB 文件应完成固定列完整导入和分批导入。
    /// </summary>
    [Fact]
    public void Import_ShouldReadRealXlsbFixtureAndDeliverBatches()
    {
        using var source = OpenFixture("xlsb/Issue635.xlsb");
        var importer = new ExcelDataReaderExcelImporter();
        var request = ExcelImport.Workbook<XlsbWorkbook>(workbook => workbook
            .Sheet<XlsbRow>("Лист2", root => root.Rows));

        var result = importer.Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var row = Assert.Single(result.Workbook.Rows);
        Assert.Equal("123.456", row.A);
        Assert.False(row.B);
        Assert.Equal(new DateTime(2023, 6, 9), row.C);
        Assert.Equal("Plaintext", row.D);
        Assert.True(source.CanRead);

        source.Position = 0;
        var batches = new List<ExcelImportBatch<XlsbRow>>();
        var summary = importer.ImportBatches(source, new ExcelBatchImportRequest<XlsbRow>("Лист2"),
            batches.Add);
        Assert.True(summary.IsSuccess);
        Assert.Equal("Лист2", Assert.Single(batches).SheetName);
        Assert.Equal("123.456", Assert.Single(Assert.Single(batches).Items).A);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSB 物理单元格预算在预检阶段被拒绝。
    /// </summary>
    [Fact]
    public void Import_XlsbPhysicalCellLimits_ShouldBeExplicitlyUnsupported()
    {
        using var source = OpenFixture("xlsb/Issue635.xlsb");
        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ExcelDataReaderExcelImporter().Import(source,
                ExcelImport.Workbook<XlsbWorkbook>(workbook => workbook
                    .ResourceLimits(new ExcelResourceLimits { MaxCells = 1 })
                    .Sheet<XlsbRow>("Лист2", root => root.Rows))));

        Assert.Equal("ExcelDataReader", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("物理单元格", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证导入 NPOI 生成的 XLSX 并映射人员数据。
    /// </summary>
    [Fact]
    public void Import_ShouldReadNpoiGeneratedXlsxAndMapValues()
    {
        using var source = CreateWorkbook(ExcelFormat.Xlsx);
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "Alice", "Bob" }, result.Workbook.People.Select(item => item.Name));
        Assert.Equal(new[] { 30, 31 }, result.Workbook.People.Select(item => item.Age));
    }

    /// <summary>
    /// 验证导入 NPOI 生成的 XLS 人员数据。
    /// </summary>
    [Fact]
    public void Import_ShouldReadNpoiGeneratedXls()
    {
        using var source = CreateWorkbook(ExcelFormat.Xls);
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "Alice", "Bob" }, result.Workbook.People.Select(item => item.Name));
    }

    /// <summary>
    /// 固定列导入应保留中文、日期和空值。
    /// </summary>
    [Fact]
    public void Import_ShouldPreserveUnicodeDateAndEmptyValue()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Data", new[]
            {
                new RichPerson { Name = "张三", RegisteredAt = new DateTime(2026, 9, 25), Note = null }
            }));
        using var source = new MemoryStream();
        new Bing.Offices.Exports.NpoiExcelExporter().Export(request, source);
        source.Position = 0;

        var importRequest = ExcelImport.Workbook<RichImportWorkbook>(workbook =>
            workbook.Sheet<RichPerson>("Data", root => root.Items));
        var result = new ExcelDataReaderExcelImporter().Import(source, importRequest);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var item = Assert.Single(result.Workbook.Items);
        Assert.Equal("张三", item.Name);
        Assert.Equal(new DateTime(2026, 9, 25), item.RegisteredAt);
        Assert.Null(item.Note);
    }

    /// <summary>
    /// 1900 与 1904 日期系统都应还原为同一日期值。
    /// </summary>
    /// <param name="relativePath">相对于黄金样例目录的文件路径。</param>
    [Theory]
    [InlineData("date/date1900.xlsx")]
    [InlineData("date/date1904.xlsx")]
    public void Import_ShouldRespectWorkbookDateSystem(string relativePath)
    {
        using var source = OpenFixture(relativePath);
        var request = ExcelImport.Workbook<DateImportWorkbook>(workbook =>
            workbook.Sheet<DateValue>("Data", root => root.Items));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new DateTime(2026, 9, 4), Assert.Single(result.Workbook.Items).Date);
    }

    /// <summary>
    /// 公式字段按 ExcelDataReader 提供的缓存值读取，不尝试重新计算。
    /// </summary>
    [Fact]
    public void Import_ShouldReadCachedFormulaValue()
    {
        using var source = OpenFixture("formula/cached-number.xlsx");
        var request = ExcelImport.Workbook<FormulaImportWorkbook>(workbook =>
            workbook.Sheet<FormulaValue>("Formula", root => root.Items));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(2, Assert.Single(result.Workbook.Items).Value);
    }

    /// <summary>
    /// 验证工作簿行数超限时不返回部分实体。
    /// </summary>
    [Fact]
    public void Import_ShouldApplyWorkbookRowBudgetWithoutPartialEntities()
    {
        using var source = CreateWorkbook();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 1 })
            .Sheet<Person>("People", root => root.People));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.People);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);
    }

    /// <summary>
    /// 验证跨表共享行数预算且超限时清除先前导入实体。
    /// </summary>
    [Fact]
    public void Import_ShouldShareRowBudgetAcrossSheetsAndHideEarlierEntities()
    {
        using var source = CreateWorkbook(ExcelFormat.Xlsx, includeSecondSheet: true);
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 2 })
            .Sheet<Person>("People", root => root.People)
            .Sheet<Person>("Other", root => root.Other));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.People);
        Assert.Empty(result.Workbook.Other);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);
    }

    /// <summary>
    /// 验证分批导入交付完整批次且保持源流可用。
    /// </summary>
    [Fact]
    public void ImportBatches_ShouldDeliverBatchesAndKeepSourceOwnedByCaller()
    {
        using var source = CreateWorkbook();
        var batches = new List<ExcelImportBatch<Person>>();
        var request = new ExcelBatchImportRequest<Person>("People") { BatchSize = 1 };

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source, request, batches.Add);

        Assert.True(summary.IsSuccess);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(2, batches.SelectMany(batch => batch.Items).Count());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证异步导入等待每个批次回调完成。
    /// </summary>
    [Fact]
    public async Task ImportBatchesAsync_ShouldAwaitEachBatchAndHonorCancellation()
    {
        using var source = CreateWorkbook();
        var count = 0;
        var request = new ExcelBatchImportRequest<Person>("People") { BatchSize = 1 };

        var summary = await new ExcelDataReaderExcelImporter().ImportBatchesAsync(source, request,
            async (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                count += batch.Items.Count;
                await Task.Yield();
            });

        Assert.True(summary.IsSuccess);
        Assert.Equal(2, count);
    }

    /// <summary>
    /// 批次大小应稳定分组，行数预算只允许交付预算内的完整实体。
    /// </summary>
    /// <param name="rowCount">生成的测试数据行数。</param>
    /// <param name="expectedResourceLimitExceeded">是否预期超出工作簿行数预算。</param>
    [Theory]
    [InlineData(4, false)]
    [InlineData(5, true)]
    public void ImportBatches_BatchSizeAndMaxRowsBoundary_ShouldDeliverOnlyBudgetedRows(
        int rowCount, bool expectedResourceLimitExceeded)
    {
        var rows = Enumerable.Range(1, rowCount)
            .Select(index => new Person { Name = $"Person-{index}", Age = 20 + index })
            .ToArray();
        using var source = CreateWorkbook(people: rows);
        var batches = new List<ExcelImportBatch<Person>>();
        var request = new ExcelBatchImportRequest<Person>("People")
        {
            BatchSize = 2,
            ResourceLimits = new ExcelResourceLimits { MaxRows = 4 }
        };

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source, request, batches.Add);

        Assert.Equal(new[] { 2, 2 }, batches.Select(batch => batch.Items.Count));
        Assert.Equal(new[] { "Person-1", "Person-2", "Person-3", "Person-4" },
            batches.SelectMany(batch => batch.Items).Select(item => item.Name));
        Assert.All(batches, batch => Assert.Empty(batch.Errors));
        Assert.Equal(4, summary.RowsRead);
        Assert.Equal(4, summary.SucceededRows);
        Assert.Equal(0, summary.FailedRows);
        Assert.Equal(expectedResourceLimitExceeded, summary.ResourceLimitExceeded);
        Assert.Equal(!expectedResourceLimitExceeded, summary.IsSuccess);
    }

    /// <summary>
    /// Continue 模式应交付错误批次，并继续交付错误行之后的有效实体。
    /// </summary>
    [Fact]
    public void ImportBatches_ContinueMode_ShouldReturnErrorRowWithItsBatch()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = string.Empty, Age = 31 },
            new Person { Name = "Bob", Age = 32 }
        });
        var batches = new List<ExcelImportBatch<Person>>();
        var request = new ExcelBatchImportRequest<Person>("People")
        {
            BatchSize = 1,
            ValidationFailureMode = ExcelValidationFailureMode.Continue
        };

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source, request, batches.Add);

        Assert.Equal(3, batches.Count);
        Assert.Equal("Alice", Assert.Single(batches[0].Items).Name);
        Assert.Empty(batches[0].Errors);
        Assert.Empty(batches[1].Items);
        var error = Assert.Single(batches[1].Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal(3, error.RowIndex);
        Assert.Equal(nameof(Person.Name), error.PropertyName);
        Assert.Equal("Bob", Assert.Single(batches[2].Items).Name);
        Assert.Empty(batches[2].Errors);
        Assert.Equal(3, summary.RowsRead);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(1, summary.FailedRows);
        Assert.Equal(1, summary.ErrorCount);
    }

    /// <summary>
    /// 单工作表分批导入的 XLS 输入应在预检阶段拒绝 AnnotatedOriginal，并保持目标流不变。
    /// </summary>
    [Fact]
    public void ImportBatches_AnnotatedOriginal_Xls_ShouldRejectBeforeCreatingOutput()
    {
        using var source = CreateWorkbook(ExcelFormat.Xls);
        using var destination = new MemoryStream();
        var request = new ExcelBatchImportRequest<Person>("People")
        {
            FailureOptions = new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination
            }
        };

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ExcelDataReaderExcelImporter().ImportBatches(source, request, _ => { }));

        Assert.Equal("ExcelDataReader", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("XLSX", exception.Message);
        Assert.Equal(0, destination.Length);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 单工作表分批导入的 XLSX 输入应保留原工作簿并写入错误批注。
    /// </summary>
    [Fact]
    public void ImportBatches_AnnotatedOriginal_Xlsx_ShouldPreserveSourceAndAnnotateErrors()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        using var destination = new MemoryStream();
        var request = new ExcelBatchImportRequest<Person>("People")
        {
            FailureOptions = new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination
            }
        };

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source, request, _ => { });

        Assert.False(summary.IsSuccess);
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var people = workbook.GetSheet("People");
        Assert.NotNull(people);
        Assert.Equal(string.Empty, people.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal(31d, people.GetRow(1).GetCell(1).NumericCellValue);
        Assert.NotNull(people.GetRow(1).GetCell(0).CellComment);
        Assert.Contains("必填", people.GetRow(1).GetCell(0).CellComment.String.String);
        Assert.NotNull(workbook.GetSheet("__BingErrors"));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 单工作表同步分批失败工作簿应通过宿主提交器写入目标路径。
    /// </summary>
    [Fact]
    public void ImportBatches_ErrorRowsOnly_DestinationPath_ShouldUseRegisteredCommitter()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        var committer = new RecordingCommitter();
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-edr-batch-{Guid.NewGuid():N}.xlsx");
        try
        {
            using var provider = new ServiceCollection()
                .AddSingleton<IFileExportCommitter>(committer)
                .AddBingOfficesExcelDataReader()
                .BuildServiceProvider();
            var request = new ExcelBatchImportRequest<Person>("People")
            {
                FailureOptions = new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                    DestinationPath = path
                }
            };

            var summary = provider.GetRequiredService<IExcelBatchImporter>().ImportBatches(source, request,
                _ => { });

            Assert.False(summary.IsSuccess);
            Assert.Equal("xlsx", committer.Format);
            Assert.True(File.Exists(path));
            Assert.True(new FileInfo(path).Length > 0);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 单工作表异步分批失败工作簿应通过宿主提交器异步写入目标路径。
    /// </summary>
    [Fact]
    public async Task ImportBatchesAsync_ErrorRowsOnly_DestinationPath_ShouldUseRegisteredCommitter()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        var committer = new RecordingCommitter();
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-edr-batch-{Guid.NewGuid():N}.xlsx");
        try
        {
            using var provider = new ServiceCollection()
                .AddSingleton<IFileExportCommitter>(committer)
                .AddBingOfficesExcelDataReader()
                .BuildServiceProvider();
            var request = new ExcelBatchImportRequest<Person>("People")
            {
                FailureOptions = new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                    DestinationPath = path
                }
            };

            var summary = await provider.GetRequiredService<IExcelBatchImporter>().ImportBatchesAsync(source,
                request, async (_, cancellationToken) =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    await Task.Yield();
                });

            Assert.False(summary.IsSuccess);
            Assert.Equal("xlsx", committer.Format);
            Assert.True(File.Exists(path));
            Assert.True(new FileInfo(path).Length > 0);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 单工作表同步分批导入应将失败行和结构化错误写入 XLSX，并保持目标流打开。
    /// </summary>
    /// <param name="format">输入工作簿格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void ImportBatches_ErrorRowsOnly_ShouldWriteFailureRowsAndKeepDestinationOpen(ExcelFormat format)
    {
        using var source = CreateWorkbook(format, people: new[]
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = string.Empty, Age = 31 }
        });
        using var destination = new MemoryStream();
        var batches = new List<ExcelImportBatch<Person>>();
        var request = new ExcelBatchImportRequest<Person>("People")
        {
            BatchSize = 1,
            ValidationFailureMode = ExcelValidationFailureMode.Continue,
            FailureOptions = new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            }
        };

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source, request, batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(1, summary.SucceededRows);
        Assert.Equal(1, summary.FailedRows);
        Assert.Equal(1, summary.ErrorCount);
        Assert.True(destination.CanRead);
        Assert.True(destination.Length > 0);

        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var summarySheet = workbook.GetSheet("__BingErrors");
        Assert.NotNull(summarySheet);
        Assert.Equal("Sheet", summarySheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Row", summarySheet.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal("People", summarySheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal(3d, summarySheet.GetRow(1).GetCell(1).NumericCellValue);
        Assert.Equal(nameof(ExcelImportErrorCode.Validation), summarySheet.GetRow(1).GetCell(3).StringCellValue);

        var people = workbook.GetSheet("People");
        Assert.NotNull(people);
        Assert.Equal("Name", people.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Age", people.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal("__Errors", people.GetRow(0).GetCell(2).StringCellValue);
        Assert.Equal(string.Empty, people.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal(31d, people.GetRow(1).GetCell(1).NumericCellValue);
        Assert.Contains("必填", people.GetRow(1).GetCell(2).StringCellValue);
    }

    /// <summary>
    /// 单工作表异步分批导入应等待回调并写出可读取的失败工作簿。
    /// </summary>
    /// <param name="format">输入工作簿格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public async Task ImportBatchesAsync_ErrorRowsOnly_ShouldWriteFailureRowsAndKeepDestinationOpen(
        ExcelFormat format)
    {
        using var source = CreateWorkbook(format, people: new[]
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = string.Empty, Age = 31 }
        });
        using var destination = new MemoryStream();
        var callbackCount = 0;
        var request = new ExcelBatchImportRequest<Person>("People")
        {
            BatchSize = 1,
            ValidationFailureMode = ExcelValidationFailureMode.Continue,
            FailureOptions = new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            }
        };

        var summary = await new ExcelDataReaderExcelImporter().ImportBatchesAsync(source, request,
            async (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                callbackCount++;
                await Task.Yield();
            });

        Assert.False(summary.IsSuccess);
        Assert.Equal(2, callbackCount);
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(1, summary.SucceededRows);
        Assert.Equal(1, summary.FailedRows);
        Assert.Equal(1, summary.ErrorCount);
        Assert.True(destination.CanRead);
        Assert.True(destination.Length > 0);

        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var summarySheet = workbook.GetSheet("__BingErrors");
        Assert.NotNull(summarySheet);
        Assert.Equal("Sheet", summarySheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Row", summarySheet.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal("People", summarySheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal(3d, summarySheet.GetRow(1).GetCell(1).NumericCellValue);
        Assert.Equal(nameof(ExcelImportErrorCode.Validation), summarySheet.GetRow(1).GetCell(3).StringCellValue);

        var people = workbook.GetSheet("People");
        Assert.NotNull(people);
        Assert.Equal("Name", people.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Age", people.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal("__Errors", people.GetRow(0).GetCell(2).StringCellValue);
        Assert.Equal(string.Empty, people.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal(31d, people.GetRow(1).GetCell(1).NumericCellValue);
        Assert.Contains("必填", people.GetRow(1).GetCell(2).StringCellValue);
    }

    /// <summary>
    /// 异步导入在真实输入读取后取消时，不得释放调用方拥有的源流。
    /// </summary>
    [Fact]
    public async Task ImportAsync_MidReadCancellation_ShouldKeepCallerSourceOpen()
    {
        using var generated = CreateWorkbook();
        using var cancellation = new CancellationTokenSource();
        using var source = new CancelAfterFirstReadStream(generated, cancellation);
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new ExcelDataReaderExcelImporter().ImportAsync(source, request, cancellation.Token));

        Assert.True(source.HasRead);
        Assert.True(cancellation.IsCancellationRequested);
        Assert.False(source.WasDisposed);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 损坏输入应转换为带 Provider 和阶段元数据的公共异常。
    /// </summary>
    [Fact]
    public void Import_MalformedWorkbook_ShouldReturnStructuredProviderException()
    {
        using var source = new MemoryStream(new byte[] { 0x42, 0x49, 0x4E, 0x47, 0x00, 0x01 });
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesImportException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Equal("ExcelDataReader", exception.Provider);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(BingOfficesStage.Read, exception.Stage);
        Assert.Equal(BingOfficesErrorCode.ImportFailed, exception.Code);
        Assert.NotNull(exception.InnerException);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// AnnotatedOriginal 应保留原工作簿并写入错误摘要和批注。
    /// </summary>
    [Fact]
    public void Import_AnnotatedOriginal_ShouldPreserveSourceAndAnnotateErrors()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var people = workbook.GetSheet("People");
        Assert.NotNull(people);
        Assert.Equal(31d, people.GetRow(1).GetCell(1).NumericCellValue);
        Assert.NotNull(people.GetRow(1).GetCell(0).CellComment);
        Assert.Contains("必填", people.GetRow(1).GetCell(0).CellComment.String.String);
        Assert.NotNull(workbook.GetSheet("__BingErrors"));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// AnnotatedOriginal 异步导入应保留原工作簿并保持目标流可用。
    /// </summary>
    [Fact]
    public async Task ImportAsync_AnnotatedOriginal_ShouldPreserveSourceAndAnnotateErrors()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var result = await new ExcelDataReaderExcelImporter().ImportAsync(source, request);

        Assert.False(result.IsSuccess);
        Assert.True(destination.CanRead);
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        Assert.NotNull(workbook.GetSheet("People").GetRow(1).GetCell(0).CellComment);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// AnnotatedOriginal 的批注冲突策略为 Fail 时应保持原目标内容不变。
    /// </summary>
    [Fact]
    public void Import_AnnotatedOriginal_CommentConflictFail_ShouldKeepDestinationUnchanged()
    {
        using var source = CreateWorkbookWithComment();
        using var destination = new MemoryStream(new byte[] { 7, 8, 9 });
        var original = destination.ToArray();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination,
                CommentConflictPolicy = ExcelImportCommentConflictPolicy.Fail
            })
            .Sheet<Person>("People", root => root.People));

        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Equal(original, destination.ToArray());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// ErrorRowsOnly 应保留失败行、原始表头和结构化错误摘要。
    /// </summary>
    /// <param name="format">输入工作簿格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Import_ErrorRowsOnly_ShouldWriteFailureRowsToXlsx(ExcelFormat format)
    {
        using var source = CreateWorkbook(format, people: new[]
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = string.Empty, Age = 31 },
            new Person { Name = "Bob", Age = 32 }
        });
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.Single(result.Errors);
        Assert.True(destination.Length > 0);
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var errors = workbook.GetSheet("__BingErrors");
        Assert.NotNull(errors);
        Assert.Equal("Sheet", errors.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("People", errors.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal(3d, errors.GetRow(1).GetCell(1).NumericCellValue);
        var people = workbook.GetSheet("People");
        Assert.NotNull(people);
        Assert.Equal("Name", people.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Age", people.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal("__Errors", people.GetRow(0).GetCell(2).StringCellValue);
        Assert.Equal(string.Empty, people.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal(31d, people.GetRow(1).GetCell(1).NumericCellValue);
        Assert.Contains("必填", people.GetRow(1).GetCell(2).StringCellValue);
    }

    /// <summary>
    /// ErrorRowsOnly 应保留源数据列宽和隐藏状态，但不隐藏新增错误列。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_XlsxColumnMetadata_ShouldPreserveSourceColumns()
    {
        using var source = CreateWorkbookWithColumnMetadata();
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.Single(result.Errors);
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var people = workbook.GetSheet("People");
        Assert.NotNull(people);
        Assert.Equal(22 * 256, people.GetColumnWidth(0));
        Assert.Equal(16 * 256, people.GetColumnWidth(1));
        Assert.True(people.IsColumnHidden(1));
        Assert.False(people.IsColumnHidden(2));
        Assert.Equal("__Errors", people.GetRow(0).GetCell(2).StringCellValue);
    }

    /// <summary>
    /// 列元数据超出目标对象预算时应保持原目标内容不变。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_XlsxColumnMetadata_ObjectBudget_ShouldKeepDestinationUnchanged()
    {
        using var source = CreateWorkbookWithColumnMetadata();
        var original = new byte[] { 9, 8, 7 };
        using var destination = new MemoryStream(original, writable: true);
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination,
                MaxEstimatedTargetObjects = 32
            })
            .Sheet<Person>("People", root => root.People));

        Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Equal(original, destination.ToArray());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// ErrorRowsOnly 遇到列样式元数据时应在输出前结构化拒绝。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_XlsxColumnMetadata_UnsupportedAttribute_ShouldRejectBeforeWriting()
    {
        using var source = CreateWorkbookWithColumnMetadata(unsupportedColumnStyle: true);
        var original = new byte[] { 6, 5, 4 };
        using var destination = new MemoryStream(original, writable: true);
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("列元数据", exception.Message);
        Assert.Equal(original, destination.ToArray());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ErrorRowsOnly 遇到条件格式时在输出前结构化拒绝。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_XlsxConditionalFormatting_ShouldRejectBeforeWriting()
    {
        using var source = CreateWorkbookWithConditionalFormatting(
            unresolvedDifferentialStyle: true);
        using var destination = new MemoryStream(new byte[] { 4, 5, 6 });
        var original = destination.ToArray();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("条件格式", exception.Message);
        Assert.Equal(original, destination.ToArray());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证源坐标未发生压缩时 ErrorRowsOnly 会保留条件格式。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_XlsxConditionalFormatting_ShouldPreserveStableRows()
    {
        using var source = CreateWorkbookWithConditionalFormatting(includeSuccessfulRow: false);
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.Single(result.Errors);
        destination.Position = 0;
        using var archive = new ZipArchive(destination, ZipArchiveMode.Read, leaveOpen: true);
        var sheet = archive.GetEntry("xl/worksheets/sheet2.xml");
        Assert.NotNull(sheet);
        using var sheetStream = sheet.Open();
        var document = XDocument.Load(sheetStream);
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var conditionalFormatting = Assert.Single(document.Root.Elements(spreadsheet + "conditionalFormatting"));
        Assert.Equal("B2", (string)conditionalFormatting.Attribute("sqref"));
        Assert.NotNull(conditionalFormatting.Element(spreadsheet + "cfRule"));
    }

    /// <summary>
    /// 验证条件格式覆盖被压缩的源行时在输出前结构化拒绝。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_XlsxConditionalFormatting_CompressedRows_ShouldRejectBeforeWriting()
    {
        using var source = CreateWorkbookWithConditionalFormatting();
        using var destination = new MemoryStream(new byte[] { 8, 9, 10 });
        var original = destination.ToArray();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("条件格式", exception.Message);
        Assert.Equal(original, destination.ToArray());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证包含外部工作表引用的条件格式在输出前结构化拒绝。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_XlsxConditionalFormatting_ExternalFormula_ShouldRejectBeforeWriting()
    {
        using var source = CreateWorkbookWithConditionalFormatting(formula: "Other!A1>0");
        using var destination = new MemoryStream(new byte[] { 11, 12, 13 });
        var original = destination.ToArray();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("条件格式", exception.Message);
        Assert.Equal(original, destination.ToArray());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 资源预检提前失败时仍应拒绝包含条件格式的低保留输出。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_XlsxConditionalFormatting_ResourceLimit_ShouldRejectBeforeWriting()
    {
        using var source = CreateWorkbookWithConditionalFormatting(
            unresolvedDifferentialStyle: true);
        using var destination = new MemoryStream(new byte[] { 7, 8, 9 });
        var original = destination.ToArray();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxCells = 1 })
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("条件格式", exception.Message);
        Assert.Equal(original, destination.ToArray());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// ErrorRowsOnly 异步输出应保留调用方目标流并生成可读取的 XLSX。
    /// </summary>
    [Fact]
    public async Task ImportAsync_ErrorRowsOnly_ShouldWriteFailureRowsAndKeepDestinationOpen()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var result = await new ExcelDataReaderExcelImporter().ImportAsync(source, request);

        Assert.False(result.IsSuccess);
        Assert.True(destination.CanRead);
        Assert.True(destination.Length > 0);
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        Assert.NotNull(workbook.GetSheet("__BingErrors"));
    }

    /// <summary>
    /// ErrorRowsOnly 应在同步和异步 XLSX 输出中重映射可安全复制的原生校验区域。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Import_ErrorRowsOnly_Xlsx_ShouldRemapNativeValidationMetadata(
        bool asynchronous)
    {
        using var source = CreateWorkbookWithNativeValidationFailure();
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var result = asynchronous
            ? await new ExcelDataReaderExcelImporter().ImportAsync(source, request)
            : new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.Single(result.Errors);
        destination.Position = 0;
        using var archive = new ZipArchive(destination, ZipArchiveMode.Read, leaveOpen: true);
        var sheetEntry = archive.GetEntry("xl/worksheets/sheet2.xml");
        Assert.NotNull(sheetEntry);
        using var sheetStream = sheetEntry.Open();
        var document = XDocument.Load(sheetStream);
        XNamespace spreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var validation = document.Descendants(spreadsheetNamespace + "dataValidation").Single();
        Assert.Equal("list", (string)validation.Attribute("type"));
        Assert.Equal("A2", (string)validation.Attribute("sqref"));
        Assert.NotNull(validation.Element(spreadsheetNamespace + "formula1"));
    }

    /// <summary>
    /// 验证通用前向读取路径只重映射安全的全局和局部定义名称。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Import_ErrorRowsOnly_Xlsx_ShouldRemapSafeDefinedNames(
        bool asynchronous)
    {
        using var baseSource = CreateWorkbookWithDefinedNamesFailure();
        using var source = RewriteXlsxEntry(baseSource,
            name => string.Equals(name, "xl/workbook.xml", StringComparison.OrdinalIgnoreCase),
            AddHiddenDefinedName);
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var result = asynchronous
            ? await new ExcelDataReaderExcelImporter().ImportAsync(source, request)
            : new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        destination.Position = 0;
        using var archive = new ZipArchive(destination, ZipArchiveMode.Read, leaveOpen: true);
        var workbookEntry = archive.GetEntry("xl/workbook.xml");
        Assert.NotNull(workbookEntry);
        using var workbookStream = workbookEntry.Open();
        var document = XDocument.Load(workbookStream);
        var names = document.Descendants().Where(element =>
                string.Equals(element.Name.LocalName, "definedName", StringComparison.Ordinal))
            .ToDictionary(element => (string)element.Attribute("name"), StringComparer.OrdinalIgnoreCase);

        Assert.EndsWith("!$B$2", names["FailedQuantity"].Value, StringComparison.Ordinal);
        Assert.Equal("1", (string)names["LocalFailedQuantity"].Attribute("localSheetId"));
        Assert.EndsWith("!$B$2", names["LocalFailedQuantity"].Value, StringComparison.Ordinal);
        Assert.DoesNotContain("SuccessfulQuantity", names.Keys, StringComparer.OrdinalIgnoreCase);
        Assert.DoesNotContain("HiddenFailedQuantity", names.Keys, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 没有导入错误时不应创建失败工作簿。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_WithoutErrors_ShouldNotWriteDestination()
    {
        using var source = CreateWorkbook();
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 序列化字节预算失败时不得污染调用方已有目标内容。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_MaxSerializedBytes_ShouldKeepDestinationUnchanged()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        var original = new byte[] { 0x42, 0x49, 0x4E, 0x47 };
        using var destination = new MemoryStream(original, writable: true);
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination,
                MaxSerializedBytes = 1
            })
            .Sheet<Person>("People", root => root.People));

        Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));
        Assert.Equal(original, destination.ToArray());
    }

    /// <summary>
    /// AnnotatedOriginal 应在复制源图片前执行数量预算检查。
    /// </summary>
    [Fact]
    public void Import_AnnotatedOriginal_PictureCountBudget_ShouldKeepDestinationUnchanged()
    {
        using var source = CreateWorkbookWithPictures(2);
        using var destination = new MemoryStream(new byte[] { 7, 8, 9 });
        var original = destination.ToArray();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination,
                MaxCopiedPictures = 1
            })
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Contains("图片数量", exception.Message);
        Assert.Equal(original, destination.ToArray());
    }

    /// <summary>
    /// AnnotatedOriginal 应在复制源图片前执行图片字节预算检查。
    /// </summary>
    [Fact]
    public void Import_AnnotatedOriginal_PictureByteBudget_ShouldKeepDestinationUnchanged()
    {
        using var source = CreateWorkbookWithPictures(1);
        using var destination = new MemoryStream(new byte[] { 7, 8, 9 });
        var original = destination.ToArray();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination,
                MaxCopiedPictureBytes = 1
            })
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Contains("图片字节数", exception.Message);
        Assert.Equal(original, destination.ToArray());
    }

    /// <summary>
    /// AnnotatedOriginal 应拒绝包含外部关系目标的 XLSX 包。
    /// </summary>
    [Fact]
    public void Import_AnnotatedOriginal_ExternalRelationship_ShouldKeepDestinationUnchanged()
    {
        using var baseSource = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        using var source = RewriteXlsxEntry(baseSource,
            name => string.Equals(name, "xl/_rels/workbook.xml.rels", StringComparison.OrdinalIgnoreCase),
            AddExternalRelationship);
        using var destination = new MemoryStream(new byte[] { 7, 8, 9 });
        var original = destination.ToArray();
        var request = CreateAnnotatedOriginalRequest(destination);

        var exception = Assert.Throws<BingOfficesImportException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Contains("外部", exception.InnerException?.Message ?? exception.Message);
        Assert.Equal(original, destination.ToArray());
    }

    /// <summary>
    /// AnnotatedOriginal 应拒绝宏工作簿，避免破坏宏部件关系。
    /// </summary>
    [Fact]
    public void Import_AnnotatedOriginal_MacroWorkbook_ShouldKeepDestinationUnchanged()
    {
        using var baseSource = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        using var source = RewriteXlsxEntry(baseSource,
            name => string.Equals(name, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase),
            ConvertToMacroWorkbook);
        using var destination = new MemoryStream(new byte[] { 7, 8, 9 });
        var original = destination.ToArray();
        var request = CreateAnnotatedOriginalRequest(destination);

        var exception = Assert.Throws<BingOfficesImportException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Contains("宏", exception.InnerException?.Message ?? exception.Message);
        Assert.Equal(original, destination.ToArray());
    }

    /// <summary>
    /// AnnotatedOriginal 应拒绝包含非批注 VML 形状的工作簿。
    /// </summary>
    [Fact]
    public void Import_AnnotatedOriginal_ComplexVml_ShouldKeepDestinationUnchanged()
    {
        using var baseSource = CreateWorkbookWithComment();
        using var source = RewriteXlsxEntry(baseSource,
            name => name.EndsWith(".vml", StringComparison.OrdinalIgnoreCase),
            AddForeignVmlShape);
        using var destination = new MemoryStream(new byte[] { 7, 8, 9 });
        var original = destination.ToArray();
        var request = CreateAnnotatedOriginalRequest(destination);

        var exception = Assert.Throws<BingOfficesImportException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Contains("非批注", exception.InnerException?.Message ?? exception.Message);
        Assert.Equal(original, destination.ToArray());
    }

    /// <summary>
    /// 指定临时目录时同步失败工作簿应在提交后清理临时文件。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_TemporaryDirectory_ShouldBeCleaned()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        using var destination = new MemoryStream();
        var directory = Path.Combine(Path.GetTempPath(), $"bing-offices-edr-temp-{Guid.NewGuid():N}");
        try
        {
            var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .FailureWorkbook(new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                    Destination = destination,
                    TemporaryDirectory = directory
                })
                .Sheet<Person>("People", root => root.People));

            var result = new ExcelDataReaderExcelImporter().Import(source, request);

            Assert.False(result.IsSuccess);
            Assert.True(Directory.Exists(directory));
            Assert.Empty(Directory.EnumerateFileSystemEntries(directory));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 指定临时目录时异步失败工作簿应在提交后清理临时文件。
    /// </summary>
    [Fact]
    public async Task ImportAsync_ErrorRowsOnly_TemporaryDirectory_ShouldBeCleaned()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        using var destination = new MemoryStream();
        var directory = Path.Combine(Path.GetTempPath(), $"bing-offices-edr-temp-{Guid.NewGuid():N}");
        try
        {
            var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .FailureWorkbook(new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                    Destination = destination,
                    TemporaryDirectory = directory
                })
                .Sheet<Person>("People", root => root.People));

            var result = await new ExcelDataReaderExcelImporter().ImportAsync(source, request);

            Assert.False(result.IsSuccess);
            Assert.True(Directory.Exists(directory));
            Assert.Empty(Directory.EnumerateFileSystemEntries(directory));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 目标路径输出应使用宿主注册的文件提交器。
    /// </summary>
    [Fact]
    public void Import_ErrorRowsOnly_DestinationPath_ShouldUseRegisteredCommitter()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        var committer = new RecordingCommitter();
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-edr-{Guid.NewGuid():N}.xlsx");
        try
        {
            using var provider = new ServiceCollection()
                .AddSingleton<IFileExportCommitter>(committer)
                .AddBingOfficesExcelDataReader()
                .BuildServiceProvider();
            var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
                .FailureWorkbook(new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                    DestinationPath = path
                })
                .Sheet<Person>("People", root => root.People));

            var result = provider.GetRequiredService<IExcelImporter>().Import(source, request);

            Assert.False(result.IsSuccess);
            Assert.True(File.Exists(path));
            Assert.Equal("xlsx", committer.Format);
            Assert.True(new FileInfo(path).Length > 0);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 验证注册完整导入和分批导入服务。
    /// </summary>
    [Fact]
    public void AddBingOfficesExcelDataReader_ShouldRegisterReadAndBatchServices()
    {
        using var provider = new ServiceCollection()
            .AddBingOfficesExcelDataReader()
            .BuildServiceProvider();

        Assert.IsType<ExcelDataReaderExcelImporter>(provider.GetRequiredService<IExcelImporter>());
        Assert.IsType<ExcelDataReaderExcelImporter>(provider.GetRequiredService<IExcelBatchImporter>());
        Assert.IsType<ExcelDataReaderExcelImporter>(provider.GetRequiredService<IExcelWorkbookBatchImporter>());
    }

    /// <summary>
    /// 创建用于完整导入和分批导入的人员工作簿。
    /// </summary>
    /// <param name="format">生成的工作簿格式。</param>
    /// <param name="includeSecondSheet">是否生成第二个工作表。</param>
    /// <param name="people">主要工作表的数据；null 时使用默认人员数据。</param>
    /// <returns>包含测试工作簿内容且位于起始位置的流。</returns>
    private static MemoryStream CreateWorkbook(ExcelFormat format = ExcelFormat.Xlsx,
        bool includeSecondSheet = false, IEnumerable<Person> people = null)
    {
        var request = ExcelExport.Workbook(workbook =>
        {
            workbook.Format(format).AddSheet("People", people ?? new[]
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 31 }
            });
            if (includeSecondSheet)
                workbook.AddSheet("Other", new[]
                {
                    new Person { Name = "Carol", Age = 32 },
                    new Person { Name = "Dan", Age = 33 }
                });
        });
        var stream = new MemoryStream();
        new Bing.Offices.Exports.NpoiExcelExporter().Export(request, stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建包含显式列表校验和必填失败行的 XLSX 工作簿。
    /// </summary>
    /// <returns>位于起始位置的 XLSX 输入流。</returns>
    private static MemoryStream CreateWorkbookWithNativeValidationFailure()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("People");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Name");
        header.CreateCell(1).SetCellValue("Age");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("Alice");
        sheet.GetRow(1).CreateCell(1).SetCellValue(30);
        sheet.CreateRow(2).CreateCell(0).SetCellValue(string.Empty);
        sheet.GetRow(2).CreateCell(1).SetCellValue(31);
        var helper = sheet.GetDataValidationHelper();
        var validation = helper.CreateValidation(
            helper.CreateExplicitListConstraint(new[] { "Alice", "Bob" }),
            new CellRangeAddressList(1, 2, 0, 0));
        sheet.AddValidationData(validation);

        using var output = new MemoryStream();
        workbook.Write(output, false);
        return new MemoryStream(output.ToArray(), writable: false);
    }

    /// <summary>
    /// 创建包含显式列宽和隐藏列的失败 XLSX 工作簿。
    /// </summary>
    /// <param name="unsupportedColumnStyle">是否注入不支持的列样式属性。</param>
    /// <returns>位于起始位置的 XLSX 输入流。</returns>
    private static MemoryStream CreateWorkbookWithColumnMetadata(
        bool unsupportedColumnStyle = false)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("People");
        sheet.SetColumnWidth(0, 22 * 256);
        sheet.SetColumnWidth(1, 16 * 256);
        sheet.SetColumnHidden(1, true);
        sheet.SetColumnHidden(2, true);
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Name");
        header.CreateCell(1).SetCellValue("Age");
        sheet.CreateRow(1).CreateCell(0).SetCellValue(string.Empty);
        sheet.GetRow(1).CreateCell(1).SetCellValue(31);

        using var output = new MemoryStream();
        workbook.Write(output, false);
        var bytes = output.ToArray();
        if (unsupportedColumnStyle)
            bytes = SetWorksheetColumnAttribute(bytes, "style", "1");
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>
    /// 为测试夹具添加源工作表列属性。
    /// </summary>
    /// <param name="source">原始 XLSX 内容。</param>
    /// <param name="name">属性名称。</param>
    /// <param name="value">属性值。</param>
    /// <returns>更新后的 XLSX 内容。</returns>
    private static byte[] SetWorksheetColumnAttribute(byte[] source, string name, string value)
    {
        using var package = new MemoryStream();
        package.Write(source, 0, source.Length);
        package.Position = 0;
        using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/worksheets/sheet1.xml");
            Assert.NotNull(entry);
            string xml;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                xml = reader.ReadToEnd();

            var document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var column = document.Descendants(spreadsheet + "col").First();
            column.SetAttributeValue(name, value);

            entry.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet1.xml",
                CompressionLevel.Optimal);
            using var writer = new StreamWriter(replacement.Open(), new UTF8Encoding(false));
            writer.Write(document.Declaration == null
                ? document.ToString(SaveOptions.DisableFormatting)
                : document.Declaration + document.ToString(SaveOptions.DisableFormatting));
        }
        return package.ToArray();
    }

    /// <summary>
    /// 创建包含条件格式和失败数据的 XLSX 工作簿。
    /// </summary>
    /// <returns>位于起始位置的 XLSX 输入流。</returns>
    private static MemoryStream CreateWorkbookWithConditionalFormatting(
        bool includeSuccessfulRow = true, bool unresolvedDifferentialStyle = false,
        string formula = "0")
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("People");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Name");
        header.CreateCell(1).SetCellValue("Age");
        if (includeSuccessfulRow)
        {
            sheet.CreateRow(1).CreateCell(0).SetCellValue("Alice");
            sheet.GetRow(1).CreateCell(1).SetCellValue(30);
        }
        var failureRow = includeSuccessfulRow ? 2 : 1;
        sheet.CreateRow(failureRow).CreateCell(0).SetCellValue(string.Empty);
        sheet.GetRow(failureRow).CreateCell(1).SetCellValue(31);

        var formatting = sheet.SheetConditionalFormatting;
        var rule = formatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThan,
            formula);
        formatting.AddConditionalFormatting(new[] { new CellRangeAddress(
            includeSuccessfulRow ? 1 : failureRow, failureRow, 1, 1) }, rule);

        using var output = new MemoryStream();
        workbook.Write(output, false);
        var bytes = output.ToArray();
        if (unresolvedDifferentialStyle)
            bytes = SetConditionalFormattingDifferentialStyleId(bytes, "1");
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>
    /// 为测试夹具写入没有对应差异样式的条件格式引用。
    /// </summary>
    /// <param name="source">原始 XLSX 内容。</param>
    /// <param name="differentialStyleId">待写入的差异样式索引。</param>
    /// <returns>修改后的 XLSX 内容。</returns>
    private static byte[] SetConditionalFormattingDifferentialStyleId(byte[] source,
        string differentialStyleId)
    {
        using var package = new MemoryStream();
        package.Write(source, 0, source.Length);
        package.Position = 0;
        using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.Entries.Single(item => item.FullName == "xl/worksheets/sheet1.xml");
            string xml;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                xml = reader.ReadToEnd();

            var document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            foreach (var rule in document.Descendants(spreadsheet + "cfRule"))
                rule.SetAttributeValue("dxfId", differentialStyleId);

            entry.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet1.xml",
                CompressionLevel.Optimal);
            using var writer = new StreamWriter(replacement.Open(), new UTF8Encoding(false));
            writer.Write(document.Declaration == null
                ? document.ToString(SaveOptions.DisableFormatting)
                : document.Declaration + document.ToString(SaveOptions.DisableFormatting));
        }

        return package.ToArray();
    }

    /// <summary>
    /// 创建包含全局、局部、成功行和隐藏定义名称的 XLSX 工作簿。
    /// </summary>
    /// <returns>位于起始位置的 XLSX 输入流。</returns>
    private static MemoryStream CreateWorkbookWithDefinedNamesFailure()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("People");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Name");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Age");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("Alice");
        sheet.GetRow(1).CreateCell(1).SetCellValue(30);
        sheet.CreateRow(2).CreateCell(0).SetCellValue(string.Empty);
        sheet.GetRow(2).CreateCell(1).SetCellValue(31);

        var failed = workbook.CreateName();
        failed.NameName = "FailedQuantity";
        failed.RefersToFormula = "'People'!$B$3";
        var successful = workbook.CreateName();
        successful.NameName = "SuccessfulQuantity";
        successful.RefersToFormula = "'People'!$B$2";
        var local = workbook.CreateName();
        local.NameName = "LocalFailedQuantity";
        local.SheetIndex = workbook.GetSheetIndex(sheet);
        local.RefersToFormula = "'People'!$B$3";
        var hidden = workbook.CreateName();
        hidden.NameName = "HiddenFailedQuantity";
        hidden.RefersToFormula = "'People'!$B$3";

        using var output = new MemoryStream();
        workbook.Write(output, false);
        return new MemoryStream(output.ToArray(), writable: false);
    }

    /// <summary>
    /// 创建包含已有批注和失败数据的 XLSX 测试工作簿。
    /// </summary>
    /// <returns>位于起始位置的 XLSX 输入流。</returns>
    private static MemoryStream CreateWorkbookWithComment()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        source.Position = 0;
        using var workbook = WorkbookFactory.Create(source);
        var sheet = workbook.GetSheet("People");
        var drawing = sheet.CreateDrawingPatriarch();
        var anchor = workbook.GetCreationHelper().CreateClientAnchor();
        anchor.Col1 = 0;
        anchor.Row1 = 1;
        anchor.Col2 = 3;
        anchor.Row2 = 4;
        var comment = drawing.CreateCellComment(anchor);
        comment.String = new XSSFRichTextString("已有批注");
        comment.Author = "Existing";
        sheet.GetRow(1).GetCell(0).CellComment = comment;
        using var output = new MemoryStream();
        workbook.Write(output, false);
        return new MemoryStream(output.ToArray());
    }

    /// <summary>
    /// 创建包含指定数量图片和失败数据的 XLSX 测试工作簿。
    /// </summary>
    /// <param name="count">图片数量。</param>
    /// <returns>位于起始位置的 XLSX 输入流。</returns>
    private static MemoryStream CreateWorkbookWithPictures(int count)
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = string.Empty, Age = 31 }
        });
        source.Position = 0;
        using var workbook = (XSSFWorkbook)WorkbookFactory.Create(source);
        var sheet = workbook.GetSheet("People");
        var drawing = sheet.CreateDrawingPatriarch();
        for (var index = 0; index < count; index++)
        {
            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = index;
            anchor.Row1 = 0;
            anchor.Col2 = index + 1;
            anchor.Row2 = 1;
            var pictureIndex = workbook.AddPicture(new byte[] { 1, 2, (byte)(index + 3) },
                PictureType.PNG);
            drawing.CreatePicture(anchor, pictureIndex);
        }

        using var output = new MemoryStream();
        workbook.Write(output, false);
        return new MemoryStream(output.ToArray());
    }

    /// <summary>
    /// 创建 AnnotatedOriginal 测试请求。
    /// </summary>
    /// <param name="destination">失败工作簿目标流。</param>
    /// <returns>带失败工作簿配置的导入请求。</returns>
    private static ExcelWorkbookImportRequest<ImportWorkbook> CreateAnnotatedOriginalRequest(
        Stream destination) => ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

    /// <summary>
    /// 修改 XLSX 包中的一个部件并返回新的输入流。
    /// </summary>
    /// <param name="source">原始 XLSX 流。</param>
    /// <param name="selector">部件选择器。</param>
    /// <param name="rewrite">部件字节转换器。</param>
    /// <returns>位于起始位置的修改后 XLSX 流。</returns>
    private static MemoryStream RewriteXlsxEntry(Stream source, Func<string, bool> selector,
        Func<byte[], byte[]> rewrite)
    {
        source.Position = 0;
        using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        using var output = new MemoryStream();
        using (var target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in archive.Entries)
            {
                using var input = entry.Open();
                using var buffer = new MemoryStream();
                input.CopyTo(buffer);
                var bytes = selector(entry.FullName) ? rewrite(buffer.ToArray()) : buffer.ToArray();
                var targetEntry = target.CreateEntry(entry.FullName, CompressionLevel.NoCompression);
                using var targetStream = targetEntry.Open();
                targetStream.Write(bytes, 0, bytes.Length);
            }
        }
        return new MemoryStream(output.ToArray());
    }

    /// <summary>
    /// 为关系部件加入外部目标关系。
    /// </summary>
    /// <param name="bytes">关系部件字节。</param>
    /// <returns>修改后的关系部件字节。</returns>
    private static byte[] AddExternalRelationship(byte[] bytes)
    {
        const string relationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        var document = ParseZipXml(bytes);
        document.Root.Add(new XElement(XNamespace.Get(relationshipNamespace) + "Relationship",
            new XAttribute("Id", "rIdExternal"),
            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"),
            new XAttribute("Target", "https://example.com"),
            new XAttribute("TargetMode", "External")));
        return Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
    }

    /// <summary>
    /// 将测试工作簿中的隐藏定义名称标记为隐藏。
    /// </summary>
    /// <param name="bytes">工作簿部件字节。</param>
    /// <returns>修改后的工作簿部件字节。</returns>
    private static byte[] AddHiddenDefinedName(byte[] bytes)
    {
        var document = ParseZipXml(bytes);
        var name = document.Descendants().First(element =>
            string.Equals(element.Name.LocalName, "definedName", StringComparison.Ordinal)
            && string.Equals((string)element.Attribute("name"), "HiddenFailedQuantity",
                StringComparison.Ordinal));
        name.SetAttributeValue("hidden", "1");
        return Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
    }

    /// <summary>
    /// 将普通工作簿清单改成宏工作簿清单。
    /// </summary>
    /// <param name="bytes">内容类型部件字节。</param>
    /// <returns>修改后的内容类型部件字节。</returns>
    private static byte[] ConvertToMacroWorkbook(byte[] bytes)
    {
        const string contentTypeNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
        var document = ParseZipXml(bytes);
        var workbookOverride = document.Root.Elements(XNamespace.Get(contentTypeNamespace) + "Override")
            .First(element => ((string)element.Attribute("PartName"))?.EndsWith("/workbook.xml",
                StringComparison.OrdinalIgnoreCase) == true);
        workbookOverride.SetAttributeValue("ContentType",
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml");
        return Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
    }

    /// <summary>
    /// 为已有批注 VML 加入一个非批注形状。
    /// </summary>
    /// <param name="bytes">VML 部件字节。</param>
    /// <returns>修改后的 VML 部件字节。</returns>
    private static byte[] AddForeignVmlShape(byte[] bytes)
    {
        const string vmlNamespace = "urn:schemas-microsoft-com:vml";
        var document = ParseZipXml(bytes);
        document.Root.Add(new XElement(XNamespace.Get(vmlNamespace) + "shape",
            new XAttribute("id", "foreign-shape")));
        return Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
    }

    /// <summary>
    /// 读取测试 ZIP 中的 XML 部件。
    /// </summary>
    /// <param name="bytes">XML 部件字节。</param>
    /// <returns>解析后的 XML 文档。</returns>
    private static XDocument ParseZipXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    /// <summary>
    /// 打开导入测试使用的黄金样例文件。
    /// </summary>
    /// <param name="relativePath">相对于黄金样例目录的文件路径。</param>
    /// <returns>已打开的只读样例文件流。</returns>
    private static FileStream OpenFixture(string relativePath)
        => File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Resources", "Golden", relativePath));

    /// <summary>
    /// 承载人员导入结果的测试工作簿。
    /// </summary>
    private sealed class ImportWorkbook
    {
        /// <summary>
        /// 获取主要工作表导入的人员集合。
        /// </summary>
        public List<Person> People { get; } = new List<Person>();
        /// <summary>
        /// 获取第二个工作表导入的人员集合。
        /// </summary>
        public List<Person> Other { get; } = new List<Person>();
    }

    /// <summary>
    /// 承载异构工作表测试结果的 Workbook 模型。
    /// </summary>
    private sealed class HeterogeneousWorkbook
    {
        /// <summary>
        /// 获取人员工作表结果。
        /// </summary>
        public List<Person> People { get; } = new List<Person>();
        /// <summary>
        /// 获取备注工作表结果。
        /// </summary>
        public List<RichPerson> Notes { get; } = new List<RichPerson>();
    }

    /// <summary>
    /// 记录失败工作簿路径提交参数并委托默认原子实现的测试提交器。
    /// </summary>
    private sealed class RecordingCommitter : IFileExportCommitter
    {
        /// <summary>
        /// 默认原子文件提交实现。
        /// </summary>
        private readonly DefaultFileExportCommitter _inner = new DefaultFileExportCommitter();

        /// <summary>
        /// 最后一次提交使用的格式名称。
        /// </summary>
        internal string Format { get; private set; }

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken, string format)
        {
            Format = format;
            _inner.Commit(path, write, cancellationToken, format);
        }

        /// <inheritdoc />
        public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format)
        {
            Format = format;
            return _inner.CommitAsync(path, writeAsync, cancellationToken, format);
        }
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
    /// 承载文本、日期和空值导入结果的测试工作簿。
    /// </summary>
    private sealed class RichImportWorkbook
    {
        /// <summary>
        /// 获取文本、日期和空值导入结果集合。
        /// </summary>
        public List<RichPerson> Items { get; } = new List<RichPerson>();
    }

    /// <summary>
    /// 用于文本、日期和空值映射的人员测试模型。
    /// </summary>
    private sealed class RichPerson
    {
        /// <summary>
        /// 获取或设置测试人员姓名。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置测试人员注册日期。
        /// </summary>
        public DateTime RegisteredAt { get; set; }
        /// <summary>
        /// 获取或设置测试备注。
        /// </summary>
        public string Note { get; set; }
    }

    /// <summary>
    /// 承载日期系统测试结果的工作簿。
    /// </summary>
    private sealed class DateImportWorkbook
    {
        /// <summary>
        /// 获取日期导入结果集合。
        /// </summary>
        public List<DateValue> Items { get; } = new List<DateValue>();
    }

    /// <summary>
    /// 用于日期系统还原的测试数据行。
    /// </summary>
    private sealed class DateValue
    {
        /// <summary>
        /// 获取或设置还原后的日期。
        /// </summary>
        public DateTime Date { get; set; }
    }

    /// <summary>
    /// 承载公式缓存读取结果的测试工作簿。
    /// </summary>
    private sealed class FormulaImportWorkbook
    {
        /// <summary>
        /// 获取公式缓存导入结果集合。
        /// </summary>
        public List<FormulaValue> Items { get; } = new List<FormulaValue>();
    }

    /// <summary>
    /// 用于公式缓存值映射的测试数据行。
    /// </summary>
    private sealed class FormulaValue
    {
        /// <summary>
        /// 获取或设置导入的公式缓存数值。
        /// </summary>
        public int Value { get; set; }
    }

    /// <summary>
    /// 承载真实 XLSB 样例导入结果的测试工作簿。
    /// </summary>
    private sealed class XlsbWorkbook
    {
        /// <summary>
        /// 获取工作表导入的数据行集合。
        /// </summary>
        public List<XlsbRow> Rows { get; } = new List<XlsbRow>();
    }

    /// <summary>
    /// 匹配真实 XLSB 样例表头的测试数据行。
    /// </summary>
    private sealed class XlsbRow
    {
        /// <summary>
        /// 获取或设置 XLSB 样例 A 列的文本值。
        /// </summary>
        public string A { get; set; }
        /// <summary>
        /// 获取或设置 XLSB 样例 B 列的布尔值。
        /// </summary>
        public bool B { get; set; }
        /// <summary>
        /// 获取或设置 XLSB 样例 C 列的日期值。
        /// </summary>
        public DateTime C { get; set; }
        /// <summary>
        /// 获取或设置 XLSB 样例 D 列的文本值。
        /// </summary>
        public string D { get; set; }
        /// <summary>
        /// 获取或设置 XLSB 样例 E 列的文本值。
        /// </summary>
        public string E { get; set; }
        /// <summary>
        /// 获取或设置 XLSB 样例 F 列的文本值。
        /// </summary>
        public string F { get; set; }
    }

    /// <summary>
    /// 首次异步读取到数据后触发取消的测试流。
    /// </summary>
    /// <remarks>不接管底层流的所有权。</remarks>
    private sealed class CancelAfterFirstReadStream : Stream
    {
        /// <summary>
        /// 提供测试输入内容的底层流，其所有权保留给调用方。
        /// </summary>
        private readonly Stream _inner;
        /// <summary>
        /// 在首次读取到数据后触发取消的令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 初始化一个 <see cref="CancelAfterFirstReadStream" /> 类型的实例。
        /// </summary>
        /// <param name="inner">提供测试输入内容的底层流。</param>
        /// <param name="cancellation">首次读取到数据后需要触发的取消令牌源。</param>
        internal CancelAfterFirstReadStream(Stream inner, CancellationTokenSource cancellation)
        {
            _inner = inner;
            _cancellation = cancellation;
        }

        /// <summary>
        /// 获取是否已读取到输入数据。
        /// </summary>
        internal bool HasRead { get; private set; }
        /// <summary>
        /// 获取包装流是否已被释放。
        /// </summary>
        internal bool WasDisposed { get; private set; }
/// <inheritdoc />
        public override bool CanRead => !WasDisposed && _inner.CanRead;
/// <inheritdoc />
        public override bool CanSeek => !WasDisposed && _inner.CanSeek;
/// <inheritdoc />
        public override bool CanWrite => false;
/// <inheritdoc />
        public override long Length => _inner.Length;
/// <inheritdoc />
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
/// <inheritdoc />
        public override void Flush() => _inner.Flush();
/// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
/// <inheritdoc />
        public override void SetLength(long value) => throw new System.NotSupportedException();
/// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => throw new System.NotSupportedException();
/// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) =>
            throw new System.NotSupportedException("ImportAsync must use asynchronous input IO.");
/// <inheritdoc />
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            var read = await _inner.ReadAsync(buffer, offset, count, cancellationToken);
            if (read > 0 && !HasRead)
            {
                HasRead = true;
                _cancellation.Cancel();
            }
            return read;
        }
/// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            WasDisposed = true;
            base.Dispose(disposing);
        }
    }
}
