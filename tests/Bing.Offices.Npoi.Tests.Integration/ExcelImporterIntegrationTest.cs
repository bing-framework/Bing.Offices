using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Csv;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Extensions;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Validations;
using Microsoft.Extensions.DependencyInjection;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests.Integration;

/// <summary>
/// Excel 导入器集成测试。
/// </summary>
public class ExcelImporterIntegrationTest
{
    /// <summary>
    /// 测试 - Windows 目标文件被占用时，原子 Excel 文件导出应保留旧目标内容和临时文件清理。
    /// </summary>
    [Fact]
    public void ExportToFile_TargetLocked_ShouldKeepExistingTarget()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return;

        // Arrange
        var path = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Locked.{Guid.NewGuid():N}.xlsx");
        File.WriteAllText(path, "原始内容", System.Text.Encoding.UTF8);
        using var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var provider = BuildProvider();
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data",
            new[] { new FileIntegrationRow { Name = "新内容", Count = 1 } }));

        try
        {
            // Act
            var exception = Assert.Throws<BingOfficesFileCommitException>(() => provider.GetRequiredService<IExcelExporter>()
                .ExportToFile(request, path));
            Assert.IsType<IOException>(exception.InnerException);

            // Assert
            Assert.Equal("原始内容", File.ReadAllText(path, System.Text.Encoding.UTF8));
            Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(path), Path.GetFileName(path) + ".*.tmp"));
        }
        finally
        {
            locked.Dispose();
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 测试 - Failure Workbook 临时目录路径与文件冲突时，应返回稳定的目录创建错误。
    /// </summary>
    [Fact]
    public void FailureWorkbook_TemporaryDirectoryPathConflict_ShouldClassifyError()
    {
        // Arrange
        var conflictPath = Path.Combine(Path.GetTempPath(), $"Bing.Offices.FailureConflict.{Guid.NewGuid():N}");
        File.WriteAllText(conflictPath, "directory-conflict", System.Text.Encoding.UTF8);
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(ConflictIntegrationRow.Count));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        }));
        using var destination = new MemoryStream(new byte[] { 7 });
        using var provider = BuildProvider();
        var request = ExcelImport.Workbook<IntegrationWorkbook<ConflictIntegrationRow>>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination,
                TemporaryDirectory = conflictPath
            })
            .Sheet("Data", root => root.Rows));

        try
        {
            // Act
            var exception = Assert.Throws<BingOfficesImportException>(() => provider.GetRequiredService<IExcelImporter>()
                .Import(source, request));

            // Assert
            Assert.Equal("失败工作簿临时目录创建失败。", exception.Message);
            Assert.IsType<IOException>(exception.InnerException);
            Assert.Equal(1, destination.Length);
        }
        finally
        {
            if (File.Exists(conflictPath))
                File.Delete(conflictPath);
        }
    }

    /// <summary>
    /// 代码自身无法证明 Windows 文件占用时 Failure Workbook 复制阶段的异常合同。
    /// </summary>
    [Fact]
    public void FailureWorkbook_LockedDestination_ShouldClassifyCopyFailureAndCleanup()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return;

        // Arrange
        var directory = Path.Combine(Path.GetTempPath(), $"Bing.Offices.FailureLocked.{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var destinationPath = Path.Combine(directory, "failure.xlsx");
        var temporaryDirectory = Path.Combine(directory, "temporary");
        Directory.CreateDirectory(temporaryDirectory);
        File.WriteAllText(destinationPath, "原始失败文件", System.Text.Encoding.UTF8);
        using var destination = new LockedDestinationStream(destinationPath);
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(ConflictIntegrationRow.Count));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        }));
        using var provider = BuildProvider();
        var request = ExcelImport.Workbook<IntegrationWorkbook<ConflictIntegrationRow>>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination,
                TemporaryDirectory = temporaryDirectory
            })
            .Sheet("Data", root => root.Rows));

        try
        {
            // Act
            var exception = Assert.Throws<BingOfficesImportException>(() => provider.GetRequiredService<IExcelImporter>()
                .Import(source, request));

            // Assert
            Assert.Equal("失败工作簿复制到目标流失败。", exception.Message);
            Assert.IsType<IOException>(exception.InnerException);
            Assert.Equal("原始失败文件", File.ReadAllText(destinationPath, System.Text.Encoding.UTF8));
            Assert.Empty(Directory.GetFiles(temporaryDirectory));
        }
        finally
        {
            destination.Dispose();
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 测试 - 通过 DI 解析的导入器应在真实 XLSX 流中执行调用方注册的自定义校验规则。
    /// </summary>
    [Fact]
    public void AddBingOfficesNpoi_RegisteredCustomValidationRule_ShouldValidateRealWorkbook()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        services.AddSingleton<IExcelValidationRule, StartsWithExcelValidationRule>();
        using var provider = services.BuildServiceProvider();
        using var source = new MemoryStream(CreateWorkbook());

        // Act
        var result = provider.GetRequiredService<IExcelImporter>().Import(source,
            CreateImportRequest<IntegrationRow>());

        // Assert
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal(nameof(IntegrationRow.Code), error.PropertyName);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 测试 - 通过 NPOI 注册解析的导入器应执行最大值特性校验。
    /// </summary>
    [Fact]
    public void AddBingOfficesNpoi_MaxValueAttribute_ShouldValidateRealWorkbook()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        using var provider = services.BuildServiceProvider();
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Amount");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("11");
        }));

        // Act
        var result = provider.GetRequiredService<IExcelImporter>().Import(source,
            ExcelImport.Workbook<IntegrationWorkbook<MaxValueIntegrationRow>>(
                builder => builder.Sheet("Data", root => root.Rows)));

        // Assert
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.MaxValue, error.Code);
        Assert.Equal(nameof(MaxValueIntegrationRow.Amount), error.PropertyName);
    }

    /// <summary>
    /// 测试 - 通过 DI 注册的双向转换器应参与真实 XLSX 的导出与导入。
    /// </summary>
    [Fact]
    public void AddBingOfficesNpoi_RegisteredCustomValueConverter_ShouldRoundTripRealWorkbook()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        services.AddSingleton<IExcelValueConverter, IntegrationCodeExcelValueConverter>();
        using var provider = services.BuildServiceProvider();
        using var destination = new MemoryStream();

        // Act
        provider.GetRequiredService<IExcelExporter>().Export(ExcelExport.Workbook(builder => builder.AddSheet("Data",
            new[] { new ConvertedIntegrationRow { Code = new IntegrationCode("42") } })), destination);
        destination.Position = 0;
        var result = provider.GetRequiredService<IExcelImporter>().Import(destination,
            CreateImportRequest<ConvertedIntegrationRow>());

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("42", Assert.Single(result.Workbook.Rows).Code.Value);
        Assert.True(destination.CanRead);
        Assert.True(destination.CanWrite);
    }

    /// <summary>
    /// 测试 - NPOI 注册的 CSV 实体服务应通过真实流完成往返且不关闭调用方流。
    /// </summary>
    [Fact]
    public void AddBingOfficesNpoi_CsvServices_ShouldRoundTripStream()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        using var provider = services.BuildServiceProvider();
        using var destination = new MemoryStream();

        // Act
        provider.GetRequiredService<ICsvExporter>().Export(new[]
        {
            new CsvIntegrationRow { Name = "A,\"B\"", Count = 7 }
        }, destination);
        destination.Position = 0;
        var result = provider.GetRequiredService<ICsvImporter>().Import<CsvIntegrationRow>(destination);

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("A,\"B\"", Assert.Single(result.Items).Name);
        Assert.Equal(7, result.Items[0].Count);
        Assert.True(destination.CanRead);
    }

    /// <summary>
    /// 测试 - 通过 DI 注册的命名校验规则应由 JSON 映射配置在真实 XLSX 导入中解析。
    /// </summary>
    [Fact]
    public void AddBingOfficesNpoi_RegisteredNamedValidationRule_ShouldValidateConfiguredWorkbook()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        services.AddSingleton<INamedExcelValidationRule, NamedStartsWithValidationRule>();
        using var provider = services.BuildServiceProvider();
        using var source = new MemoryStream(CreateWorkbook());
        var document = ExcelMappingConfigurationLoader.FromJsonDocument(
            "{\"version\":2,\"import\":{\"columns\":[{\"propertyName\":\"Code\",\"validationRuleNames\":[\"starts-with-ok\"]}]}}");
        var configuration = document.Import;

        // Act
        var result = provider.GetRequiredService<IExcelImporter>().Import(source,
            CreateImportRequest<ConfiguredIntegrationRow>(sheet => sheet.Mapping(configuration)));

        // Assert
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal(nameof(ConfiguredIntegrationRow.Code), error.PropertyName);
    }

    /// <summary>
    /// 测试 - 真实 XLS 与 XLSX 文件均应通过路径适配完成实体往返。
    /// </summary>
    /// <param name="format">要测试的 Excel 文件格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void StreamExtensions_RealExcelFormats_ShouldRoundTripThroughPath(ExcelFormat format)
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        using var provider = services.BuildServiceProvider();
        var extension = format == ExcelFormat.Xls ? "xls" : "xlsx";
        var path = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Integration.{Guid.NewGuid():N}.{extension}");

        try
        {
            // Act
            var exporter = provider.GetRequiredService<IExcelExporter>();
            var importer = provider.GetRequiredService<IExcelImporter>();
            using (var destination = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                exporter.Export(ExcelExport.Workbook(builder => builder.Format(format).AddSheet("Data",
                    new[] { new FileIntegrationRow { Name = "路径", Count = 7 } })), destination);
            using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var result = importer.Import(source, CreateImportRequest<FileIntegrationRow>());

            // Assert
            Assert.Empty(result.Errors);
            var item = Assert.Single(result.Workbook.Rows);
            Assert.Equal("路径", item.Name);
            Assert.Equal(7, item.Count);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 测试 - 真实 XLSX 文件应保留跨列表头、批注和图表绘制结果。
    /// </summary>
    [Fact]
    public void ExportToFile_ShouldPersistHeaderCommentMergedCellsAndCharts()
    {
        var path = Path.Combine(Path.GetTempPath(), $"Bing.Offices.RichWorkbook.{Guid.NewGuid():N}.xlsx");
        using var provider = BuildProvider();
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("统计",
            new[]
            {
                new ChartIntegrationRow { Month = "一月", Amount = 10, Forecast = 12 },
                new ChartIntegrationRow { Month = "二月", Amount = 15, Forecast = 16 }
            }, sheet => sheet
                .HeaderRowIndex(1)
                .DataRowStartIndex(2)
                .HeaderRows(new[] { new ExcelHeaderRow(0, new[]
                {
                    new ExcelHeaderCell(0, "统计摘要", columnSpan: 2,
                        comment: new ExcelComment("真实文件批注", "integration", true))
                }) })
                .Chart(new ExcelChartDefinition
                {
                    Title = "金额柱状图",
                    Type = ExcelChartType.Column,
                    Categories = new ExcelChartRange { ColumnKey = "Month" },
                    Series = new[] { new ExcelChartSeries
                    {
                        Name = "金额", Values = new ExcelChartRange { ColumnKey = "Amount" }
                    } },
                    Anchor = new ExcelChartAnchor { StartRow = 0, StartColumn = 5, EndRow = 12, EndColumn = 14 }
                })));

        try
        {
            provider.GetRequiredService<IExcelExporter>().ExportToFile(request, path);
            using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var workbook = (XSSFWorkbook)WorkbookFactory.Create(source);
            var sheet = workbook.GetSheet("统计");

            Assert.Contains(sheet.MergedRegions,
                region => region.FirstRow == 0 && region.LastRow == 0 &&
                          region.FirstColumn == 0 && region.LastColumn == 1);
            var comment = sheet.GetRow(0).GetCell(0).CellComment;
            Assert.NotNull(comment);
            Assert.Equal("真实文件批注", comment.String.String);
            Assert.Equal("integration", comment.Author);
            var drawing = Assert.IsType<XSSFDrawing>(sheet.CreateDrawingPatriarch());
            Assert.Contains(drawing.GetCharts(), chart => chart.Title.String == "金额柱状图");
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 测试 - 真实文件导出应保留模板其它 Sheet，并把命名区域作为写入起点。
    /// </summary>
    [Fact]
    public void ExportToFile_ShouldPersistTemplateRegionAndOtherSheets()
    {
        var path = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Template.{Guid.NewGuid():N}.xlsx");
        byte[] templateBytes;
        using (var workbook = new XSSFWorkbook())
        {
            var sheet = workbook.CreateSheet("订单");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("保留内容");
            var name = workbook.CreateName();
            name.NameName = "OrderTable";
            name.RefersToFormula = "'订单'!$B$3:$C$20";
            workbook.CreateSheet("说明").CreateRow(0).CreateCell(0).SetCellValue("说明");
            using var templateOutput = new MemoryStream();
            workbook.Write(templateOutput, false);
            templateBytes = templateOutput.ToArray();
        }
        using var template = new MemoryStream(templateBytes);
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("订单", new[] { new FileIntegrationRow { Name = "客户 A", Count = 1 } },
                sheet => sheet.UseTemplateRegion("OrderTable")));
        using var provider = BuildProvider();

        try
        {
            provider.GetRequiredService<IExcelExporter>().ExportToFile(request, path);
            using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var result = (XSSFWorkbook)WorkbookFactory.Create(source);
            Assert.Equal("保留内容", result.GetSheet("订单").GetRow(0).GetCell(0).StringCellValue);
            Assert.Equal("客户 A", result.GetSheet("订单").GetRow(3).GetCell(1).StringCellValue);
            Assert.Equal("说明", result.GetSheet("说明").GetRow(0).GetCell(0).StringCellValue);
            Assert.True(template.CanRead);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 测试 - 真实 XLSX 文件中的图片锚点应绑定到 byte[] 属性。
    /// </summary>
    [Fact]
    public void ImportFromFile_ShouldBindImageAnchoredCell()
    {
        var path = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Image.{Guid.NewGuid():N}.xlsx");
        var imageBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x00
        };
        using (var workbook = new XSSFWorkbook())
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Photo");
            sheet.CreateRow(1).CreateCell(0).SetCellType(CellType.Blank);
            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Row1 = 1;
            anchor.Row2 = 2;
            anchor.Col1 = 0;
            anchor.Col2 = 1;
            sheet.CreateDrawingPatriarch().CreatePicture(anchor,
                workbook.AddPicture(imageBytes, PictureType.PNG));
            using var output = File.Create(path);
            workbook.Write(output, false);
        }
        using var provider = BuildProvider();

        try
        {
            using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var result = provider.GetRequiredService<IExcelImporter>().Import(source,
                ExcelImport.Workbook<ImageIntegrationWorkbook>(builder =>
                    builder.Sheet("Data", root => root.Rows)));
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(imageBytes, Assert.Single(result.Workbook.Rows).Photo);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 测试 - 显式 Large 门禁验证 NPOI 500K/1M 行真实文件往返；普通 PR 通过 Category 过滤跳过。
    /// </summary>
    /// <param name="rowCount">要生成和导入的数据行数。</param>
    [Trait("Category", "Large")]
    [Theory]
    [InlineData(500_000)]
    [InlineData(1_000_000)]
    public void LargeFileRoundTrip_ShouldUseRealPath(int rowCount)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Npoi.Large.{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "large.xlsx");
        using var provider = BuildProvider();
        var rows = Enumerable.Range(0, rowCount).Select(index => new FileIntegrationRow
        {
            Name = $"large-{index}",
            Count = index
        });
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data", rows));
        var importRequest = ExcelImport.Workbook<IntegrationWorkbook<FileIntegrationRow>>(builder =>
            builder.ResourceLimits(new ExcelResourceLimits
            {
                MaxInputBytes = null,
                MaxZipEntryUncompressedBytes = null,
                MaxZipTotalUncompressedBytes = null,
                MaxXmlCharacters = null,
                MaxSharedStringsBytes = null,
                MaxWorksheetBytes = null,
                MaxTotalWorksheetBytes = null
            }).Sheet("Data", root => root.Rows));

        try
        {
            provider.GetRequiredService<IExcelExporter>().ExportToFile(request, path);
            var result = ExcelStreamExtensions.ImportFromFile<IntegrationWorkbook<FileIntegrationRow>>(
                provider.GetRequiredService<IExcelImporter>(), path, importRequest);

            Assert.Empty(result.Errors);
            Assert.Equal(rowCount, result.Workbook.Rows.Count);
            Assert.Equal("large-0", result.Workbook.Rows[0].Name);
            Assert.Equal(0, result.Workbook.Rows[0].Count);
            Assert.Equal($"large-{rowCount - 1}", result.Workbook.Rows[rowCount - 1].Name);
            Assert.Equal(rowCount - 1, result.Workbook.Rows[rowCount - 1].Count);
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 测试 - 已取消的真实导出请求不应向调用方目标流写入内容。
    /// </summary>
    [Fact]
    public void AddBingOfficesNpoi_PreCancelledExport_ShouldNotWriteDestination()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        using var provider = services.BuildServiceProvider();
        using var destination = new MemoryStream();
        using var cancellationTokenSource = new System.Threading.CancellationTokenSource();
        cancellationTokenSource.Cancel();

        // Act
        var action = () => provider.GetRequiredService<IExcelExporter>().Export(
            ExcelExport.Workbook(builder => builder.AddSheet("Data",
                new[] { new FileIntegrationRow { Name = "取消", Count = 1 } })), destination,
            cancellationTokenSource.Token);

        // Assert
        Assert.Throws<OperationCanceledException>(action);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 测试 - 导出写入过程中取消时应停止后续写入并保持调用方目标流打开。
    /// </summary>
    [Fact]
    public void AddBingOfficesNpoi_MidWriteCancelledExport_ShouldStopWritingAndPreserveDestination()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        using var provider = services.BuildServiceProvider();
        using var cancellationTokenSource = new System.Threading.CancellationTokenSource();
        using var destination = new CancelAfterFirstWriteStream(cancellationTokenSource);
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data",
            Enumerable.Range(0, 20).Select(index => new FileIntegrationRow
            {
                Name = $"取消-{index}",
                Count = index
            })));

        // Act
        var action = () => provider.GetRequiredService<IExcelExporter>().Export(request, destination,
            cancellationTokenSource.Token);

        // Assert
        Assert.Throws<OperationCanceledException>(action);
        Assert.Equal(1, destination.WriteCount);
        Assert.True(destination.CanWrite);
    }

    /// <summary>
    /// 测试 - 通过 DI 解析的导入器应支持不可寻址 XLSX 流、保持调用方流打开，并在预取消时不读取输入。
    /// </summary>
    [Fact]
    public void AddBingOfficesNpoi_NonSeekableAndPreCancelledImport_ShouldPreserveStreamContract()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        using var provider = services.BuildServiceProvider();
        using var generatedWorkbook = new MemoryStream();
        provider.GetRequiredService<IExcelExporter>().Export(ExcelExport.Workbook(builder => builder.AddSheet("Data",
            new[] { new FileIntegrationRow { Name = "不可寻址", Count = 7 } })), generatedWorkbook);
        using var source = new NonSeekableReadStream(generatedWorkbook.ToArray());
        using var cancelledSource = new NonSeekableReadStream(generatedWorkbook.ToArray());
        using var cancellationTokenSource = new System.Threading.CancellationTokenSource();
        cancellationTokenSource.Cancel();

        // Act
        var result = provider.GetRequiredService<IExcelImporter>().Import(source,
            CreateImportRequest<FileIntegrationRow>());
        var cancelled = () => provider.GetRequiredService<IExcelImporter>().Import(cancelledSource,
            CreateImportRequest<FileIntegrationRow>(), cancellationTokenSource.Token);

        // Assert
        var item = Assert.Single(result.Workbook.Rows);
        Assert.Empty(result.Errors);
        Assert.Equal("不可寻址", item.Name);
        Assert.Equal(7, item.Count);
        Assert.True(source.CanRead);
        Assert.Throws<OperationCanceledException>(cancelled);
        Assert.True(cancelledSource.CanRead);
    }

    /// <summary>
    /// 测试 - Fluent 配置与 XML 映射配置应在真实 XLSX 导入导出中使用相同的列定义。
    /// </summary>
    [Fact]
    public void AddBingOfficesNpoi_FluentConfigurationAndXmlConfiguration_ShouldRoundTripRealWorkbook()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        using var provider = services.BuildServiceProvider();
        using var destination = new MemoryStream();
        var profile = new ExportMappingBuilder<MappingIntegrationRow>()
            .Property(row => row.Code).HasHeader("业务编码").And()
            .Build();
        var xmlConfiguration = ExcelMappingConfigurationLoader.FromXmlDocument(
            "<ExcelMappingDocument><Version>2</Version><Import><Columns><ExcelColumnConfiguration><PropertyName>Code</PropertyName><Title>业务编码</Title></ExcelColumnConfiguration></Columns></Import></ExcelMappingDocument>");

        // Act
        provider.GetRequiredService<IExcelExporter>().Export(ExcelExport.Workbook(builder => builder.AddSheet("Data",
            new[] { new MappingIntegrationRow { Code = "配置" } }, sheet => sheet.Mapping(profile))), destination);
        destination.Position = 0;
        var result = provider.GetRequiredService<IExcelImporter>().Import(destination,
            CreateImportRequest<MappingIntegrationRow>(sheet => sheet.Mapping(xmlConfiguration)));

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("配置", Assert.Single(result.Workbook.Rows).Code);
        Assert.True(destination.CanRead);
    }

    /// <summary>
    /// 测试 - NPOI 注册应解析安全配置加载器，并将已注册的双向转换器注入 CSV 实体服务。
    /// </summary>
    [Fact]
    public void AddBingOfficesNpoi_ConfigurationLoaderAndCsvConverter_ShouldResolveThroughDi()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        services.AddSingleton<IExcelValueConverter, IntegrationCodeExcelValueConverter>();
        using var provider = services.BuildServiceProvider();
        using var destination = new MemoryStream();

        // Act
        var configuration = provider.GetRequiredService<IExcelMappingConfigurationLoader>().FromJsonDocument(
            "{\"version\":2,\"import\":{\"columns\":[{\"propertyName\":\"Code\",\"converterName\":\"integration-code\"}]},\"export\":{\"columns\":[]}}");
        provider.GetRequiredService<ICsvExporter>().Export(new[]
        {
            new ConvertedIntegrationRow { Code = new IntegrationCode("42") }
        }, destination, new CsvExportOptions<ConvertedIntegrationRow> { MappingDocument = configuration });
        destination.Position = 0;
        var result = provider.GetRequiredService<ICsvImporter>().Import<ConvertedIntegrationRow>(destination,
            new CsvImportOptions<ConvertedIntegrationRow> { MappingDocument = configuration });

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("42", Assert.Single(result.Items).Code.Value);
        Assert.True(destination.CanRead);
    }

    /// <summary>
    /// 创建集成测试工作簿字节。
    /// </summary>
    /// <typeparam name="T">工作簿数据行类型。</typeparam>
    /// <param name="configure">可选的工作表导入配置回调。</param>
    /// <returns>配置为导入 <typeparamref name="T" /> 行的工作簿请求。</returns>
    private static ExcelWorkbookImportRequest<IntegrationWorkbook<T>> CreateImportRequest<T>(
        Action<ExcelSheetImportBuilder<T>> configure = null) where T : class, new()
    {
        return ExcelImport.Workbook<IntegrationWorkbook<T>>(builder =>
            builder.Sheet("Data", root => root.Rows, configure));
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    /// <typeparam name="T">工作簿行数据类型。</typeparam>
    private sealed class IntegrationWorkbook<T> where T : class, new()
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<T> Rows { get; } = new();
    }

    /// <summary>
    /// 创建测试工作簿。
    /// </summary>
    /// <param name="configure">配置回调。</param>
    /// <returns>生成的字节内容。</returns>
    private static byte[] CreateWorkbook(Action<XSSFWorkbook> configure = null)
    {
        using var workbook = new XSSFWorkbook();
        if (configure == null)
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(IntegrationRow.Code));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        }
        else
            configure(workbook);
        using var destination = new MemoryStream();
        workbook.Write(destination, false);
        return destination.ToArray();
    }

    /// <summary>
    /// 构建测试使用的服务提供程序。
    /// </summary>
    /// <returns>生成的 ServiceProvider 结果。</returns>
    private static ServiceProvider BuildProvider()
    {
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        return services.BuildServiceProvider();
    }

    /// <summary>
    /// 集成测试行模型。
    /// </summary>
    private class IntegrationRow
    {
        /// <summary>
        /// 获取或设置业务编码。
        /// </summary>
        [StartsWithAttribute]
        public string Code { get; set; }
    }

    /// <summary>
    /// 包含领域编码的集成测试行模型。
    /// </summary>
    private class ConvertedIntegrationRow
    {
        /// <summary>
        /// 获取或设置业务编码。
        /// </summary>
        public IntegrationCode Code { get; set; }
    }

    /// <summary>
    /// 包含最大值校验的集成测试行模型。
    /// </summary>
    private class MaxValueIntegrationRow
    {
        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        [ExcelMaxValue(10)]
        public string Amount { get; set; }
    }

    /// <summary>
    /// CSV 集成测试行模型。
    /// </summary>
    private class CsvIntegrationRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Count { get; set; }
    }

    /// <summary>
    /// 命名规则集成测试行模型。
    /// </summary>
    private class ConfiguredIntegrationRow
    {
        /// <summary>
        /// 获取或设置业务编码。
        /// </summary>
        public string Code { get; set; }
    }

    /// <summary>
    /// 文件路径集成测试行模型。
    /// </summary>
    private class FileIntegrationRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Count { get; set; }
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private class ChartIntegrationRow
    {
        /// <summary>
        /// 获取或设置月份。
        /// </summary>
        public string Month { get; set; }
        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public int Amount { get; set; }
        /// <summary>
        /// 获取或设置预测值。
        /// </summary>
        public int Forecast { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class ImageIntegrationWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<ImageIntegrationRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class ImageIntegrationRow
    {
        /// <summary>
        /// 获取或设置图片。
        /// </summary>
        public byte[] Photo { get; set; }
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private class ConflictIntegrationRow
    {
        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Count { get; set; }
    }

    /// <summary>
    /// 提供测试场景使用的流替身。
    /// </summary>
    private sealed class LockedDestinationStream : Stream
    {
        /// <summary>
        /// 被锁定的目标文件路径。
        /// </summary>
        private readonly string _path;

        /// <summary>
        /// 持有目标文件读取锁的文件流。
        /// </summary>
        private readonly FileStream _lock;

        /// <summary>
        /// 初始化一个 <see cref="LockedDestinationStream" /> 类型的实例。
        /// </summary>
        /// <param name="path">要打开并锁定的目标文件路径。</param>
        public LockedDestinationStream(string path)
        {
            _path = path;
            _lock = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        }

        /// <inheritdoc />
        public override bool CanRead => true;
        /// <inheritdoc />
        public override bool CanSeek => true;
        /// <inheritdoc />
        public override bool CanWrite => true;
        /// <inheritdoc />
        public override long Length => _lock.Length;
        /// <inheritdoc />
        public override long Position
        {
            get => _lock.Position;
            set => _lock.Position = value;
        }

        /// <inheritdoc />
        public override void Flush() { }

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => _lock.Read(buffer, offset, count);

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _lock.Seek(offset, origin);

        /// <inheritdoc />
        public override void SetLength(long value) => _lock.SetLength(value);

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            using var blocked = new FileStream(_path, FileMode.Open, FileAccess.Write, FileShare.None);
            blocked.Write(buffer, offset, count);
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _lock.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 映射配置集成测试行模型。
    /// </summary>
    private class MappingIntegrationRow
    {
        /// <summary>
        /// 获取或设置业务编码。
        /// </summary>
        public string Code { get; set; }
    }

    /// <summary>
    /// 集成测试领域编码。
    /// </summary>
    private sealed class IntegrationCode
    {
        /// <summary>
        /// 初始化一个<see cref="IntegrationCode"/>类型的实例。
        /// </summary>
        /// <param name="value">编码文本。</param>
        public IntegrationCode(string value) => Value = value;

        /// <summary>
        /// 获取编码文本。
        /// </summary>
        public string Value { get; }
    }

    /// <summary>
    /// 前缀校验特性。
    /// </summary>
    [BindFilter(typeof(StartsWithExcelValidationRule))]
    private sealed class StartsWithAttribute : FilterAttributeBase
    {
        /// <inheritdoc />
        public override string ErrorMsg { get; set; } = "必须以 OK- 开头";
    }

    /// <summary>
    /// 前缀校验规则。
    /// </summary>
    private sealed class StartsWithExcelValidationRule : IExcelValidationRule
    {
        /// <inheritdoc />
        public bool CanValidate(FilterAttributeBase attribute) => attribute is StartsWithAttribute;

        /// <inheritdoc />
        public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context) =>
            context.Value.StartsWith("OK-", StringComparison.Ordinal);
    }

    /// <summary>
    /// 集成测试双向编码转换器。
    /// </summary>
    private sealed class IntegrationCodeExcelValueConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "integration-code";

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(IntegrationCode);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = new IntegrationCode(((string)context.Value).Substring(3));
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = $"CV-{((IntegrationCode)context.Value).Value}";
            return true;
        }
    }

    /// <summary>
    /// 通过配置名称解析的前缀校验规则。
    /// </summary>
    private sealed class NamedStartsWithValidationRule : INamedExcelValidationRule
    {
        /// <inheritdoc />
        public string Name => "starts-with-ok";

        /// <inheritdoc />
        public string ErrorMessage => "必须以 OK- 开头";

        /// <inheritdoc />
        public bool Validate(ExcelValidationContext context) =>
            context.Value.StartsWith("OK-", StringComparison.Ordinal);
    }

    /// <summary>
    /// 不支持定位的只读内存流。
    /// </summary>
    private sealed class NonSeekableReadStream : Stream
    {
        /// <summary>
        /// 被代理的内存流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 初始化一个<see cref="NonSeekableReadStream"/>类型的实例。
        /// </summary>
        /// <param name="content">输入内容。</param>
        public NonSeekableReadStream(byte[] content) => _inner = new MemoryStream(content);

        /// <inheritdoc />
        public override bool CanRead => true;

        /// <inheritdoc />
        public override bool CanSeek => false;

        /// <inheritdoc />
        public override bool CanWrite => false;

        /// <inheritdoc />
        public override long Length => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        /// <inheritdoc />
        public override void Flush()
        {
        }

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 在首次底层写入完成后取消令牌，用于确定性验证写入边界取消语义。
    /// </summary>
    private sealed class CancelAfterFirstWriteStream : Stream
    {
        /// <summary>
        /// 承载写入测试数据的内存流。
        /// </summary>
        private readonly MemoryStream _inner = new();

        /// <summary>
        /// 用于在首次写入后取消操作的令牌源。
        /// </summary>
        private readonly System.Threading.CancellationTokenSource _cancellationTokenSource;

        /// <summary>
        /// 初始化一个<see cref="CancelAfterFirstWriteStream"/>类型的实例。
        /// </summary>
        /// <param name="cancellationTokenSource">用于在首次写入后取消的令牌源。</param>
        public CancelAfterFirstWriteStream(System.Threading.CancellationTokenSource cancellationTokenSource) =>
            _cancellationTokenSource = cancellationTokenSource;

        /// <summary>
        /// 获取或设置底层写入次数。
        /// </summary>
        public int WriteCount { get; private set; }

        /// <inheritdoc />
        public override bool CanRead => true;

        /// <inheritdoc />
        public override bool CanSeek => true;

        /// <inheritdoc />
        public override bool CanWrite => true;

        /// <inheritdoc />
        public override long Length => _inner.Length;

        /// <inheritdoc />
        public override long Position { get => _inner.Position; set => _inner.Position = value; }

        /// <inheritdoc />
        public override void Flush() => _inner.Flush();

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);

        /// <inheritdoc />
        public override void SetLength(long value) => _inner.SetLength(value);

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            WriteCount++;
            _inner.Write(buffer, offset, count);
            if (WriteCount == 1)
                _cancellationTokenSource.Cancel();
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
