using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Npoi.Imports;
using Bing.Offices.Npoi;
using Bing.Offices.Providers;
using Bing.Offices.Styles;
using Bing.Offices.Validations;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// Workbook Request、动态列和关系绑定测试。
/// </summary>
public sealed class ExcelWorkbookRequestTest
{
    private static readonly IReadOnlyList<ExcelDynamicColumnDefinition> TenantColumns =
        new[]
        {
            new ExcelDynamicColumnDefinition
            {
                Key = "region",
                Title = "地区",
                Aliases = new[] { "旧地区" },
                DataType = typeof(string),
                Order = 10
            },
            new ExcelDynamicColumnDefinition
            {
                Key = "amount",
                Title = "金额",
                DataType = typeof(decimal),
                NumberFormat = "0.00",
                Order = 20
            }
        };

    /// <summary>
    /// 测试 - Excel 导出应直接写入调用方目标流，保持目标流打开并生成可重开的工作簿。
    /// </summary>
    [Fact]
    public void Export_DestinationStream_ShouldRemainOpenAfterDirectWrite()
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook => workbook
            .Metadata(new ExcelWorkbookMetadataOptions { Author = "作者 A", Title = "标题 A" })
            .AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } }));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        Assert.True(destination.CanWrite);
        Assert.True(destination.Length > 0);
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        Assert.Equal("客户 A", workbook.GetSheet("客户").GetRow(1).GetCell(0).StringCellValue);
        var xssfWorkbook = Assert.IsType<XSSFWorkbook>(workbook);
        Assert.Equal("作者 A", xssfWorkbook.GetProperties().CoreProperties.Creator);
        Assert.Equal("标题 A", xssfWorkbook.GetProperties().CoreProperties.Title);
    }

    /// <summary>
    /// 测试 - Workbook 请求应复制 metadata，避免调用方后续修改配置影响已构建请求。
    /// </summary>
    [Fact]
    public void Export_WorkbookRequest_ShouldSnapshotMetadata()
    {
        // Arrange
        var metadata = new ExcelWorkbookMetadataOptions
        {
            Author = "作者 A",
            Company = "公司 A",
            Title = "标题 A"
        };
        var request = ExcelExport.Workbook(workbook => workbook
            .Metadata(metadata)
            .AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } }));

        // Act
        metadata = new ExcelWorkbookMetadataOptions { Author = "作者 B" };

        // Assert
        Assert.Equal("作者 A", request.Metadata.Author);
        Assert.Equal("公司 A", request.Metadata.Company);
        Assert.Equal("标题 A", request.Metadata.Title);
    }

    /// <summary>
    /// 测试 - 模板默认保留 metadata，显式 Metadata 应在 XLS/XLSX 中覆盖六个字段。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xlsx)]
    [InlineData(ExcelFormat.Xls)]
    public void Export_TemplateMetadata_ShouldPreserveByDefaultAndOverrideExplicitly(ExcelFormat format)
    {
        // Arrange
        var templateBytes = CreateMetadataTemplate(format, "模板作者", "模板公司", "模板标题", "模板主题",
            "模板类别", "模板备注");
        using var template = new MemoryStream(templateBytes);
        var request = ExcelExport.Workbook(workbook => workbook
            .Format(format)
            .UseTemplate(template, leaveOpen: true)
            .Metadata(new ExcelWorkbookMetadataOptions
            {
                Author = "请求作者",
                Company = "请求公司",
                Title = "请求标题",
                Subject = "请求主题",
                Category = "请求类别",
                Description = "请求备注"
            })
            .AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } }));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        if (format == ExcelFormat.Xlsx)
        {
            var properties = Assert.IsType<XSSFWorkbook>(result).GetProperties();
            Assert.Equal("请求作者", properties.CoreProperties.Creator);
            Assert.Equal("请求公司", properties.ExtendedProperties.GetUnderlyingProperties().Company);
            Assert.Equal("请求标题", properties.CoreProperties.Title);
            Assert.Equal("请求主题", properties.CoreProperties.Subject);
            Assert.Equal("请求类别", properties.CoreProperties.Category);
            Assert.Equal("请求备注", properties.CoreProperties.Description);
        }
        else
        {
            var workbook = Assert.IsType<HSSFWorkbook>(result);
            Assert.Equal("请求作者", workbook.SummaryInformation.Author);
            Assert.Equal("请求公司", workbook.DocumentSummaryInformation.Company);
            Assert.Equal("请求标题", workbook.SummaryInformation.Title);
            Assert.Equal("请求主题", workbook.SummaryInformation.Subject);
            Assert.Equal("请求类别", workbook.DocumentSummaryInformation.Category);
            Assert.Equal("请求备注", workbook.SummaryInformation.Comments);
        }
    }

    /// <summary>
    /// 测试 - 未显式设置请求 metadata 时，模板原有 metadata 应保持不变。
    /// </summary>
    [Fact]
    public void Export_TemplateMetadata_WhenNotSpecified_ShouldPreserveTemplateValues()
    {
        // Arrange
        using var template = new MemoryStream(CreateMetadataTemplate(ExcelFormat.Xlsx, "模板作者", "模板公司",
            "模板标题", "模板主题", "模板类别", "模板备注"));
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } }));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = Assert.IsType<XSSFWorkbook>(WorkbookFactory.Create(destination));
        var properties = result.GetProperties();
        Assert.Equal("模板作者", properties.CoreProperties.Creator);
        Assert.Equal("模板公司", properties.ExtendedProperties.GetUnderlyingProperties().Company);
        Assert.Equal("模板标题", properties.CoreProperties.Title);
        Assert.Equal("模板主题", properties.CoreProperties.Subject);
        Assert.Equal("模板类别", properties.CoreProperties.Category);
        Assert.Equal("模板备注", properties.CoreProperties.Description);
    }

    /// <summary>
    /// 测试 - XLS 模板未显式指定请求 metadata 时，应保留 SummaryInformation 和 DocumentSummaryInformation。
    /// </summary>
    [Fact]
    public void Export_XlsTemplateMetadata_WhenNotSpecified_ShouldPreserveTemplateValues()
    {
        // Arrange
        using var template = new MemoryStream(CreateMetadataTemplate(ExcelFormat.Xls, "模板作者", "模板公司",
            "模板标题", "模板主题", "模板类别", "模板备注"));
        var request = ExcelExport.Workbook(workbook => workbook
            .Format(ExcelFormat.Xls)
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } }));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = Assert.IsType<HSSFWorkbook>(WorkbookFactory.Create(destination));
        Assert.Equal("模板作者", result.SummaryInformation.Author);
        Assert.Equal("模板公司", result.DocumentSummaryInformation.Company);
        Assert.Equal("模板标题", result.SummaryInformation.Title);
        Assert.Equal("模板主题", result.SummaryInformation.Subject);
        Assert.Equal("模板类别", result.DocumentSummaryInformation.Category);
        Assert.Equal("模板备注", result.SummaryInformation.Comments);
    }

    /// <summary>
    /// 测试 - 并发导出不同请求 metadata 时，各个 XLS/XLSX 结果应保持请求级隔离。
    /// </summary>
    [Fact]
    public void Export_MetadataConcurrentRequests_ShouldRemainIsolated()
    {
        // Arrange
        var results = new MetadataSnapshot[64];

        // Act
        Parallel.For(0, results.Length, index =>
        {
            var value = $"并发-{index}";
            var request = ExcelExport.Workbook(workbook => workbook
                .Format(index % 2 == 0 ? ExcelFormat.Xlsx : ExcelFormat.Xls)
                .Metadata(new ExcelWorkbookMetadataOptions
                {
                    Author = value,
                    Company = value,
                    Title = value,
                    Subject = value,
                    Category = value,
                    Description = value
                })
                .AddSheet("客户", new[] { new ExportCustomer { Name = value } }));
            using var destination = new MemoryStream();
            new NpoiExcelExporter().Export(request, destination);
            destination.Position = 0;
            using var workbook = WorkbookFactory.Create(destination);
            results[index] = workbook is XSSFWorkbook xssf
                ? new MetadataSnapshot(value,
                    xssf.GetProperties().CoreProperties.Creator,
                    xssf.GetProperties().ExtendedProperties.GetUnderlyingProperties().Company,
                    xssf.GetProperties().CoreProperties.Title,
                    xssf.GetProperties().CoreProperties.Subject,
                    xssf.GetProperties().CoreProperties.Category,
                    xssf.GetProperties().CoreProperties.Description)
                : CreateMetadataSnapshot((HSSFWorkbook)workbook, value);
        });

        // Assert
        foreach (var result in results)
        {
            Assert.Equal(result.Expected, result.Author);
            Assert.Equal(result.Expected, result.Company);
            Assert.Equal(result.Expected, result.Title);
            Assert.Equal(result.Expected, result.Subject);
            Assert.Equal(result.Expected, result.Category);
            Assert.Equal(result.Expected, result.Description);
        }
    }

    /// <summary>
    /// 测试 - Excel 文件导出失败时应保留已存在目标文件的原始内容。
    /// </summary>
    [Fact]
    public void ExportToFile_Failure_ShouldKeepExistingTarget()
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } }));
        var filePath = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Tests.{Guid.NewGuid():N}.xlsx");
        File.WriteAllText(filePath, "原始内容", Encoding.UTF8);

        try
        {
            // Act
            Assert.Throws<InvalidOperationException>(() => new ThrowingExcelExporter()
                .ExportToFile(request, filePath));

            // Assert
            Assert.Equal("原始内容", File.ReadAllText(filePath, Encoding.UTF8));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    /// <summary>
    /// 测试 - Excel 文件导出在预取消和 staging 写后取消时均不得提交目标文件。
    /// </summary>
    [Fact]
    public void ExportToFile_Cancellation_ShouldKeepExistingTarget()
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } }));
        var filePath = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Cancel.{Guid.NewGuid():N}.xlsx");
        File.WriteAllText(filePath, "原始内容", Encoding.UTF8);
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();

        try
        {
            // Act
            Assert.Throws<OperationCanceledException>(() => new NpoiExcelExporter()
                .ExportToFile(request, filePath, canceled.Token));
            Assert.Equal("原始内容", File.ReadAllText(filePath, Encoding.UTF8));

            using var writeCanceled = new CancellationTokenSource();
            Assert.Throws<OperationCanceledException>(() => new CancelingExcelExporter()
                .ExportToFile(request, filePath, writeCanceled.Token));

            // Assert
            Assert.Equal("原始内容", File.ReadAllText(filePath, Encoding.UTF8));
            Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(filePath),
                Path.GetFileName(filePath) + ".*.tmp"));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    /// <summary>
    /// 测试 - Excel 文件导出成功时应支持新建和替换目标，并可重新打开结果。
    /// </summary>
    [Fact]
    public void ExportToFile_Success_ShouldCreateAndReplaceTarget()
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("客户", new[] { new ExportCustomer { Name = "新内容" } }));
        var filePath = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Success.{Guid.NewGuid():N}.xlsx");

        try
        {
            // Act
            new NpoiExcelExporter().ExportToFile(request, filePath);

            // Assert
            using (var first = File.OpenRead(filePath))
            using (var workbook = WorkbookFactory.Create(first))
                Assert.Equal("新内容", workbook.GetSheet("客户").GetRow(1).GetCell(0).StringCellValue);

            var replacementRequest = ExcelExport.Workbook(workbook => workbook
                .AddSheet("客户", new[] { new ExportCustomer { Name = "替换内容" } }));
            new NpoiExcelExporter().ExportToFile(replacementRequest, filePath);
            using var second = File.OpenRead(filePath);
            using var replaced = WorkbookFactory.Create(second);
            Assert.Equal("替换内容", replaced.GetSheet("客户").GetRow(1).GetCell(0).StringCellValue);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    /// <summary>
    /// 测试 - Excel 文件提交目标不可替换时应保留 staging 清理和目标目录。
    /// </summary>
    [Fact]
    public void ExportToFile_CommitFailure_ShouldNotReplaceTargetDirectory()
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } }));
        var directoryPath = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Directory.{Guid.NewGuid():N}.xlsx");
        Directory.CreateDirectory(directoryPath);

        try
        {
            // Act
            Assert.ThrowsAny<IOException>(() => new NpoiExcelExporter().ExportToFile(request, directoryPath));

            // Assert
            Assert.True(Directory.Exists(directoryPath));
            Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(directoryPath),
                Path.GetFileName(directoryPath) + ".*.tmp"));
        }
        finally
        {
            if (Directory.Exists(directoryPath))
                Directory.Delete(directoryPath, true);
        }
    }

    /// <summary>
    /// 测试 - Workbook 导出请求应使用同一 Cell Writer 写入多个 Sheet，并按动态 Key 读取字典值。
    /// </summary>
    [Fact]
    public void Export_WorkbookRequest_ShouldWriteMultipleSheetsAndTypedDynamicKeys()
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook =>
        {
            workbook.AddSheet("订单", new[]
            {
                new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>
                {
                    ["region"] = "华东", ["amount"] = 12.5m
                } }
            }, sheet => sheet
                .HeaderRowIndex(2)
                .DynamicColumns(order => order.CustomFields, TenantColumns));
            workbook.AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } });
        });
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        Assert.Equal(2, workbook.NumberOfSheets);
        var orderSheet = workbook.GetSheet("订单");
        Assert.Equal("地区", orderSheet.GetRow(2).GetCell(1).StringCellValue);
        Assert.Equal("金额", orderSheet.GetRow(2).GetCell(2).StringCellValue);
        Assert.Equal("华东", orderSheet.GetRow(3).GetCell(1).StringCellValue);
        Assert.Equal(CellType.Numeric, orderSheet.GetRow(3).GetCell(2).CellType);
        Assert.Equal("客户 A", workbook.GetSheet("客户").GetRow(1).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 测试 - 同一映射的多个 Sheet 应由一次 Workbook Plan 构建覆盖，并实际写入两个 Sheet。
    /// </summary>
    [Fact]
    public void Export_WorkbookPlan_ShouldBuildOnceForSameMappingSheets()
    {
        // Arrange
        var recorder = new RecordingMappingPlanFactory();
        var request = ExcelExport.Workbook(workbook =>
        {
            workbook.AddSheet("订单一", new[] { new ExportOrder { OrderNo = "O-1" } });
            workbook.AddSheet("订单二", new[] { new ExportOrder { OrderNo = "O-2" } });
        });
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter(mappingPlanFactory: recorder).Export(request, destination);

        // Assert
        Assert.Equal(new[] { 2 }, recorder.WorkbookPlanSheetCounts);
        destination.Position = 0;
        using var workbookResult = WorkbookFactory.Create(destination);
        Assert.Equal("O-1", workbookResult.GetSheet("订单一").GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("O-2", workbookResult.GetSheet("订单二").GetRow(1).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 测试 - 异构 Workbook 导入应按每个 Sheet 的 HeaderRowIndex/DataRowStartIndex 解析别名和动态类型。
    /// </summary>
    [Fact]
    public void Import_WorkbookRequest_ShouldUsePerSheetHeaderAndDynamicDefinitions()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var orders = workbook.CreateSheet("订单");
            orders.CreateRow(2).CreateCell(0).SetCellValue("订单号");
            orders.GetRow(2).CreateCell(1).SetCellValue("旧地区");
            orders.GetRow(2).CreateCell(2).SetCellValue("金额");
            orders.CreateRow(3).CreateCell(0).SetCellValue("O-1");
            orders.GetRow(3).CreateCell(1).SetCellValue("华东");
            orders.GetRow(3).CreateCell(2).SetCellValue("12.5");
            var details = workbook.CreateSheet("明细");
            details.CreateRow(1).CreateCell(0).SetCellValue("订单号");
            details.GetRow(1).CreateCell(1).SetCellValue("名称");
            details.CreateRow(2).CreateCell(0).SetCellValue("O-1");
            details.GetRow(2).CreateCell(1).SetCellValue("商品");
        }));
        var request = ExcelImport.Workbook<OrderWorkbook>(workbook =>
        {
            workbook.Sheet("订单", root => root.Orders, sheet => sheet
                .HeaderRowIndex(2)
                .DataRowStartIndex(3)
                .DynamicColumns(order => order.CustomFields, TenantColumns));
            workbook.Sheet("明细", root => root.Details, sheet => sheet
                .HeaderRowIndex(1)
                .DataRowStartIndex(2));
            workbook.HasMany(root => root.Orders, root => root.Details,
                order => order.OrderNo, detail => detail.OrderNo,
                order => order.DetailItems);
        });

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.IsSuccess);
        var order = Assert.Single(result.Workbook.Orders);
        Assert.Equal("华东", order.CustomFields["region"]);
        Assert.Equal(12.5m, order.CustomFields["amount"]);
        Assert.Equal("商品", Assert.Single(order.DetailItems).Name);
        Assert.Equal(2, result.Sheets.Count);
        Assert.All(result.Sheets, sheet => Assert.Empty(sheet.Errors));
    }

    /// <summary>
    /// 测试 - 泛型整型关系键应绑定成功，找不到父项时错误应保留子 Sheet、行号和 Key。
    /// </summary>
    [Fact]
    public void Import_RelationWithNonStringKey_ShouldKeepSourceLocation()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var parents = workbook.CreateSheet("Parents");
            parents.CreateRow(0).CreateCell(0).SetCellValue("Id");
            parents.CreateRow(1).CreateCell(0).SetCellValue(1);
            var children = workbook.CreateSheet("Children");
            children.CreateRow(0).CreateCell(0).SetCellValue("ParentId");
            children.CreateRow(1).CreateCell(0).SetCellValue(2);
        }));
        var request = ExcelImport.Workbook<IntRelationWorkbook>(builder =>
        {
            builder.Sheet("Parents", root => root.Parents);
            builder.Sheet("Children", root => root.Children);
            builder.HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId, parent => parent.Children);
        });

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.Errors.Count == 1,
            string.Join("; ", result.Errors.Select(error =>
                $"{error.Code}|{error.SheetName}|{error.RowIndex}|{error.ColumnKey}|{error.RawValue}")));
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Relationship, error.Code);
        Assert.Equal("Children", error.SheetName);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal("Key", error.ColumnKey);
        Assert.Equal(2, error.RawValue);
    }

    /// <summary>
    /// 测试 - 关系绑定应使用调用方提供的 comparer，大小写不同的字符串键也应关联成功。
    /// </summary>
    [Fact]
    public void Import_RelationWithCustomComparer_ShouldBindCaseInsensitiveKeys()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var parents = workbook.CreateSheet("Parents");
            parents.CreateRow(0).CreateCell(0).SetCellValue("订单号");
            parents.CreateRow(1).CreateCell(0).SetCellValue("A-1");
            var children = workbook.CreateSheet("Children");
            children.CreateRow(0).CreateCell(0).SetCellValue("订单号");
            children.GetRow(0).CreateCell(1).SetCellValue("名称");
            children.CreateRow(1).CreateCell(0).SetCellValue("a-1");
            children.GetRow(1).CreateCell(1).SetCellValue("商品");
        }));
        var request = ExcelImport.Workbook<OrderWorkbook>(builder =>
        {
            builder.Sheet("Parents", root => root.Orders);
            builder.Sheet("Children", root => root.Details);
            builder.HasMany(root => root.Orders, root => root.Details,
                parent => parent.OrderNo, child => child.OrderNo,
                parent => parent.DetailItems, StringComparer.OrdinalIgnoreCase);
        });

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        Assert.Equal("商品", Assert.Single(Assert.Single(result.Workbook.Orders).DetailItems).Name);
    }

    /// <summary>
    /// 测试 - fixed 与 dynamic 列应由同一不可变计划承载执行元数据和预编译访问器。
    /// </summary>
    [Fact]
    public void ColumnPlan_FixedAndDynamic_ShouldShareCompiledExecutionMetadata()
    {
        // Arrange
        var map = new Bing.Offices.Mappings.ExcelMappingPlanFactory().Create<ImportOrder>(new ExcelMappingDocument
        {
            UseConventionFallback = true
        }, null,
            MappingDirection.Import);
        var fixedProperty = map.Columns.Single(property => property.Name == nameof(ImportOrder.OrderNo));
        var dynamicProperty = map.Columns.Single(property => property.Name == nameof(ImportOrder.CustomFields));
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "地区",
            DataType = typeof(string),
            ConverterName = "region",
            ValidatorName = "required"
        };
        var fixedPlan = new ExcelColumnPlan(fixedProperty.Title, fixedProperty, false, 0, null,
            reflectionProperty: typeof(ImportOrder).GetProperty(nameof(ImportOrder.OrderNo)));
        var dynamicPlan = new ExcelColumnPlan(definition.Title, dynamicProperty, true, 1, definition,
            definition.Key, reflectionProperty: typeof(ImportOrder).GetProperty(nameof(ImportOrder.CustomFields)));

        // Assert
        Assert.NotNull(fixedPlan.Getter);
        Assert.NotNull(fixedPlan.Setter);
        Assert.Equal(typeof(string), fixedPlan.ValueType);
        Assert.Equal(typeof(string), dynamicPlan.ValueType);
        Assert.Equal("region", dynamicPlan.Key);
        Assert.Equal("region", dynamicPlan.ConverterName);
        Assert.Equal("required", dynamicPlan.ValidatorName);
        Assert.Equal(ExcelImageMultiplicityPolicy.First, dynamicPlan.ImageMultiplicity);
    }

    /// <summary>
    /// 测试 - Before/After 与物理列索引不能同时配置。
    /// </summary>
    [Fact]
    public void DynamicColumnPlacement_CombinedRelativeAndPhysicalIndex_ShouldFail()
    {
        // Arrange
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "地区",
            Placement = ExcelColumnPlacement.Before("OrderNo"),
            PhysicalColumnIndex = 1
        };

        // Act
        var action = () => ExcelExport.Workbook(workbook => workbook.AddSheet("订单",
            new[] { new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>() } },
            sheet => sheet.DynamicColumns(order => order.CustomFields, new[] { definition })));

        // Assert
        Assert.Throws<ArgumentException>(action);
    }

    /// <summary>
    /// 测试 - 动态值应按 DataType 写为 Numeric Cell，未知 Key 在 Fail 策略下应拒绝导出。
    /// </summary>
    [Fact]
    public void DynamicColumn_DataTypeAndUnknownValuePolicy_ShouldBeEnforced()
    {
        // Arrange
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "amount",
            Title = "金额",
            DataType = typeof(decimal)
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("订单",
            new[] { new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>
            {
                ["amount"] = "12.5", ["unknown"] = "拒绝"
            } } }, sheet => sheet
                .DynamicColumns(order => order.CustomFields, new[] { definition })
                .UnknownDynamicValues(ExcelUnknownDynamicValuePolicy.Fail)));
        using var destination = new MemoryStream();

        // Act
        var action = () => new NpoiExcelExporter().Export(request, destination);

        // Assert
        Assert.Throws<InvalidOperationException>(action);

        var validRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("订单",
            new[] { new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>
            {
                ["amount"] = "12.5"
            } } }, sheet => sheet.DynamicColumns(order => order.CustomFields, new[] { definition })));
        new NpoiExcelExporter().Export(validRequest, destination);
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        Assert.Equal(CellType.Numeric, workbook.GetSheet("订单").GetRow(1).GetCell(1).CellType);
    }

    /// <summary>
    /// 测试 - Before 放置应改变最终物理列顺序，并保持固定列可读。
    /// </summary>
    [Fact]
    public void DynamicColumnPlacement_BeforeFixedColumn_ShouldChangePhysicalOrder()
    {
        // Arrange
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "地区",
            Placement = ExcelColumnPlacement.Before("OrderNo")
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("订单",
            new[] { new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "华东"
            } } }, sheet => sheet.DynamicColumns(order => order.CustomFields, new[] { definition })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var sheet = workbook.GetSheet("订单");
        Assert.Equal("地区", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("OrderNo", sheet.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal("华东", sheet.GetRow(1).GetCell(0).StringCellValue);

        var legacyDocument = new ExcelMappingDocument
        {
            Export = new ExcelMappingConfiguration
            {
                DynamicColumns =
                {
                    new ExcelMappingDynamicColumnConfiguration
                    {
                        Key = "region", Title = "地区", DataTypeName = "string",
                        PlacementKey = "before-OrderNo"
                    }
                }
            }
        };
        using var legacyDestination = new MemoryStream();
        new NpoiExcelExporter().Export(ExcelExport.Workbook(workbook => workbook.AddSheet("订单",
            new[] { new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "华东"
            } } }, sheet => sheet.Mapping(legacyDocument))), legacyDestination);
        legacyDestination.Position = 0;
        using var legacyWorkbook = WorkbookFactory.Create(legacyDestination);
        Assert.Equal("地区", legacyWorkbook.GetSheet("订单").GetRow(0).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 测试 - 模板命名区域应确定写入起点，保留模板其它内容，并按 leaveOpen 约定处理输入流。
    /// </summary>
    [Fact]
    public void Export_TemplateRegion_ShouldPreserveTemplateAndKeepInputOpenWhenRequested()
    {
        // Arrange
        using var template = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("订单");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("保留内容");
            var name = workbook.CreateName();
            name.NameName = "OrderTable";
            name.RefersToFormula = "'订单'!$B$3:$C$20";
            workbook.CreateSheet("说明").CreateRow(0).CreateCell(0).SetCellValue("说明");
        }));
        var request = ExcelExport.Workbook(workbook =>
        {
            workbook.UseTemplate(template, leaveOpen: true)
                .AddSheet("订单", new[] { new ExportCustomer { Name = "客户 A" } },
                    sheet => sheet.UseTemplateRegion("OrderTable"));
        });
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        Assert.True(template.CanRead);
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        Assert.Equal("保留内容", result.GetSheet("订单").GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("客户 A", result.GetSheet("订单").GetRow(3).GetCell(1).StringCellValue);
        Assert.Equal("说明", result.GetSheet("说明").GetRow(0).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 测试 - 同一 Workbook 中相同请求样式应复用样式定义，并保留 XLSX 自定义颜色。
    /// </summary>
    [Fact]
    public void Export_RequestStyle_ShouldReuseStyleAndWriteCustomXlsxColor()
    {
        // Arrange
        var style = new ExcelCellStyle
        {
            Bold = true,
            FillPattern = ExcelFillPattern.Solid,
            ForegroundColor = new ExcelColor("FF123456")
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("客户",
            new[] { new ExportCustomer { Name = "客户 A" }, new ExportCustomer { Name = "客户 B" } },
            sheet => sheet.HeaderStyle(style).BodyStyle(style)));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var sheet = workbook.GetSheet("客户");
        Assert.Equal(sheet.GetRow(0).GetCell(0).CellStyle.Index, sheet.GetRow(1).GetCell(0).CellStyle.Index);
        Assert.Equal("123456", BitConverter.ToString(((XSSFCellStyle)sheet.GetRow(0).GetCell(0).CellStyle)
            .FillForegroundXSSFColor.RGB).Replace("-", string.Empty));
    }

    /// <summary>
    /// 测试 - 宽表头应用 HeaderAttribute 时应复用 Workbook 级字体和样式，并保留请求样式。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xlsx)]
    [InlineData(ExcelFormat.Xls)]
    public void Export_HeaderAttribute_ShouldReuseFontsAndStylesForWideHeaders(ExcelFormat format)
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook =>
        {
            workbook.Format(format);
            workbook.AddSheet("表头", new[]
            {
                new HeaderStyleOrder
                {
                    A = "A", B = "B", C = "C", D = "D", E = "E", F = "F",
                    G = "G", H = "H", I = "I", J = "J", K = "K", L = "L"
                }
            }, sheet => sheet.HeaderStyle(new ExcelCellStyle
            {
                FillPattern = ExcelFillPattern.Solid,
                ForegroundColor = new ExcelColor(format == ExcelFormat.Xlsx ? "FF112233" : "FF0000FF")
            }));
        });
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var header = workbook.GetSheet("表头").GetRow(0);
        var first = header.GetCell(0);
        var last = header.GetCell(11);
        var font = workbook.GetFontAt(first.CellStyle.FontIndex);
        Assert.Equal(first.CellStyle.Index, last.CellStyle.Index);
        Assert.Equal(first.CellStyle.FontIndex, last.CellStyle.FontIndex);
        Assert.Equal("Arial", font.FontName);
        Assert.Equal(11, font.FontHeightInPoints);
        Assert.False(font.IsBold);
        Assert.Equal(FillPattern.SolidForeground, first.CellStyle.FillPattern);
        if (format == ExcelFormat.Xlsx)
        {
            Assert.Equal("112233", BitConverter.ToString(((XSSFCellStyle)first.CellStyle)
                .FillForegroundXSSFColor.RGB).Replace("-", string.Empty));
        }
        else
        {
            Assert.Equal(IndexedColors.Blue.Index, first.CellStyle.FillForegroundColor);
        }
        var styleCeiling = format == ExcelFormat.Xlsx ? 8 : 23;
        Assert.True(workbook.NumCellStyles <= styleCeiling,
            $"宽表头样式数量异常: {workbook.NumCellStyles}, 上限: {styleCeiling}");
        var fontCeiling = format == ExcelFormat.Xlsx ? 4 : 5;
        Assert.True(workbook.NumberOfFonts <= fontCeiling,
            $"宽表头字体数量异常: {workbook.NumberOfFonts}, 上限: {fontCeiling}");
    }

    /// <summary>
    /// 测试 - 模板样式叠加时未指定属性应保留，显式 reset 只恢复指定属性。
    /// </summary>
    [Fact]
    public void Export_RequestStyle_Compose_ShouldPreserveAndResetSelectedProperties()
    {
        // Arrange
        using var template = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("客户");
            var style = workbook.CreateCellStyle();
            style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");
            style.FillPattern = FillPattern.SolidForeground;
            style.FillForegroundColor = IndexedColors.Yellow.Index;
            var cell = sheet.CreateRow(1).CreateCell(0);
            cell.CellStyle = style;
            var name = workbook.CreateName();
            name.NameName = "CustomerRegion";
            name.RefersToFormula = "'客户'!$A$1:$A$2";
        }));
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } },
                sheet => sheet.UseTemplateRegion("CustomerRegion")
                    .BodyStyle(new ExcelCellStyle
                    {
                        Bold = true,
                        Reset = new ExcelCellStyleReset { FillPattern = true }
                    })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        var cell = result.GetSheet("客户").GetRow(1).GetCell(0);
        Assert.True(cell.CellStyle.GetDataFormatString() == "0.00");
        Assert.True(cell.CellStyle.Index > 0);
        Assert.True(cell.CellStyle.FillPattern == FillPattern.NoFill);
        Assert.True(result.GetFontAt(cell.CellStyle.FontIndex).IsBold);
    }

    /// <summary>
    /// 测试 - XSSF 前景色和背景色应独立写入，不能因只设置背景色而改变前景色。
    /// </summary>
    [Fact]
    public void Export_RequestStyle_ShouldKeepXlsxForegroundAndBackgroundIndependent()
    {
        // Arrange
        var style = new ExcelCellStyle
        {
            FillPattern = ExcelFillPattern.Solid,
            ForegroundColor = new ExcelColor("FF112233"),
            BackgroundColor = new ExcelColor("FF445566")
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("客户",
            new[] { new ExportCustomer { Name = "客户 A" } }, sheet => sheet.BodyStyle(style)));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = (XSSFWorkbook)WorkbookFactory.Create(destination);
        var cellStyle = (XSSFCellStyle)result.GetSheet("客户").GetRow(1).GetCell(0).CellStyle;
        Assert.Equal("112233", BitConverter.ToString(cellStyle.FillForegroundXSSFColor.RGB)
            .Replace("-", string.Empty));
        Assert.Equal("445566", BitConverter.ToString(cellStyle.FillBackgroundXSSFColor.RGB)
            .Replace("-", string.Empty));
    }

    /// <summary>
    /// 测试 - 同一 Workbook 中字体 reset 与未指定 reset 的样式不得共享错误字体。
    /// </summary>
    [Fact]
    public void Export_RequestStyle_FontReset_ShouldIsolateFontCacheEntries()
    {
        // Arrange
        using var workbook = new XSSFWorkbook();
        var baseFont = workbook.CreateFont();
        baseFont.IsBold = true;
        var baseStyle = workbook.CreateCellStyle();
        baseStyle.SetFont(baseFont);
        var resetStyle = new ExcelCellStyle
        {
            Reset = new ExcelCellStyleReset { Bold = true }
        };
        var preserveStyle = new ExcelCellStyle();

        // Act
        var reset = NpoiStyleCache.Compose(workbook, baseStyle, resetStyle);
        var preserve = NpoiStyleCache.Compose(workbook, baseStyle, preserveStyle);

        // Assert
        Assert.False(workbook.GetFontAt(reset.FontIndex).IsBold);
        Assert.True(workbook.GetFontAt(preserve.FontIndex).IsBold);
        Assert.NotEqual(reset.FontIndex, preserve.FontIndex);

        var preserveFirst = NpoiStyleCache.Compose(workbook, baseStyle, new ExcelCellStyle());
        var resetAfterPreserve = NpoiStyleCache.Compose(workbook, baseStyle, resetStyle);
        Assert.True(workbook.GetFontAt(preserveFirst.FontIndex).IsBold);
        Assert.False(workbook.GetFontAt(resetAfterPreserve.FontIndex).IsBold);
    }

    /// <summary>
    /// 测试 - 样式 reset 应逐属性恢复模板默认状态，并保持缓存结果可重用。
    /// </summary>
    [Fact]
    public void Export_RequestStyle_Reset_ShouldRestoreAllCoveredProperties()
    {
        // Arrange
        using var workbook = new XSSFWorkbook();
        var baseStyle = workbook.CreateCellStyle();
        baseStyle.FillPattern = FillPattern.SolidForeground;
        baseStyle.Alignment = HorizontalAlignment.Center;
        baseStyle.VerticalAlignment = VerticalAlignment.Center;
        baseStyle.WrapText = true;
        baseStyle.Indention = 3;
        baseStyle.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");
        baseStyle.BorderTop = BorderStyle.Thin;
        baseStyle.BorderBottom = BorderStyle.Thin;
        baseStyle.BorderLeft = BorderStyle.Thin;
        baseStyle.BorderRight = BorderStyle.Thin;

        // Act
        var reset = (XSSFCellStyle)NpoiStyleCache.Compose(workbook, baseStyle, new ExcelCellStyle
        {
            Reset = new ExcelCellStyleReset
            {
                FillPattern = true,
                HorizontalAlignment = true,
                VerticalAlignment = true,
                WrapText = true,
                Indent = true,
                NumberFormat = true,
                TopBorder = true,
                BottomBorder = true,
                LeftBorder = true,
                RightBorder = true
            }
        });

        // Assert
        Assert.Equal(FillPattern.NoFill, reset.FillPattern);
        Assert.Equal(HorizontalAlignment.General, reset.Alignment);
        Assert.Equal(VerticalAlignment.Bottom, reset.VerticalAlignment);
        Assert.False(reset.WrapText);
        Assert.Equal(0, reset.Indention);
        Assert.Equal("General", reset.GetDataFormatString());
        Assert.Equal(BorderStyle.None, reset.BorderTop);
        Assert.Equal(BorderStyle.None, reset.BorderBottom);
        Assert.Equal(BorderStyle.None, reset.BorderLeft);
        Assert.Equal(BorderStyle.None, reset.BorderRight);
    }

    /// <summary>
    /// 测试 - 公开导出请求中的固定列和动态列样式 reset 应在 XLS/XLSX 重开后保持契约，且重复行不应无限创建资源。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xlsx)]
    [InlineData(ExcelFormat.Xls)]
    public void Export_RequestStyle_PublicFixedAndDynamicColumns_ShouldResetAndBoundResources(ExcelFormat format)
    {
        // Arrange
        var rows = Enumerable.Range(1, 32).Select(index => new ExportOrder
        {
            OrderNo = $"O-{index}",
            CustomFields = new Dictionary<string, object> { ["amount"] = index + 0.5m }
        }).ToArray();
        var request = ExcelExport.Workbook(workbook => workbook
            .Format(format)
            .AddSheet("订单", rows, sheet => sheet
                .HeaderStyle(new ExcelCellStyle { Bold = true, FillPattern = ExcelFillPattern.Solid,
                    Reset = new ExcelCellStyleReset { ForegroundColor = true, BackgroundColor = true } })
                .BodyStyle(new ExcelCellStyle { Bold = true, HorizontalAlignment = ExcelHorizontalAlignment.Center,
                    NumberFormat = "0.00", Reset = new ExcelCellStyleReset { FillPattern = true } })
                .DynamicColumns(order => order.CustomFields, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "amount",
                        Title = "金额",
                        DataType = typeof(decimal),
                        NumberFormat = "0.00",
                        HeaderStyle = new ExcelCellStyle { Italic = true },
                        BodyStyle = new ExcelCellStyle { Underline = true, NumberFormat = "0.00" }
                    }
                })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        var sheet = result.GetSheet("订单");
        var fixedHeader = sheet.GetRow(0).GetCell(0);
        var fixedBody = sheet.GetRow(1).GetCell(0);
        var dynamicHeader = sheet.GetRow(0).GetCell(1);
        var dynamicBody = sheet.GetRow(1).GetCell(1);
        Assert.True(result.GetFontAt(fixedHeader.CellStyle.FontIndex).IsBold);
        Assert.True(result.GetFontAt(fixedBody.CellStyle.FontIndex).IsBold);
        Assert.True(result.GetFontAt(dynamicHeader.CellStyle.FontIndex).IsItalic);
        Assert.NotEqual(FontUnderlineType.None, result.GetFontAt(dynamicBody.CellStyle.FontIndex).Underline);
        Assert.Equal("0.00", fixedBody.CellStyle.GetDataFormatString());
        Assert.Equal("0.00", dynamicBody.CellStyle.GetDataFormatString());
        Assert.Equal(FillPattern.NoFill, fixedBody.CellStyle.FillPattern);
        Assert.True(result.NumCellStyles <= 32, $"样式数量异常: {result.NumCellStyles}");
        Assert.True(result.NumberOfFonts <= 8, $"字体数量异常: {result.NumberOfFonts}");
    }

    /// <summary>
    /// 测试 - 公共导出入口应从模板非默认样式清除 Sheet、固定列和动态列的全部 reset 属性。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xlsx)]
    [InlineData(ExcelFormat.Xls)]
    public void Export_RequestStyle_PublicReset_ShouldClearNonDefaultTemplateProperties(ExcelFormat format)
    {
        // Arrange
        using var template = new MemoryStream(CreateStyledExportTemplate(format));
        var reset = new ExcelCellStyleReset
        {
            FontName = true,
            FontSize = true,
            Bold = true,
            Italic = true,
            Underline = true,
            FontColor = true,
            ForegroundColor = true,
            BackgroundColor = true,
            FillPattern = true,
            TopBorder = true,
            BottomBorder = true,
            LeftBorder = true,
            RightBorder = true,
            HorizontalAlignment = true,
            VerticalAlignment = true,
            WrapText = true,
            Indent = true,
            NumberFormat = true
        };
        var rows = new[]
        {
            new ExportOrder
            {
                OrderNo = "O-1",
                CustomFields = new Dictionary<string, object> { ["amount"] = 12.5m }
            }
        };
        var request = ExcelExport.Workbook(workbook =>
        {
            workbook.Format(format)
                .UseTemplate(template, leaveOpen: true)
                .AddSheet("订单", rows, sheet => sheet
                    .SheetStyle(new ExcelCellStyle { Reset = reset })
                    .HeaderStyle(new ExcelCellStyle { Reset = reset })
                    .BodyStyle(new ExcelCellStyle { Reset = reset })
                    .DynamicColumns(order => order.CustomFields, new[]
                    {
                        new ExcelDynamicColumnDefinition
                        {
                            Key = "amount",
                            Title = "金额",
                            DataType = typeof(decimal),
                            HeaderStyle = new ExcelCellStyle { Reset = reset },
                            BodyStyle = new ExcelCellStyle { Reset = reset }
                        }
                    }))
                .AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } },
                    sheet => sheet.SheetStyle(new ExcelCellStyle { Reset = reset }));
        });
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        AssertResetStyle(result.GetSheet("订单").GetRow(0).GetCell(0).CellStyle,
            result.GetSheet("订单").GetRow(0).GetCell(0));
        AssertResetStyle(result.GetSheet("订单").GetRow(1).GetCell(0).CellStyle,
            result.GetSheet("订单").GetRow(1).GetCell(0));
        AssertResetStyle(result.GetSheet("订单").GetRow(0).GetCell(1).CellStyle,
            result.GetSheet("订单").GetRow(0).GetCell(1));
        AssertResetStyle(result.GetSheet("订单").GetRow(1).GetCell(1).CellStyle,
            result.GetSheet("订单").GetRow(1).GetCell(1));
        AssertResetStyle(result.GetSheet("客户").GetRow(0).GetCell(0).CellStyle,
            result.GetSheet("客户").GetRow(0).GetCell(0));
        AssertResetStyle(result.GetSheet("客户").GetRow(1).GetCell(0).CellStyle,
            result.GetSheet("客户").GetRow(1).GetCell(0));
        Assert.True(result.NumCellStyles <= 40, $"样式数量异常: {result.NumCellStyles}");
        Assert.True(result.NumberOfFonts <= 12, $"字体数量异常: {result.NumberOfFonts}");
    }

    /// <summary>
    /// 测试 - 样式 reset 先调用或后调用均不得与保留样式共享错误的字体，且两个提供程序行为一致。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xlsx, true)]
    [InlineData(ExcelFormat.Xlsx, false)]
    [InlineData(ExcelFormat.Xls, true)]
    [InlineData(ExcelFormat.Xls, false)]
    public void Export_RequestStyle_ResetOrder_ShouldIsolateFonts(ExcelFormat format, bool resetFirst)
    {
        // Arrange
        using var workbook = format == ExcelFormat.Xls
            ? (IWorkbook)new HSSFWorkbook()
            : new XSSFWorkbook();
        var baseFont = workbook.CreateFont();
        baseFont.IsBold = true;
        var baseStyle = workbook.CreateCellStyle();
        baseStyle.SetFont(baseFont);
        var resetStyle = new ExcelCellStyle { Reset = new ExcelCellStyleReset { Bold = true } };
        var preserveStyle = new ExcelCellStyle();

        // Act
        var first = resetFirst
            ? NpoiStyleCache.Compose(workbook, baseStyle, resetStyle)
            : NpoiStyleCache.Compose(workbook, baseStyle, preserveStyle);
        var second = resetFirst
            ? NpoiStyleCache.Compose(workbook, baseStyle, preserveStyle)
            : NpoiStyleCache.Compose(workbook, baseStyle, resetStyle);

        // Assert
        var reset = resetFirst ? first : second;
        var preserve = resetFirst ? second : first;
        Assert.False(workbook.GetFontAt(reset.FontIndex).IsBold);
        Assert.True(workbook.GetFontAt(preserve.FontIndex).IsBold);
        Assert.NotEqual(reset.FontIndex, preserve.FontIndex);
        Assert.True(workbook.NumCellStyles <= 32, $"样式数量异常: {workbook.NumCellStyles}");
        Assert.True(workbook.NumberOfFonts <= 6, $"字体数量异常: {workbook.NumberOfFonts}");
    }

    /// <summary>
    /// 测试 - XSSF 显式清除前景色和背景色后应恢复无色状态，而不是写入黑色。
    /// </summary>
    [Fact]
    public void Export_RequestStyle_XlsxFillColorReset_ShouldClearColors()
    {
        // Arrange
        using var workbook = new XSSFWorkbook();
        var baseStyle = workbook.CreateCellStyle();
        var baseXssfStyle = (XSSFCellStyle)baseStyle;
        baseXssfStyle.SetFillForegroundColor(new XSSFColor(new byte[] { 17, 34, 51 }, null));
        baseXssfStyle.SetFillBackgroundColor(new XSSFColor(new byte[] { 68, 85, 102 }, null));
        baseXssfStyle.FillPattern = FillPattern.SolidForeground;

        // Act
        var result = (XSSFCellStyle)NpoiStyleCache.Compose(workbook, baseStyle, new ExcelCellStyle
        {
            Reset = new ExcelCellStyleReset { ForegroundColor = true, BackgroundColor = true }
        });

        // Assert
        Assert.Null(result.FillForegroundColorColor);
        Assert.Null(result.FillBackgroundColorColor);
    }

    /// <summary>
    /// 测试 - XLS 支持的 indexed 颜色应写入，无法表示的自定义颜色应明确拒绝。
    /// </summary>
    [Fact]
    public void Export_RequestStyle_Xls_ShouldHonorColorCapabilityBoundary()
    {
        // Arrange
        var supported = ExcelExport.Workbook(workbook => workbook.Format(ExcelFormat.Xls).AddSheet("客户",
            new[] { new ExportCustomer { Name = "客户 A" } }, sheet => sheet.BodyStyle(new ExcelCellStyle
            {
                FillPattern = ExcelFillPattern.Solid,
                ForegroundColor = new ExcelColor("FFFF0000")
            })));
        var unsupported = ExcelExport.Workbook(workbook => workbook.Format(ExcelFormat.Xls).AddSheet("客户",
            new[] { new ExportCustomer { Name = "客户 A" } }, sheet => sheet.BodyStyle(new ExcelCellStyle
            {
                TopBorder = new ExcelBorderStyle
                {
                    LineStyle = ExcelBorderLineStyle.Thin,
                    Color = new ExcelColor("FF123456")
                }
            })));

        // Act
        using var supportedOutput = new MemoryStream();
        new NpoiExcelExporter().Export(supported, supportedOutput);
        using var unsupportedOutput = new MemoryStream();

        // Assert
        supportedOutput.Position = 0;
        using var workbook = (HSSFWorkbook)WorkbookFactory.Create(supportedOutput);
        Assert.Equal(IndexedColors.Red.Index, workbook.GetSheet("客户").GetRow(1).GetCell(0)
            .CellStyle.FillForegroundColor);
        Assert.Throws<NotSupportedException>(() => new NpoiExcelExporter().Export(unsupported, unsupportedOutput));
    }

    /// <summary>
    /// 测试 - 请求样式的 XSSF 四边框自定义颜色应按 RGB 写入并在重开文件后保持一致。
    /// </summary>
    [Fact]
    public void Export_RequestStyle_ShouldWriteCustomXlsxBorderColors()
    {
        // Arrange
        var style = new ExcelCellStyle
        {
            TopBorder = new ExcelBorderStyle
            {
                LineStyle = ExcelBorderLineStyle.Thin,
                Color = new ExcelColor("FF112233")
            },
            BottomBorder = new ExcelBorderStyle
            {
                LineStyle = ExcelBorderLineStyle.Thin,
                Color = new ExcelColor("FF445566")
            },
            LeftBorder = new ExcelBorderStyle
            {
                LineStyle = ExcelBorderLineStyle.Thin,
                Color = new ExcelColor("FF778899")
            },
            RightBorder = new ExcelBorderStyle
            {
                LineStyle = ExcelBorderLineStyle.Thin,
                Color = new ExcelColor("FFAABBCC")
            }
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("客户",
            new[] { new ExportCustomer { Name = "客户 A" } }, sheet => sheet.BodyStyle(style)));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var cellStyle = (XSSFCellStyle)workbook.GetSheet("客户").GetRow(1).GetCell(0).CellStyle;
        Assert.Equal("112233", BitConverter.ToString(cellStyle.TopBorderXSSFColor.RGB)
            .Replace("-", string.Empty));
        Assert.Equal("445566", BitConverter.ToString(cellStyle.BottomBorderXSSFColor.RGB)
            .Replace("-", string.Empty));
        Assert.Equal("778899", BitConverter.ToString(cellStyle.LeftBorderXSSFColor.RGB)
            .Replace("-", string.Empty));
        Assert.Equal("AABBCC", BitConverter.ToString(cellStyle.RightBorderXSSFColor.RGB)
            .Replace("-", string.Empty));
    }

    /// <summary>
    /// 测试 - 请求样式应逐属性叠加到模板样式，不得覆盖模板数字格式。
    /// </summary>
    [Fact]
    public void Export_RequestStyle_ShouldPreserveTemplateNumberFormat()
    {
        // Arrange
        using var template = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("订单");
            var style = workbook.CreateCellStyle();
            style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");
            sheet.CreateRow(3).CreateCell(1).CellStyle = style;
            var name = workbook.CreateName();
            name.NameName = "OrderTable";
            name.RefersToFormula = "'订单'!$B$3:$C$20";
        }));
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("订单", new[] { new ExportCustomer { Name = "客户 A" } },
                sheet => sheet.UseTemplateRegion("OrderTable")
                    .BodyStyle(new ExcelCellStyle { Bold = true })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        var cell = result.GetSheet("订单").GetRow(3).GetCell(1);
        Assert.Equal("0.00", cell.CellStyle.GetDataFormatString());
        Assert.True(result.GetSheet("订单").GetRow(3).GetCell(1).CellStyle.Index > 0);
    }

    /// <summary>
    /// 测试 - 默认模板策略应保留表头和正文的模板样式、批注，并用导出值替换公式。
    /// </summary>
    [Fact]
    public void Export_TemplateCellOverwrite_Default_ShouldPreserveTemplateMetadata()
    {
        // Arrange
        using var template = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("客户");
            var headerStyle = workbook.CreateCellStyle();
            headerStyle.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");
            var header = sheet.CreateRow(0).CreateCell(0);
            header.SetCellFormula("1+1");
            header.CellStyle = headerStyle;
            var body = sheet.CreateRow(1).CreateCell(0);
            body.SetCellFormula("2+2");
            body.CellStyle = headerStyle;
            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = 0;
            anchor.Col2 = 2;
            anchor.Row1 = 0;
            anchor.Row2 = 3;
            var comment = sheet.CreateDrawingPatriarch().CreateCellComment(anchor);
            comment.String = workbook.GetCreationHelper().CreateRichTextString("模板批注");
            comment.Author = "template";
            header.CellComment = comment;
        }));
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } }));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        var sheet = result.GetSheet("客户");
        Assert.Equal("Name", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("客户 A", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("0.00", sheet.GetRow(0).GetCell(0).CellStyle.GetDataFormatString());
        Assert.Equal("0.00", sheet.GetRow(1).GetCell(0).CellStyle.GetDataFormatString());
        Assert.NotNull(sheet.GetRow(0).GetCell(0).CellComment);
        Assert.Equal("模板批注", sheet.GetRow(0).GetCell(0).CellComment.String.String);
    }

    /// <summary>
    /// 测试 - ReplaceTemplate 策略应清除模板样式和批注后写入导出值。
    /// </summary>
    [Fact]
    public void Export_TemplateCellOverwrite_Replace_ShouldClearTemplateMetadata()
    {
        // Arrange
        using var template = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("客户");
            var style = workbook.CreateCellStyle();
            style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");
            var header = sheet.CreateRow(0).CreateCell(0);
            header.CellStyle = style;
            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = 0;
            anchor.Col2 = 2;
            anchor.Row1 = 0;
            anchor.Row2 = 3;
            var comment = sheet.CreateDrawingPatriarch().CreateCellComment(anchor);
            comment.String = workbook.GetCreationHelper().CreateRichTextString("模板批注");
            header.CellComment = comment;
        }));
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("客户", new[] { new ExportCustomer { Name = "客户 A" } },
                sheet => sheet.TemplateCellOverwrite(ExcelTemplateCellOverwritePolicy.ReplaceTemplate)));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        var sheet = result.GetSheet("客户");
        Assert.Equal("Name", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("客户 A", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("General", sheet.GetRow(0).GetCell(0).CellStyle.GetDataFormatString());
        Assert.Null(sheet.GetRow(0).GetCell(0).CellComment);
    }

    /// <summary>
    /// 测试 - 模板区域起点应同时偏移自定义表头的相对行列。
    /// </summary>
    [Fact]
    public void Export_TemplateRegion_ShouldOffsetCustomHeaders()
    {
        // Arrange
        using var template = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("订单");
            var name = workbook.CreateName();
            name.NameName = "OrderTable";
            name.RefersToFormula = "'订单'!$B$3:$C$20";
        }));
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("订单", new[] { new ExportCustomer { Name = "客户 A" } },
                sheet => sheet.UseTemplateRegion("OrderTable")
                    .HeaderRowIndex(1)
                    .DataRowStartIndex(2)
                    .HeaderRows(new[] { new ExcelHeaderRow(0,
                        new[] { new ExcelHeaderCell(2, "扩展表头") }) })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        Assert.Equal("扩展表头", result.GetSheet("订单").GetRow(2).GetCell(3).StringCellValue);
    }

    /// <summary>
    /// 测试 - 动态列指定命名转换器时，应在默认类型转换前使用该转换器。
    /// </summary>
    [Fact]
    public void Export_DynamicColumnNamedConverter_ShouldConvertValue()
    {
        // Arrange
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "地区",
            DataType = typeof(string),
            ConverterName = "upper-region"
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("订单",
            new[] { new ExportOrder { OrderNo = "O-1", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "east"
            } } }, sheet => sheet.DynamicColumns(order => order.CustomFields, new[] { definition })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter(new[] { new UpperRegionConverter() }).Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        Assert.Equal("EAST", result.GetSheet("订单").GetRow(1).GetCell(1).StringCellValue);
    }

    /// <summary>
    /// 测试 - Document 动态列、布局、数字格式和样式应进入真实 XLSX 导出计划。
    /// </summary>
    [Fact]
    public void Export_DocumentDynamicPlan_ShouldApplyLayoutFormatAndStyle()
    {
        // Arrange
        var document = new ExcelMappingDocument
        {
            Export = new ExcelMappingConfiguration
            {
                DynamicColumns =
                {
                    new ExcelMappingDynamicColumnConfiguration
                    {
                        Key = "amount", Title = "金额", DataTypeName = "decimal",
                        NumberFormat = "0.00", PlacementKey = "before:OrderNo"
                    }
                },
                Style = new ExcelMappingStyleConfiguration { HeaderStyleKey = "header" }
            }
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("订单", new[]
        {
            new ExportOrder
            {
                OrderNo = "O-1",
                CustomFields = new Dictionary<string, object> { ["amount"] = 12.5m }
            }
        }, sheet => sheet.Mapping(document)));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var sheet = workbook.GetSheet("订单");
        Assert.Equal("金额", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.True(sheet.GetRow(0).GetCell(0).CellStyle.GetFont(workbook).IsBold);
        Assert.Equal(CellType.Numeric, sheet.GetRow(1).GetCell(0).CellType);
        Assert.Equal("0.00", sheet.GetRow(1).GetCell(0).CellStyle.GetDataFormatString());
        Assert.Equal("O-1", sheet.GetRow(1).GetCell(1).StringCellValue);
    }

    /// <summary>
    /// 测试 - Document 命名动态校验器应在 CSV 与 XLSX 主链产生等价失败。
    /// </summary>
    [Fact]
    public void Import_DocumentDynamicValidator_ShouldMatchCsvAndXlsx()
    {
        // Arrange
        var document = new ExcelMappingDocument
        {
            Import = new ExcelMappingConfiguration
            {
                DynamicColumns =
                {
                    new ExcelMappingDynamicColumnConfiguration
                    {
                        Key = "region", Title = "地区", DataTypeName = "string",
                        ValidatorName = "starts-ok",
                        ValidationRuleNames = new List<string> { "contains-region" }
                    }
                }
            }
        };
        var validator = new StartsWithOkValidationRule();
        var secondValidator = new ContainsRegionValidationRule();
        using var xlsx = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("地区");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("BAD");
        }));
        var request = ExcelImport.Workbook<DynamicTargetWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet.Mapping(document)));
        using var csv = new MemoryStream(Encoding.UTF8.GetBytes("地区\r\nBAD\r\n"));

        // Act
        var validators = new INamedExcelValidationRule[] { validator, secondValidator };
        var plan = new ExcelMappingPlanFactory(namedValidationRules: validators)
            .Create<DynamicTargetRow>(document, MappingDirection.Import);
        var dynamicPlan = Assert.Single(plan.DynamicColumns);
        Assert.Equal(new[] { "starts-ok", "contains-region" }, dynamicPlan.ValidationRuleNames);
        Assert.Equal(2, dynamicPlan.ValidationBindings.Count);
        var xlsxResult = new NpoiExcelImporter(namedValidationRules: validators).Import(xlsx, request);
        var csvResult = new CsvEntityImporter(namedValidationRules: validators).Import<DynamicTargetRow>(csv,
            new CsvImportOptions<DynamicTargetRow> { MappingDocument = document, HeaderMatch = false });

        // Assert
        Assert.Single(xlsxResult.Errors);
        Assert.Single(csvResult.Errors);
        Assert.Equal(2, xlsxResult.Errors[0].RowIndex);
        Assert.Equal(2, csvResult.Errors[0].RowIndex);
        Assert.Empty(xlsxResult.Workbook.Rows);
        Assert.Empty(csvResult.Items);
    }

    /// <summary>
    /// 测试 - 七类内置动态规则应在 CSV 与 XLSX 中产生一致失败行和唯一值回滚结果。
    /// </summary>
    [Fact]
    public void Import_DynamicBuiltInValidationMatrix_ShouldMatchCsvAndXlsx()
    {
        // Arrange
        var document = new ExcelMappingDocument
        {
            Import = new ExcelMappingConfiguration
            {
                DynamicColumns = CreateDynamicValidationColumns()
            }
        };
        var headers = new[] { "必填", "正则", "日期", "最大值", "区间", "长度", "唯一" };
        var rows = new[]
        {
            new[] { "ok", "OK-1", "2026-08-24", "5", "5", "abc", "U-1" },
            new[] { "ok", "OK-1", "2026-08-24", "10", "1", "abcde", "U-2" },
            new[] { "ok", "OK-1", "2026-08-24", "10", "10", "abcde", "U-3" },
            new[] { "", "OK-1", "2026-08-24", "5", "5", "abc", "U-4" },
            new[] { "ok", "BAD", "2026-08-24", "5", "5", "abc", "U-5" },
            new[] { "ok", "OK-1", "not-date", "5", "5", "abc", "U-6" },
            new[] { "ok", "OK-1", "2026-08-24", "11", "5", "abc", "U-7" },
            new[] { "ok", "OK-1", "2026-08-24", "5", "0", "abc", "U-8" },
            new[] { "ok", "OK-1", "2026-08-24", "5", "5", "abcdef", "U-9" },
            new[] { "ok", "OK-1", "2026-08-24", "5", "5", "abc", "u-1" }
        };
        using var xlsx = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0);
            for (var column = 0; column < headers.Length; column++)
                sheet.GetRow(0).CreateCell(column).SetCellValue(headers[column]);
            for (var row = 0; row < rows.Length; row++)
            {
                sheet.CreateRow(row + 1);
                for (var column = 0; column < rows[row].Length; column++)
                    sheet.GetRow(row + 1).CreateCell(column).SetCellValue(rows[row][column]);
            }
        }));
        var request = ExcelImport.Workbook<DynamicTargetWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet.Mapping(document)));
        using var csv = new MemoryStream(Encoding.UTF8.GetBytes(
            string.Join("\r\n", new[] { string.Join(",", headers) }
                .Concat(rows.Select(row => string.Join(",", row)))) + "\r\n"));

        // Act
        var plan = new ExcelMappingPlanFactory().Create<DynamicTargetRow>(document, MappingDirection.Import);
        var xlsxResult = new NpoiExcelImporter().Import(xlsx, request);
        var csvResult = new CsvEntityImporter().Import<DynamicTargetRow>(csv,
            new CsvImportOptions<DynamicTargetRow> { MappingDocument = document, HeaderMatch = false });

        // Assert
        Assert.Equal(7, plan.DynamicColumns.Count);
        Assert.All(plan.DynamicColumns, column => Assert.Single(column.ValidationBindings));
        Assert.True(plan.DynamicColumns.Single(column => column.Key == "unique").IsUnique);
        Assert.Equal(7, xlsxResult.Errors.Count);
        Assert.Equal(7, csvResult.Errors.Count);
        Assert.Equal(3, xlsxResult.Workbook.Rows.Count);
        Assert.Equal(3, csvResult.Items.Count);
        var expected = new[]
        {
            (Row: 5, Column: 1, Key: "required", Code: ExcelImportErrorCode.Validation),
            (Row: 6, Column: 2, Key: "regex", Code: ExcelImportErrorCode.Validation),
            (Row: 7, Column: 3, Key: "date", Code: ExcelImportErrorCode.ValueConversion),
            (Row: 8, Column: 4, Key: "max", Code: ExcelImportErrorCode.MaxValue),
            (Row: 9, Column: 5, Key: "range", Code: ExcelImportErrorCode.Validation),
            (Row: 10, Column: 6, Key: "length", Code: ExcelImportErrorCode.MaxLength),
            (Row: 11, Column: 7, Key: "unique", Code: ExcelImportErrorCode.Validation)
        };
        Assert.Equal(expected.Select(item => $"{item.Code}|{item.Row}|{item.Column}|{item.Key}"),
            xlsxResult.Errors.Select(error => $"{error.Code}|{error.RowIndex}|{error.ColumnIndex}|{error.ColumnKey}"));
        Assert.Equal(expected.Select(item => $"{item.Row}|{item.Column}|{item.Key}"),
            csvResult.Errors.Select(error => $"{error.RowIndex}|{error.ColumnIndex}|{error.PropertyName}"));
        Assert.Equal(2, xlsxResult.Errors.Single(error => error.ColumnKey == "unique").FirstRowNumber);
        Assert.Equal(2, csvResult.Errors.Single(error => error.PropertyName == "unique").FirstRowNumber);
    }

    /// <summary>
    /// 测试 - 动态唯一规则应在 CSV 与 XLSX 中一致处理空值、大小写比较和跟踪上限。
    /// </summary>
    [Fact]
    public void Import_DynamicUniqueOptions_ShouldMatchCsvAndXlsx()
    {
        // Arrange
        var values = new[] { " ", " ", "ABC", "abc" };
        byte[] CreateSource() => CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("唯一");
            sheet.GetRow(0).CreateCell(1).SetCellValue("辅助");
            for (var row = 0; row < values.Length; row++)
            {
                sheet.CreateRow(row + 1).CreateCell(0).SetCellValue(values[row]);
                sheet.GetRow(row + 1).CreateCell(1).SetCellValue("x");
            }
        });
        var csvText = "唯一,辅助\r\n,x\r\n,x\r\nABC,x\r\nabc,x\r\n";

        // Act
        var ignoredDocument = CreateDynamicUniqueDocument(ignoreEmpty: true);
        var ignoredXlsx = ImportDynamicUniqueXlsx(CreateSource(), ignoredDocument,
            new ExcelResourceLimits { UniqueComparison = StringComparison.OrdinalIgnoreCase });
        var ignoredCsv = ImportDynamicUniqueCsv(csvText, ignoredDocument,
            new CsvImportOptions<DynamicTargetRow>
            {
                HeaderMatch = false,
                UniqueComparison = StringComparison.OrdinalIgnoreCase
            });

        var trackedDocument = CreateDynamicUniqueDocument(ignoreEmpty: false);
        var trackedXlsx = ImportDynamicUniqueXlsx(CreateSource(), trackedDocument,
            new ExcelResourceLimits { UniqueComparison = StringComparison.OrdinalIgnoreCase });
        var trackedCsv = ImportDynamicUniqueCsv(csvText, trackedDocument,
            new CsvImportOptions<DynamicTargetRow>
            {
                HeaderMatch = false,
                UniqueComparison = StringComparison.OrdinalIgnoreCase
            });

        var limitedValues = new[] { "ABC", "DEF" };
        var limitedSource = CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("唯一");
            for (var row = 0; row < limitedValues.Length; row++)
                sheet.CreateRow(row + 1).CreateCell(0).SetCellValue(limitedValues[row]);
        });
        var limitedDocument = CreateDynamicUniqueDocument(ignoreEmpty: true);
        var limitedXlsx = ImportDynamicUniqueXlsx(limitedSource, limitedDocument,
            new ExcelResourceLimits
            {
                MaxTrackedUniqueValues = 1,
                UniqueComparison = StringComparison.OrdinalIgnoreCase
            });
        var limitedCsv = ImportDynamicUniqueCsv("唯一\r\nABC\r\nDEF\r\n", limitedDocument,
            new CsvImportOptions<DynamicTargetRow>
            {
                HeaderMatch = false,
                MaxTrackedUniqueValues = 1,
                UniqueComparison = StringComparison.OrdinalIgnoreCase
            });

        // Assert
        Assert.Equal(1, ignoredXlsx.Errors.Count);
        Assert.Equal(1, ignoredCsv.Errors.Count);
        Assert.Equal(3, ignoredXlsx.Workbook.Rows.Count);
        Assert.Equal(3, ignoredCsv.Items.Count);
        Assert.Equal(5, ignoredXlsx.Errors[0].RowIndex);
        Assert.Equal(5, ignoredCsv.Errors[0].RowIndex);
        Assert.Equal(4, ignoredXlsx.Errors[0].FirstRowNumber);
        Assert.Equal(4, ignoredCsv.Errors[0].FirstRowNumber);

        Assert.Equal(2, trackedXlsx.Errors.Count);
        Assert.Equal(2, trackedCsv.Errors.Count);
        Assert.Equal(2, trackedXlsx.Workbook.Rows.Count);
        Assert.Equal(2, trackedCsv.Items.Count);
        Assert.Equal(new[] { 3, 5 }, trackedXlsx.Errors.Select(error => error.RowIndex));
        Assert.Equal(new[] { 3, 5 }, trackedCsv.Errors.Select(error => error.RowIndex));
        Assert.Equal(new int?[] { 2, 4 }, trackedXlsx.Errors.Select(error => error.FirstRowNumber));
        Assert.Equal(new int?[] { 2, 4 }, trackedCsv.Errors.Select(error => error.FirstRowNumber));
        Assert.All(trackedXlsx.Errors, error => Assert.Equal("unique", error.ColumnKey));
        Assert.All(trackedCsv.Errors, error => Assert.Equal("unique", error.PropertyName));

        Assert.Equal(1, limitedXlsx.Errors.Count);
        Assert.Equal(1, limitedCsv.Errors.Count);
        Assert.Equal(3, limitedXlsx.Errors[0].RowIndex);
        Assert.Equal(3, limitedCsv.Errors[0].RowIndex);
        Assert.Equal("unique", limitedXlsx.Errors[0].ColumnKey);
        Assert.Equal("unique", limitedCsv.Errors[0].PropertyName);
    }

    /// <summary>
    /// 测试 - 七类固定属性规则应在 CSV 与 XLSX 中保持相同 binding、坐标和失败行回滚契约。
    /// </summary>
    [Fact]
    public void Import_FixedBuiltInValidationMatrix_ShouldMatchCsvAndXlsx()
    {
        // Arrange
        var document = new ExcelMappingDocument
        {
            Import = new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    new() { PropertyName = nameof(FixedValidationRow.Required), Title = "必填" },
                    new() { PropertyName = nameof(FixedValidationRow.Regex), Title = "正则" },
                    new() { PropertyName = nameof(FixedValidationRow.Date), Title = "日期" },
                    new() { PropertyName = nameof(FixedValidationRow.MaxValue), Title = "最大值" },
                    new() { PropertyName = nameof(FixedValidationRow.Range), Title = "区间" },
                    new() { PropertyName = nameof(FixedValidationRow.MaxLength), Title = "长度" },
                    new() { PropertyName = nameof(FixedValidationRow.Unique), Title = "唯一" }
                }
            }
        };
        var headers = new[] { "必填", "正则", "日期", "最大值", "区间", "长度", "唯一" };
        var rows = new[]
        {
            new[] { "ok", "OK-1", "2026-08-24", "5", "5", "abc", "U-1" },
            new[] { "ok", "OK-1", "2026-08-24", "10", "1", "abcde", "U-2" },
            new[] { "ok", "OK-1", "2026-08-24", "10", "10", "abcde", "U-3" },
            new[] { "", "OK-1", "2026-08-24", "5", "5", "abc", "U-4" },
            new[] { "ok", "BAD", "2026-08-24", "5", "5", "abc", "U-5" },
            new[] { "ok", "OK-1", "not-date", "5", "5", "abc", "U-6" },
            new[] { "ok", "OK-1", "2026-08-24", "11", "5", "abc", "U-7" },
            new[] { "ok", "OK-1", "2026-08-24", "5", "11", "abc", "U-8" },
            new[] { "ok", "OK-1", "2026-08-24", "5", "5", "abcdef", "U-9" },
            new[] { "ok", "OK-1", "2026-08-24", "5", "5", "abc", "u-1" }
        };
        using var xlsx = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0);
            for (var column = 0; column < headers.Length; column++)
                sheet.GetRow(0).CreateCell(column).SetCellValue(headers[column]);
            for (var row = 0; row < rows.Length; row++)
            {
                sheet.CreateRow(row + 1);
                for (var column = 0; column < rows[row].Length; column++)
                    sheet.GetRow(row + 1).CreateCell(column).SetCellValue(rows[row][column]);
            }
        }));
        var request = ExcelImport.Workbook<FixedValidationWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet.Mapping(document)));
        using var csv = new MemoryStream(Encoding.UTF8.GetBytes(
            string.Join("\r\n", new[] { string.Join(",", headers) }
                .Concat(rows.Select(row => string.Join(",", row)))) + "\r\n"));

        // Act
        var xlsxResult = new NpoiExcelImporter().Import(xlsx, request);
        var csvResult = new CsvEntityImporter().Import<FixedValidationRow>(csv,
            new CsvImportOptions<FixedValidationRow> { MappingDocument = document, HeaderMatch = false });

        // Assert
        Assert.Equal(7, xlsxResult.Errors.Count);
        Assert.Equal(7, csvResult.Errors.Count);
        Assert.Equal(3, xlsxResult.Workbook.Rows.Count);
        Assert.Equal(3, csvResult.Items.Count);
        var expected = new[]
        {
            (Row: 5, Column: 1, Property: nameof(FixedValidationRow.Required), Code: ExcelImportErrorCode.Validation),
            (Row: 6, Column: 2, Property: nameof(FixedValidationRow.Regex), Code: ExcelImportErrorCode.Validation),
            (Row: 7, Column: 3, Property: nameof(FixedValidationRow.Date), Code: ExcelImportErrorCode.ValueConversion),
            (Row: 8, Column: 4, Property: nameof(FixedValidationRow.MaxValue), Code: ExcelImportErrorCode.MaxValue),
            (Row: 9, Column: 5, Property: nameof(FixedValidationRow.Range), Code: ExcelImportErrorCode.Validation),
            (Row: 10, Column: 6, Property: nameof(FixedValidationRow.MaxLength), Code: ExcelImportErrorCode.MaxLength),
            (Row: 11, Column: 7, Property: nameof(FixedValidationRow.Unique), Code: ExcelImportErrorCode.Validation)
        };
        Assert.Equal(expected.Select(item => $"{item.Code}|{item.Row}|{item.Column}|{item.Property}"),
            xlsxResult.Errors.Select(error => $"{error.Code}|{error.RowIndex}|{error.ColumnIndex}|{error.PropertyName}"));
        Assert.Equal(expected.Select(item => $"{item.Row}|{item.Column}|{item.Property}"),
            csvResult.Errors.Select(error => $"{error.RowIndex}|{error.ColumnIndex}|{error.PropertyName}"));
        Assert.Equal(2, xlsxResult.Errors.Single(error => error.PropertyName == nameof(FixedValidationRow.Unique)).FirstRowNumber);
        Assert.Equal(2, csvResult.Errors.Single(error => error.PropertyName == nameof(FixedValidationRow.Unique)).FirstRowNumber);
    }

    private static List<ExcelMappingDynamicColumnConfiguration> CreateDynamicValidationColumns() => new()
    {
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "required", Title = "必填", DataTypeName = "string", Order = 1,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "required" }
            }
        },
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "regex", Title = "正则", DataTypeName = "string", Order = 2,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "regex", Pattern = "^OK-" }
            }
        },
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "date", Title = "日期", DataTypeName = "datetime", Order = 3,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "date", Format = "yyyy-MM-dd" }
            }
        },
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "max", Title = "最大值", DataTypeName = "decimal", Order = 4,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "maxValue", MaxValue = 10 }
            }
        },
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "range", Title = "区间", DataTypeName = "decimal", Order = 5,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "range", Min = 1, Max = 10 }
            }
        },
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "length", Title = "长度", DataTypeName = "string", Order = 6,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "maxLength", MaxLength = 5 }
            }
        },
        new ExcelMappingDynamicColumnConfiguration
        {
            Key = "unique", Title = "唯一", DataTypeName = "string", Order = 7,
            ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
            {
                new() { Name = "unique", IgnoreEmpty = false }
            }
        }
    };

    private static ExcelMappingDocument CreateDynamicUniqueDocument(bool ignoreEmpty) => new()
    {
        Import = new ExcelMappingConfiguration
        {
            DynamicColumns =
            {
                new ExcelMappingDynamicColumnConfiguration
                {
                    Key = "unique", Title = "唯一", DataTypeName = "string",
                    ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
                    {
                        new() { Name = "unique", IgnoreEmpty = ignoreEmpty }
                    }
                }
            }
        }
    };

    private static ExcelWorkbookImportResult<DynamicTargetWorkbook> ImportDynamicUniqueXlsx(
        byte[] source, ExcelMappingDocument document, ExcelResourceLimits resourceLimits)
    {
        using var stream = new MemoryStream(source);
        var request = ExcelImport.Workbook<DynamicTargetWorkbook>(workbook => workbook
            .ResourceLimits(resourceLimits)
            .Sheet("Data", root => root.Rows, sheet => sheet.Mapping(document)));
        return new NpoiExcelImporter().Import(stream, request);
    }

    private static CsvImportResult<DynamicTargetRow> ImportDynamicUniqueCsv(string source,
        ExcelMappingDocument document, CsvImportOptions<DynamicTargetRow> options)
    {
        options.MappingDocument = document;
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));
        return new CsvEntityImporter().Import(stream, options);
    }

    /// <summary>
    /// 测试 - 按从零开始的索引选择 Sheet 时，应读取指定工作表而不是依赖名称。
    /// </summary>
    [Fact]
    public void Import_SheetSelectorByIndex_ShouldReadSelectedSheet()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            workbook.CreateSheet("忽略").CreateRow(0).CreateCell(0).SetCellValue("Code");
            var sheet = workbook.CreateSheet("目标");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("B-1");
        }));
        var request = ExcelImport.Workbook<SelectorWorkbook>(workbook =>
            workbook.Sheet(ExcelSheetSelector.ByIndex(1), root => root.Rows));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.Equal("B-1", Assert.Single(result.Workbook.Rows).Code);
        Assert.Equal("目标", Assert.Single(result.Sheets).Name);
    }

    /// <summary>
    /// 测试 - 重复索引和名称索引混用指向同一物理 Sheet 时，应在计划构建前给出确定性错误。
    /// </summary>
    [Fact]
    public void Import_DuplicateSheetSelectors_ShouldFailBeforePlanExecution()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("目标");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("A-1");
        }));

        var duplicateIndex = Record.Exception(() => ExcelImport.Workbook<SelectorWorkbook>(workbook =>
        {
            workbook.Sheet(ExcelSheetSelector.ByIndex(0), root => root.Rows);
            workbook.Sheet(ExcelSheetSelector.ByIndex(0), root => root.Rows);
        }));
        var mixedSelectorRequest = ExcelImport.Workbook<SelectorWorkbook>(workbook =>
        {
            workbook.Sheet(ExcelSheetSelector.ByName("目标"), root => root.Rows);
            workbook.Sheet(ExcelSheetSelector.ByIndex(0), root => root.Rows);
        });

        // Act
        source.Position = 0;
        var mixedSelector = Record.Exception(() => new NpoiExcelImporter().Import(source, mixedSelectorRequest));

        // Assert
        var duplicateIndexArgument = Assert.IsType<ArgumentException>(duplicateIndex);
        Assert.Contains("#0", duplicateIndexArgument.Message);
        var mixedSelectorArgument = Assert.IsType<ArgumentException>(mixedSelector);
        Assert.Contains("目标", mixedSelectorArgument.Message);
        Assert.Contains("#0", mixedSelectorArgument.Message);
        Assert.Contains("同一物理 Sheet", mixedSelectorArgument.Message);
    }

    /// <summary>
    /// 测试 - ReadColumns 应限制表头绑定范围，并允许范围外的辅助列存在。
    /// </summary>
    [Fact]
    public void Import_ReadColumns_ShouldIgnoreColumnsOutsideRange()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("辅助列");
            sheet.GetRow(0).CreateCell(1).SetCellValue("Code");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("ignored");
            sheet.GetRow(1).CreateCell(1).SetCellValue("A-1");
        }));
        var request = ExcelImport.Workbook<SelectorWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet.ReadColumns(1, 1)));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.Equal("A-1", Assert.Single(result.Workbook.Rows).Code);
    }

    /// <summary>
    /// 测试 - Header Trim 与 Body Preserve 应分别生效，正文原始首尾空白不得被 Reader 提前丢弃。
    /// </summary>
    [Fact]
    public void Import_WhitespacePolicies_ShouldApplyHeaderAndBodySeparately()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(" Code ");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("  A-1  ");
        }));
        var request = ExcelImport.Workbook<SelectorWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet
                .HeaderWhitespace(ExcelWhitespacePolicy.Trim)
                .BodyWhitespace(ExcelWhitespacePolicy.Preserve)));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.Equal("  A-1  ", Assert.Single(result.Workbook.Rows).Code);
    }

    /// <summary>
    /// 测试 - 动态列应写入请求指定的动态目标，而不是映射属性的默认 Setter。
    /// </summary>
    [Fact]
    public void Import_DynamicTargetGetter_ShouldWriteRequestedDictionary()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Extra");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("value");
        }));
        var definition = new ExcelDynamicColumnDefinition { Key = "extra", Title = "Extra" };
        var request = ExcelImport.Workbook<DynamicTargetWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet
                .DynamicColumns(row => row.TargetFields, new[] { definition })));

        // Act
        var result = new NpoiExcelImporter().Import(source, request);
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        var item = Assert.Single(result.Workbook.Rows);

        // Assert
        Assert.Equal("value", item.TargetFields["extra"]);
        Assert.Null(item.MappedFields);
    }

    /// <summary>
    /// 测试 - 动态列导入应使用定义中的命名双向转换器。
    /// </summary>
    [Fact]
    public void Import_DynamicColumnNamedConverter_ShouldConvertValue()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("地区");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("east");
        }));
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "地区",
            ConverterName = "upper-region"
        };
        var request = ExcelImport.Workbook<DynamicTargetWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet
                .DynamicColumns(row => row.TargetFields, new[] { definition })));

        // Act
        var result = new NpoiExcelImporter(valueConverters: new IExcelValueConverter[] { new UpperRegionConverter() })
            .Import(source, request);

        // Assert
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        Assert.Equal("EAST", Assert.Single(result.Workbook.Rows).TargetFields["region"]);
    }

    /// <summary>
    /// 测试 - 固定列转换器应在列计划创建时绑定，而不是按单元格重复解析。
    /// </summary>
    [Fact]
    public void Import_FixedColumnConverter_ShouldBindOnceBeforeCellConversion()
    {
        // Arrange
        var converter = new CountingConverter();
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("订单号");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("a-1");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("b-2");
        }));
        var configuration = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(ImportOrder.OrderNo), ConverterName = "counting" }
            }
        };
        var request = ExcelImport.Workbook<OrderWorkbook>(builder => builder
            .Sheet("Data", root => root.Orders, sheet => sheet.Mapping(configuration)));

        // Act
        var result = new NpoiExcelImporter(valueConverters: new IExcelValueConverter[] { converter })
            .Import(source, request);

        // Assert
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        Assert.Equal(2, result.Workbook.Orders.Count);
        Assert.Equal(1, converter.CanConvertCalls);
        Assert.Equal(2, converter.ConvertFromCalls);
    }

    /// <summary>
    /// 测试 - 动态列转换器应在列计划创建时绑定一次，并复用到每个数据行。
    /// </summary>
    [Fact]
    public void Import_DynamicColumnConverter_ShouldBindOnceBeforeCellConversion()
    {
        // Arrange
        var converter = new CountingConverter();
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("地区");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("east");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("west");
        }));
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "地区",
            DataType = typeof(string),
            ConverterName = "counting"
        };
        var request = ExcelImport.Workbook<DynamicTargetWorkbook>(builder =>
            builder.Sheet("Data", root => root.Rows, sheet => sheet
                .DynamicColumns(row => row.TargetFields, new[] { definition })));

        // Act
        var result = new NpoiExcelImporter(valueConverters: new IExcelValueConverter[] { converter })
            .Import(source, request);

        // Assert
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        Assert.Equal(2, result.Workbook.Rows.Count);
        Assert.Equal(1, converter.CanConvertCalls);
        Assert.Equal(2, converter.ConvertFromCalls);
    }

    /// <summary>
    /// 测试 - 导出固定列转换器应在列计划创建时绑定并复用到每个数据行。
    /// </summary>
    [Fact]
    public void Export_FixedColumnConverter_ShouldBindOnceBeforeCellConversion()
    {
        // Arrange
        var converter = new CountingConverter();
        var configuration = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(ExportCustomer.Name), ConverterName = "counting" }
            }
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", new[]
        {
            new ExportCustomer { Name = "a" },
            new ExportCustomer { Name = "b" }
        }, sheet => sheet.Mapping(configuration)));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter(new IExcelValueConverter[] { converter }).Export(request, destination);

        // Assert
        Assert.Equal(1, converter.CanConvertCalls);
        Assert.Equal(2, converter.ConvertToCalls);
    }

    /// <summary>
    /// 测试 - Fixed 列宽应按字符宽度写入工作表。
    /// </summary>
    [Fact]
    public void Export_ColumnWidthFixed_ShouldWriteConfiguredWidth()
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("客户",
            new[] { new ExportCustomer { Name = "客户 A" } }, sheet => sheet
                .ColumnWidth(new ExcelColumnWidthOptions
                {
                    Mode = ExcelColumnWidthMode.Fixed,
                    FixedWidth = 24
                })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        Assert.Equal(24 * 256, result.GetSheet("客户").GetColumnWidth(0));
    }

    /// <summary>
    /// 测试 - 自定义表头批注应写入传统 Note，并保留作者和可见性。
    /// </summary>
    [Fact]
    public void Export_HeaderComment_ShouldWriteLegacyNote()
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("客户",
            new[] { new ExportCustomer { Name = "客户 A" } }, sheet => sheet
                .HeaderRowIndex(1)
                .DataRowStartIndex(2)
                .HeaderRows(new[] { new ExcelHeaderRow(0,
                    new[] { new ExcelHeaderCell(1, "说明", comment: new ExcelComment("表头说明", "tester", true)) }) })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var result = WorkbookFactory.Create(destination);
        var comment = result.GetSheet("客户").GetRow(0).GetCell(1).CellComment;
        Assert.NotNull(comment);
        Assert.Equal("表头说明", comment.String.String);
        Assert.Equal("tester", comment.Author);
        Assert.True(comment.Visible);
    }

    /// <summary>
    /// 测试 - XLSX 应按列 Key 创建柱状、折线和饼图，并拒绝饼图多系列配置。
    /// </summary>
    [Fact]
    public void Export_XlsxCharts_ShouldCreateSupportedChartTypes()
    {
        // Arrange
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("统计",
            new[]
            {
                new ChartRow { Month = "一月", Amount = 10, Forecast = 12 },
                new ChartRow { Month = "二月", Amount = 15, Forecast = 16 }
            }, sheet => sheet
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
                })
                .Chart(new ExcelChartDefinition
                {
                    Title = "预测折线图",
                    Type = ExcelChartType.Line,
                    Categories = new ExcelChartRange { ColumnKey = "Month" },
                    Series = new[] { new ExcelChartSeries
                    {
                        Name = "预测", Values = new ExcelChartRange { ColumnKey = "Forecast" }
                    } },
                    Anchor = new ExcelChartAnchor { StartRow = 13, StartColumn = 5, EndRow = 25, EndColumn = 14 }
                })
                .Chart(new ExcelChartDefinition
                {
                    Title = "金额饼图",
                    Type = ExcelChartType.Pie,
                    Categories = new ExcelChartRange { ColumnKey = "Month" },
                    Series = new[] { new ExcelChartSeries
                    {
                        Name = "金额", Values = new ExcelChartRange { ColumnKey = "Amount" }
                    } },
                    Anchor = new ExcelChartAnchor { StartRow = 26, StartColumn = 5, EndRow = 38, EndColumn = 14 }
                })));
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        destination.Position = 0;
        using var workbook = (XSSFWorkbook)WorkbookFactory.Create(destination);
        var drawing = Assert.IsType<XSSFDrawing>(workbook.GetSheet("统计").CreateDrawingPatriarch());
        Assert.Equal(3, drawing.GetCharts().Count);
        Assert.Contains(drawing.GetCharts(), chart => chart.Title.String == "金额柱状图");
        Assert.Contains(drawing.GetCharts(), chart => chart.Title.String == "预测折线图");
        Assert.Contains(drawing.GetCharts(), chart => chart.Title.String == "金额饼图");
    }

    private static byte[] CreateWorkbook(Action<IWorkbook> configure)
    {
        using var workbook = new XSSFWorkbook();
        configure(workbook);
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    private static MetadataSnapshot CreateMetadataSnapshot(HSSFWorkbook workbook, string expected) => new(
        expected,
        workbook.SummaryInformation.Author,
        workbook.DocumentSummaryInformation.Company,
        workbook.SummaryInformation.Title,
        workbook.SummaryInformation.Subject,
        workbook.DocumentSummaryInformation.Category,
        workbook.SummaryInformation.Comments);

    private static byte[] CreateMetadataTemplate(ExcelFormat format, string author, string company, string title,
        string subject, string category, string description)
    {
        using IWorkbook workbook = format == ExcelFormat.Xls ? new HSSFWorkbook() : new XSSFWorkbook();
        if (workbook is XSSFWorkbook xssf)
        {
            var properties = xssf.GetProperties();
            properties.CoreProperties.Creator = author;
            properties.ExtendedProperties.GetUnderlyingProperties().Company = company;
            properties.CoreProperties.Title = title;
            properties.CoreProperties.Subject = subject;
            properties.CoreProperties.Category = category;
            properties.CoreProperties.Description = description;
        }
        else
        {
            var hssf = (HSSFWorkbook)workbook;
            var document = NPOI.HPSF.PropertySetFactory.CreateDocumentSummaryInformation();
            document.Company = company;
            document.Category = category;
            hssf.DocumentSummaryInformation = document;
            var summary = NPOI.HPSF.PropertySetFactory.CreateSummaryInformation();
            summary.Author = author;
            summary.Title = title;
            summary.Subject = subject;
            summary.Comments = description;
            hssf.SummaryInformation = summary;
        }
        workbook.CreateSheet("客户");
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    private static byte[] CreateStyledExportTemplate(ExcelFormat format)
    {
        using IWorkbook workbook = format == ExcelFormat.Xls
            ? new HSSFWorkbook()
            : new XSSFWorkbook();
        var orders = workbook.CreateSheet("订单");
        var customers = workbook.CreateSheet("客户");
        var orderStyle = CreateNonDefaultStyle(workbook);
        var customerStyle = CreateNonDefaultStyle(workbook);
        orders.CreateRow(0).CreateCell(0).CellStyle = orderStyle;
        orders.GetRow(0).CreateCell(1).CellStyle = orderStyle;
        orders.CreateRow(1).CreateCell(0).CellStyle = orderStyle;
        orders.GetRow(1).CreateCell(1).CellStyle = orderStyle;
        customers.CreateRow(0).CreateCell(0).CellStyle = customerStyle;
        customers.CreateRow(1).CreateCell(0).CellStyle = customerStyle;
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    private static ICellStyle CreateNonDefaultStyle(IWorkbook workbook)
    {
        var style = workbook.CreateCellStyle();
        var font = workbook.CreateFont();
        font.FontName = "Arial";
        font.FontHeightInPoints = 16;
        font.IsBold = true;
        font.IsItalic = true;
        font.Underline = FontUnderlineType.Single;
        font.Color = IndexedColors.Red.Index;
        style.SetFont(font);
        style.FillPattern = FillPattern.SolidForeground;
        style.FillForegroundColor = IndexedColors.Yellow.Index;
        style.FillBackgroundColor = IndexedColors.Green.Index;
        style.BorderTop = BorderStyle.Thin;
        style.BorderBottom = BorderStyle.Medium;
        style.BorderLeft = BorderStyle.Dashed;
        style.BorderRight = BorderStyle.Double;
        style.Alignment = HorizontalAlignment.Center;
        style.VerticalAlignment = VerticalAlignment.Center;
        style.WrapText = true;
        style.Indention = 4;
        style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");
        if (style is XSSFCellStyle xssfStyle)
        {
            xssfStyle.SetFillForegroundColor(new XSSFColor(new byte[] { 17, 34, 51 }, null));
            xssfStyle.SetFillBackgroundColor(new XSSFColor(new byte[] { 68, 85, 102 }, null));
        }
        return style;
    }

    private static void AssertResetStyle(ICellStyle style, ICell cell)
    {
        var font = cell.Sheet.Workbook.GetFontAt(style.FontIndex);
        var defaultFont = cell.Sheet.Workbook.GetFontAt(0);
        Assert.Equal(FillPattern.NoFill, style.FillPattern);
        Assert.Equal(BorderStyle.None, style.BorderTop);
        Assert.Equal(BorderStyle.None, style.BorderBottom);
        Assert.Equal(BorderStyle.None, style.BorderLeft);
        Assert.Equal(BorderStyle.None, style.BorderRight);
        Assert.Equal(HorizontalAlignment.General, style.Alignment);
        Assert.Equal(VerticalAlignment.Bottom, style.VerticalAlignment);
        Assert.False(style.WrapText);
        Assert.Equal(0, style.Indention);
        Assert.Equal("General", style.GetDataFormatString());
        Assert.Equal(defaultFont.FontName, font.FontName);
        Assert.Equal(defaultFont.FontHeightInPoints, font.FontHeightInPoints);
        Assert.Equal(defaultFont.Color, font.Color);
        Assert.Equal(defaultFont.IsBold, font.IsBold);
        Assert.Equal(defaultFont.IsItalic, font.IsItalic);
        Assert.Equal(defaultFont.Underline, font.Underline);
        if (style is XSSFCellStyle xssfStyle)
        {
            Assert.Null(xssfStyle.FillForegroundColorColor);
            Assert.Null(xssfStyle.FillBackgroundColorColor);
        }
        else
        {
            Assert.Equal(IndexedColors.Automatic.Index, style.FillForegroundColor);
            Assert.Equal(IndexedColors.Automatic.Index, style.FillBackgroundColor);
        }
    }

    private sealed class ExportOrder
    {
        public string OrderNo { get; set; }

        [DynamicColumn]
        public IDictionary<string, object> CustomFields { get; set; }
    }

    private sealed class ExportCustomer
    {
        public string Name { get; set; }
    }

    private sealed class MetadataSnapshot
    {
        public MetadataSnapshot(string expected, string author, string company, string title, string subject,
            string category, string description)
        {
            Expected = expected;
            Author = author;
            Company = company;
            Title = title;
            Subject = subject;
            Category = category;
            Description = description;
        }

        public string Expected { get; }
        public string Author { get; }
        public string Company { get; }
        public string Title { get; }
        public string Subject { get; }
        public string Category { get; }
        public string Description { get; }
    }

    private sealed class ThrowingExcelExporter : IExcelExporter
    {
        public void Export(ExcelWorkbookExportRequest request, Stream destination,
            CancellationToken cancellationToken = default)
        {
            destination.WriteByte(1);
            throw new InvalidOperationException("测试导出失败");
        }
    }

    private sealed class CancelingExcelExporter : IExcelExporter
    {
        public void Export(ExcelWorkbookExportRequest request, Stream destination,
            CancellationToken cancellationToken = default)
        {
            destination.WriteByte(1);
            throw new OperationCanceledException(cancellationToken);
        }
    }

    [Header(FontName = "Arial", FontSize = 11, Bold = false, Color = Color.Blue)]
    private sealed class HeaderStyleOrder
    {
        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string D { get; set; }
        public string E { get; set; }
        public string F { get; set; }
        public string G { get; set; }
        public string H { get; set; }
        public string I { get; set; }
        public string J { get; set; }
        public string K { get; set; }
        public string L { get; set; }
    }

    private sealed class ChartRow
    {
        public string Month { get; set; }
        public double Amount { get; set; }
        public double Forecast { get; set; }
    }

    private sealed class OrderWorkbook
    {
        public List<ImportOrder> Orders { get; } = new List<ImportOrder>();
        public List<ImportDetail> Details { get; } = new List<ImportDetail>();
    }

    private sealed class ImportOrder
    {
        [ColumnName("订单号")]
        public string OrderNo { get; set; }

        [DynamicColumn]
        public IDictionary<string, object> CustomFields { get; set; }

        [ExcelIgnore]
        public List<ImportDetail> DetailItems { get; } = new List<ImportDetail>();
    }

    private sealed class ImportDetail
    {
        [ColumnName("订单号")]
        public string OrderNo { get; set; }
        [ColumnName("名称")]
        public string Name { get; set; }
    }

    private sealed class IntRelationWorkbook
    {
        public List<IntParent> Parents { get; } = new();
        public List<IntChild> Children { get; } = new();
    }

    private sealed class IntParent
    {
        public int Id { get; set; }

        [ExcelIgnore]
        public List<IntChild> Children { get; } = new();
    }

    private sealed class IntChild
    {
        public int ParentId { get; set; }
    }

    private sealed class SelectorWorkbook
    {
        public List<SelectorRow> Rows { get; } = new List<SelectorRow>();
    }

    private sealed class SelectorRow
    {
        public string Code { get; set; }
    }

    private sealed class DynamicTargetWorkbook
    {
        public List<DynamicTargetRow> Rows { get; } = new List<DynamicTargetRow>();
    }

    private sealed class DynamicTargetRow
    {
        [DynamicColumn]
        public IDictionary<string, object> MappedFields { get; set; }

        [ExcelIgnore]
        public IDictionary<string, object> TargetFields { get; } =
            new Dictionary<string, object>(StringComparer.Ordinal);
    }

    private sealed class FixedValidationWorkbook
    {
        public List<FixedValidationRow> Rows { get; } = new();
    }

    private sealed class FixedValidationRow
    {
        [ExcelRequired]
        public string Required { get; set; }

        [ExcelRegex("^OK-")]
        public string Regex { get; set; }

        [ExcelDate("yyyy-MM-dd")]
        public DateTime Date { get; set; }

        [ExcelMaxValue(10)]
        public decimal MaxValue { get; set; }

        [ExcelRange(1, 10)]
        public decimal Range { get; set; }

        [ExcelMaxLength(5)]
        public string MaxLength { get; set; }

        [ExcelUnique(IgnoreEmpty = false)]
        public string Unique { get; set; }
    }

    private sealed class UpperRegionConverter : INamedExcelValueConverter
    {
        public string Name => "upper-region";

        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = Convert.ToString(context.Value)?.ToUpperInvariant();
            return true;
        }

        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = Convert.ToString(context.Value)?.ToUpperInvariant();
            return true;
        }
    }

    private sealed class CountingConverter : INamedExcelValueConverter
    {
        public string Name => "counting";

        public int CanConvertCalls { get; private set; }

        public int ConvertFromCalls { get; private set; }

        public int ConvertToCalls { get; private set; }

        public bool CanConvert(Type propertyType)
        {
            CanConvertCalls++;
            return propertyType == typeof(string);
        }

        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            ConvertFromCalls++;
            value = Convert.ToString(context.Value)?.ToUpperInvariant();
            return true;
        }

        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            ConvertToCalls++;
            value = context.Value;
            return true;
        }
    }

    private sealed class StartsWithOkValidationRule : INamedExcelValidationRule
    {
        public string Name => "starts-ok";
        public string ErrorMessage => "必须以 OK 开头";
        public bool Validate(ExcelValidationContext context) => context.Value.StartsWith("OK",
            StringComparison.Ordinal);
    }

    private sealed class ContainsRegionValidationRule : INamedExcelValidationRule
    {
        public string Name => "contains-region";
        public string ErrorMessage => "必须包含地区";
        public bool Validate(ExcelValidationContext context) => context.Value.Contains("地区",
            StringComparison.Ordinal);
    }

    private sealed class RecordingMappingPlanFactory : IExcelMappingPlanFactory
    {
        private readonly ExcelMappingPlanFactory _inner = new();
        private readonly List<int> _workbookPlanSheetCounts = new();

        public IReadOnlyList<int> WorkbookPlanSheetCounts => _workbookPlanSheetCounts;

        public IExcelMappingPlan Create<T>(ExcelMappingDocument document, MappingDirection direction)
            where T : class, new() => _inner.Create<T>(document, direction);

        public IExcelMappingPlan Create<T>(ExcelMappingDocument document,
            ExcelMappingConfiguration requestConfiguration, MappingDirection direction)
            where T : class, new() => _inner.Create<T>(document, requestConfiguration, direction);

        public IExcelMappingWorkbookPlan CreateWorkbook<T>(ExcelMappingDocument document,
            MappingDirection direction, IReadOnlyList<string> sheetNames) where T : class, new()
        {
            _workbookPlanSheetCounts.Add(sheetNames.Count);
            return _inner.CreateWorkbook<T>(document, direction, sheetNames);
        }

        public IExcelMappingWorkbookPlan CreateWorkbook<T>(ExcelMappingDocument document,
            ExcelMappingConfiguration requestConfiguration, MappingDirection direction,
            IReadOnlyList<string> sheetNames) where T : class, new()
        {
            _workbookPlanSheetCounts.Add(sheetNames.Count);
            return _inner.CreateWorkbook<T>(document, requestConfiguration, direction, sheetNames);
        }
    }
}
