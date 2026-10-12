using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证 Provider 原样保留透视模板的输出路径。
/// </summary>
public sealed class TemplatePivotPreserveExportTest
{
    /// <summary>
    /// 验证同步和异步保留模式复制完整模板部件。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task PreserveOnly_ShouldCopyCompletePivotTemplate(string provider, bool asynchronous)
    {
        var expected = CreatePivotTemplate();
        using var template = new MemoryStream(expected, writable: false);
        template.Position = 5;
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .PreserveTemplatePivotTables());
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        Assert.Equal(expected, output.ToArray());
        Assert.Equal(5, template.Position);
        Assert.True(template.CanRead);
        Assert.Contains("pivotTables/", output.ToArray().Length == 0
            ? string.Empty
            : string.Join("|", ReadEntries(output.ToArray())));
    }

    /// <summary>
    /// 验证写入模板业务 Sheet 后仍保留完整透视部件。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task PreserveWithBusinessSheet_ShouldKeepPivotParts(string provider, bool asynchronous)
    {
        var templateBytes = CreatePivotTemplate();
        using var template = new MemoryStream(templateBytes, writable: false);
        template.Position = 9;
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", new[]
            {
                new TemplatePivotRow { Region = "West", Amount = 99 },
                new TemplatePivotRow { Region = "East", Amount = 101 }
            })
            .PreserveTemplatePivotTables());
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        var actual = output.ToArray();
        var entries = ReadEntries(actual);
        Assert.Contains(entries, name => name.Contains("pivotTables/", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(entries, name => name.Contains("pivotCacheDefinition", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(entries, name => name.Contains("pivotCacheRecords", StringComparison.OrdinalIgnoreCase));
        using var workbook = new XLWorkbook(new MemoryStream(actual));
        var values = workbook.Worksheet("Data").Range("A1:B3").Cells().Select(cell => cell.GetString()).ToArray();
        Assert.Contains("West", values);
        Assert.Contains("99", values);
        Assert.Equal(9, template.Position);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证透视部件合并完成后清理临时归档文件。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task PreserveWithBusinessSheet_ShouldCleanArchiveTemporaryFile(string provider,
        bool asynchronous)
    {
        var existingArchive = GetArchiveTemporaryFiles();
        var existingParts = GetPartsTemporaryFiles();
        var existingStripped = GetStrippedTemporaryFiles();
        var existingGenerated = GetGeneratedTemporaryFiles();
        var existingAsync = GetAsyncTemporaryFiles();
        using var template = new MemoryStream(CreatePivotTemplate(), writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", new[] { new TemplatePivotRow { Region = "West", Amount = 99 } })
            .PreserveTemplatePivotTables());
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        Assert.Equal(existingArchive, GetArchiveTemporaryFiles());
        Assert.Equal(existingParts, GetPartsTemporaryFiles());
        Assert.Equal(existingStripped, GetStrippedTemporaryFiles());
        Assert.Equal(existingGenerated, GetGeneratedTemporaryFiles());
        Assert.Equal(existingAsync, GetAsyncTemporaryFiles());
    }

    /// <summary>
    /// 验证预取消不会写入目标流或改变模板位置。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task PreserveMode_PreCanceled_ShouldKeepStreams(string provider, bool asynchronous)
    {
        var templateBytes = CreatePivotTemplate();
        using var template = new TemplatePivotTrackingStream(templateBytes);
        template.Position = 9;
        var originalTemplate = template.ToArray();
        using var output = new TemplatePivotTrackingStream(new byte[] { 4, 3, 2, 1 });
        output.Position = 2;
        var originalOutput = output.ToArray();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", new[] { new TemplatePivotRow { Region = "West", Amount = 99 } })
            .PreserveTemplatePivotTables());

        var error = asynchronous
            ? await Assert.ThrowsAsync<OperationCanceledException>(() =>
                ExportAsync(provider, request, output, cancellation.Token)).ConfigureAwait(false)
            : Assert.Throws<OperationCanceledException>(() =>
                Export(provider, request, output, cancellation.Token));

        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(9, template.Position);
        Assert.False(template.WasDisposed);
        Assert.Equal(originalOutput, output.ToArray());
        Assert.Equal(2, output.Position);
        Assert.False(output.WasDisposed);
    }

    /// <summary>
    /// 验证公开导出资源限制在业务数据枚举前约束模板快照。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task ResourceLimits_ShouldRejectTemplateBeforeBusinessEnumeration(string provider,
        bool asynchronous)
    {
        var templateBytes = CreatePivotTemplate();
        using var template = new TemplatePivotTrackingStream(templateBytes);
        template.Position = 6;
        var originalTemplate = template.ToArray();
        var enumerated = false;
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", EnumerateRows(() => enumerated = true))
            .PreserveTemplatePivotTables()
            .ResourceLimits(new ExcelResourceLimits { MaxInputBytes = templateBytes.Length - 1 }));
        using var output = new TemplatePivotTrackingStream(new byte[] { 8, 7, 6 });
        output.Position = 1;
        var originalOutput = output.ToArray();

        var error = asynchronous
            ? await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
                ExportAsync(provider, request, output)).ConfigureAwait(false)
            : Assert.Throws<BingOfficesResourceLimitException>(() =>
                Export(provider, request, output));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, error.Code);
        Assert.False(enumerated);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(6, template.Position);
        Assert.False(template.WasDisposed);
        Assert.Equal(originalOutput, output.ToArray());
        Assert.Equal(1, output.Position);
        Assert.False(output.WasDisposed);
    }

    /// <summary>
    /// 验证仅保留模板的路径也使用公开导出资源限制。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task ResourceLimits_ShouldRejectPreserveOnlyTemplate(string provider, bool asynchronous)
    {
        var templateBytes = CreatePivotTemplate();
        using var template = new TemplatePivotTrackingStream(templateBytes);
        template.Position = 4;
        var originalTemplate = template.ToArray();
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .PreserveTemplatePivotTables()
            .ResourceLimits(new ExcelResourceLimits { MaxInputBytes = templateBytes.Length - 1 }));
        using var output = new TemplatePivotTrackingStream(new byte[] { 5, 4, 3 });
        output.Position = 2;
        var originalOutput = output.ToArray();

        var error = asynchronous
            ? await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
                ExportAsync(provider, request, output)).ConfigureAwait(false)
            : Assert.Throws<BingOfficesResourceLimitException>(() =>
                Export(provider, request, output));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, error.Code);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(4, template.Position);
        Assert.False(template.WasDisposed);
        Assert.Equal(originalOutput, output.ToArray());
        Assert.Equal(2, output.Position);
        Assert.False(output.WasDisposed);
    }

    /// <summary>
    /// 验证不添加业务工作表时仍会执行模板透视更新。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdateOnly_ShouldRebuildTemplatePivotWithoutBusinessSheet(string provider,
        bool asynchronous)
    {
        var templateBytes = CreatePivotTemplate();
        using var template = new MemoryStream(templateBytes, writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .UpdatePivotTable(new ExcelPivotTableDefinition
            {
                Name = "SalesByRegion",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition { EndRow = 3, EndColumn = 1 },
                TargetSheet = "Summary",
                TargetRow = 1,
                TargetColumn = 2,
                RowField = "Region",
                ValueField = "Amount"
            }));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        var actual = output.ToArray();
        Assert.False(templateBytes.SequenceEqual(actual));
        var cacheXml = ReadPivotCacheXml(actual);
        Assert.Contains("10", cacheXml, StringComparison.Ordinal);
        Assert.Contains("20", cacheXml, StringComparison.Ordinal);
        Assert.Contains("30", cacheXml, StringComparison.Ordinal);
        Assert.Equal(3, ReadPivotCacheRecordRows(actual));
        Assert.Contains(true, ReadPivotCacheRefreshFlags(actual));
        using var workbook = new XLWorkbook(new MemoryStream(actual));
        Assert.Equal("SalesByRegion", Assert.Single(workbook.Worksheet("Summary").PivotTables).Name);
        Assert.DoesNotContain("__BingPivotUpdate_001", workbook.Worksheets.Select(sheet => sheet.Name));
    }

    /// <summary>
    /// 验证更新透视时保留未配置的布局属性和样式。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task Update_ShouldPreserveUnconfiguredPivotLayout(string provider,
        bool asynchronous)
    {
        using var template = new MemoryStream(WithPivotLayout(CreatePivotTemplate(),
            "SalesByRegion"), writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .UpdatePivotTable(CreateDefinition("SalesByRegion", "Summary", "Data")));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        using var archive = new ZipArchive(new MemoryStream(output.ToArray(), writable: false));
        var pivot = archive.Entries.Single(entry => entry.FullName.Contains("pivotTables/",
            StringComparison.OrdinalIgnoreCase) && entry.FullName.EndsWith(".xml",
            StringComparison.OrdinalIgnoreCase) && string.Equals((string)ReadXml(archive,
                entry.FullName).Root.Attribute("name"), "SalesByRegion", StringComparison.Ordinal));
        var definition = ReadXml(archive, pivot.FullName).Root;
        Assert.Equal("0", (string)definition.Attribute("compact"));
        Assert.Equal("0", (string)definition.Attribute("showEmptyCol"));
        Assert.Equal("0", (string)definition.Attribute("showDataTips"));
        Assert.Equal("0", (string)definition.Attribute("showHeaders"));
        Assert.Equal("0", (string)definition.Attribute("rowGrandTotals"));
        var style = definition.Elements().Single(item => item.Name.LocalName == "pivotTableStyleInfo");
        Assert.Equal("PivotStyleDark4", (string)style.Attribute("name"));
        Assert.Equal("0", (string)style.Attribute("showRowHeaders"));
        Assert.Equal("1", (string)style.Attribute("showRowStripes"));
        var format = definition.Elements().Single(item => item.Name.LocalName == "formats")
            .Elements().Single(item => item.Name.LocalName == "format");
        Assert.Equal("formatting", (string)format.Attribute("action"));
        Assert.Equal("0", (string)format.Elements().Single(item =>
            item.Name.LocalName == "pivotArea").Attribute("dataOnly"));
    }

    /// <summary>
    /// 验证未知透视扩展导致更新失败且不改写目标流。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task Update_UnknownExtension_ShouldRejectWithoutChangingDestination(string provider,
        bool asynchronous)
    {
        using var template = new TemplatePivotTrackingStream(
            WithPivotExtension(CreatePivotTemplate(), "SalesByRegion"));
        template.Position = 7;
        var originalTemplate = template.ToArray();
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .UpdatePivotTable(CreateDefinition("SalesByRegion", "Summary", "Data")));
        using var output = new MemoryStream(new byte[] { 9, 8, 7 });
        output.Position = 1;
        var originalOutput = output.ToArray();

        var error = asynchronous
            ? await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
                ExportAsync(provider, request, output)).ConfigureAwait(false)
            : Assert.Throws<BingOfficesConfigurationException>(() =>
                Export(provider, request, output));

        Assert.Equal(BingOfficesStage.Preflight, error.Stage);
        Assert.Equal(originalOutput, output.ToArray());
        Assert.Equal(1, output.Position);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(7, template.Position);
        Assert.True(output.CanWrite);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证采购与仓储模板中的多轴、多聚合透视可共同更新。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdateBusinessTemplates_ShouldKeepMeasuresFiltersAndFormats(string provider,
        bool asynchronous)
    {
        using var template = new TemplatePivotTrackingStream(CreateBusinessPivotTemplate());
        template.Position = 11;
        var originalTemplate = template.ToArray();
        var purchaseDefinition = CreateBusinessDefinition("PurchaseByRegion", "PurchaseData",
            "PurchaseSummary", includeMultipleMeasures: true);
        var warehouseDefinition = CreateBusinessDefinition("WarehouseByRegion", "WarehouseData",
            "WarehouseSummary", includeMultipleMeasures: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("PurchaseData", CreateBusinessRows(110))
            .AddSheet("WarehouseData", CreateBusinessRows(210))
            .UpdatePivotTable(purchaseDefinition)
            .UpdatePivotTable(warehouseDefinition));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        var bytes = output.ToArray();
        var cacheXml = ReadPivotCacheXml(bytes);
        Assert.Contains("110.25", cacheXml, StringComparison.Ordinal);
        Assert.Contains("120.5", cacheXml, StringComparison.Ordinal);
        Assert.Contains("210.25", cacheXml, StringComparison.Ordinal);
        Assert.Contains("220.5", cacheXml, StringComparison.Ordinal);
        Assert.DoesNotContain("v=\"10.25\"", cacheXml, StringComparison.Ordinal);
        Assert.DoesNotContain("v=\"20.5\"", cacheXml, StringComparison.Ordinal);
        Assert.DoesNotContain("v=\"30.75\"", cacheXml, StringComparison.Ordinal);
        Assert.DoesNotContain("v=\"40.25\"", cacheXml, StringComparison.Ordinal);
        Assert.Equal(2, CountPivotCacheDefinitions(bytes));
        Assert.Equal(2, ReadPivotCacheRefreshFlags(bytes).Count(refresh => refresh));

        var purchasePivot = ReadPivotTableXml(bytes, "PurchaseByRegion").Root;
        XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var purchaseValues = purchasePivot.Element(main + "dataFields").Elements().ToArray();
        Assert.Equal(new[] { "sum", "count", "average", "min", "max" },
            purchaseValues.Select(value => (string)value.Attribute("subtotal") ?? "sum"));
        Assert.Equal(2, purchasePivot.Element(main + "rowFields").Elements().Count());
        var purchaseColumnFields = purchasePivot.Element(main + "colFields").Elements().ToArray();
        Assert.Equal(2, purchaseColumnFields.Length);
        Assert.Equal(-2, (int)purchaseColumnFields[1].Attribute("x"));
        Assert.Single(purchasePivot.Element(main + "pageFields").Elements());
        Assert.True((uint)purchaseValues[0].Attribute("numFmtId") > 0);

        var warehousePivot = ReadPivotTableXml(bytes, "WarehouseByRegion").Root;
        Assert.Single(warehousePivot.Element(main + "dataFields").Elements());
        Assert.Single(warehousePivot.Element(main + "pageFields").Elements());
        Assert.True((uint)warehousePivot.Element(main + "dataFields").Elements().Single()
            .Attribute("numFmtId") > 0);
        Assert.Contains("#,##0.000", ReadCustomNumberFormats(bytes));

        using var workbook = new XLWorkbook(new MemoryStream(bytes));
        Assert.Equal("PurchaseByRegion", Assert.Single(workbook.Worksheet("PurchaseSummary")
            .PivotTables).Name);
        Assert.Equal("WarehouseByRegion", Assert.Single(workbook.Worksheet("WarehouseSummary")
            .PivotTables).Name);
        Assert.DoesNotContain("__BingPivotUpdate_", string.Join("|",
            workbook.Worksheets.Select(sheet => sheet.Name)), StringComparison.Ordinal);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(11, template.Position);
    }

    /// <summary>
    /// 验证同一模板更新请求可重复成功导出。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdateRequest_ShouldBeReusableForRepeatedExports(string provider,
        bool asynchronous)
    {
        using var template = new TemplatePivotTrackingStream(CreatePivotTemplate());
        template.Position = 7;
        var originalTemplate = template.ToArray();
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", new[]
            {
                new TemplatePivotRow { Region = "West", Amount = 99 },
                new TemplatePivotRow { Region = "East", Amount = 101 }
            })
            .UpdatePivotTable(new ExcelPivotTableDefinition
            {
                Name = "SalesByRegion",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition { EndRow = 2, EndColumn = 1 },
                TargetSheet = "Summary",
                TargetRow = 1,
                TargetColumn = 2,
                RowField = "Region",
                ValueField = "Amount"
            }));

        for (var iteration = 0; iteration < 2; iteration++)
        {
            using var output = new MemoryStream();
            if (asynchronous)
                await ExportAsync(provider, request, output).ConfigureAwait(false);
            else
                Export(provider, request, output);

            var bytes = output.ToArray();
            Assert.Equal(2, ReadPivotCacheRecordRows(bytes));
            Assert.Contains("99", ReadPivotCacheXml(bytes), StringComparison.Ordinal);
            Assert.Contains("101", ReadPivotCacheXml(bytes), StringComparison.Ordinal);
            using var workbook = new XLWorkbook(new MemoryStream(bytes));
            Assert.Equal("SalesByRegion", Assert.Single(workbook.Worksheet("Summary")
                .PivotTables).Name);
            Assert.DoesNotContain("__BingPivotUpdate_", string.Join("|",
                workbook.Worksheets.Select(sheet => sheet.Name)), StringComparison.Ordinal);
        }

        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(7, template.Position);
    }

    /// <summary>
    /// 验证来源行数增加后透视缓存纳入新增记录。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdateSourceGrowth_ShouldIncludeAllNewRows(string provider,
        bool asynchronous)
    {
        using var template = new MemoryStream(CreatePivotTemplate(), writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", new[]
            {
                new TemplatePivotRow { Region = "North", Amount = 101 },
                new TemplatePivotRow { Region = "South", Amount = 202 },
                new TemplatePivotRow { Region = "West", Amount = 303 },
                new TemplatePivotRow { Region = "East", Amount = 404 }
            })
            .UpdatePivotTable(new ExcelPivotTableDefinition
            {
                Name = "SalesByRegion",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition { EndRow = 4, EndColumn = 1 },
                TargetSheet = "Summary",
                TargetRow = 1,
                TargetColumn = 2,
                RowField = "Region",
                ValueField = "Amount"
            }));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        var bytes = output.ToArray();
        var cacheXml = ReadPivotCacheXml(bytes);
        Assert.Equal(4, ReadPivotCacheRecordRows(bytes));
        Assert.Contains("101", cacheXml, StringComparison.Ordinal);
        Assert.Contains("202", cacheXml, StringComparison.Ordinal);
        Assert.Contains("303", cacheXml, StringComparison.Ordinal);
        Assert.Contains("404", cacheXml, StringComparison.Ordinal);
        using var workbook = new XLWorkbook(new MemoryStream(bytes));
        Assert.Equal("SalesByRegion", Assert.Single(workbook.Worksheet("Summary")
            .PivotTables).Name);
    }

    /// <summary>
    /// 验证单个无共享缓存透视可以按新来源数据更新。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdateSinglePivot_ShouldWriteNewCacheRecords(string provider, bool asynchronous)
    {
        using var template = new MemoryStream(CreatePivotTemplate(), writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", new[]
            {
                new TemplatePivotRow { Region = "West", Amount = 99 },
                new TemplatePivotRow { Region = "East", Amount = 101 }
            })
            .UpdatePivotTable(new ExcelPivotTableDefinition
            {
                Name = "SalesByRegion",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition { EndRow = 2, EndColumn = 1 },
                TargetSheet = "Summary",
                TargetRow = 1,
                TargetColumn = 2,
                RowField = "Region",
                ValueField = "Amount"
            }));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        var cacheXml = ReadPivotCacheXml(output.ToArray());
        Assert.Contains("99", cacheXml, StringComparison.Ordinal);
        Assert.Contains("101", cacheXml, StringComparison.Ordinal);
        Assert.Equal(2, ReadPivotCacheRecordRows(output.ToArray()));
        Assert.Contains(true, ReadPivotCacheRefreshFlags(output.ToArray()));
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal("SalesByRegion", Assert.Single(workbook.Worksheet("Summary").PivotTables).Name);
        Assert.DoesNotContain("__BingPivotUpdate_001", workbook.Worksheets.Select(sheet => sheet.Name));
    }

    /// <summary>
    /// 验证来源范围收缩后清除旧范围中的尾部数据。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdateSourceShrink_ShouldClearOldManagedRows(string provider,
        bool asynchronous)
    {
        var templateBytes = WithWorksheetValues(CreatePivotTemplate(), "Data",
            ("A3", "Stale row 2"), ("B3", "200"),
            ("A4", "Stale row 3"), ("B4", "300"), ("C4", "Outside source"));
        using var template = new MemoryStream(templateBytes, writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .UpdatePivotTable(new ExcelPivotTableDefinition
            {
                Name = "SalesByRegion",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition { EndRow = 1, EndColumn = 1 },
                TargetSheet = "Summary",
                TargetRow = 1,
                TargetColumn = 2,
                RowField = "Region",
                ValueField = "Amount"
            }));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        var data = workbook.Worksheet("Data");
        Assert.Equal("North", data.Cell("A2").GetString());
        Assert.False(WorksheetCellHasValue(output.ToArray(), "Data", "A3"));
        Assert.False(WorksheetCellHasValue(output.ToArray(), "Data", "B3"));
        Assert.False(WorksheetCellHasValue(output.ToArray(), "Data", "A4"));
        Assert.False(WorksheetCellHasValue(output.ToArray(), "Data", "B4"));
        Assert.Equal("Outside source", data.Cell("C4").GetString());
        Assert.Equal(1, ReadPivotCacheRecordRows(output.ToArray()));
        Assert.Contains("10", ReadPivotCacheXml(output.ToArray()), StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证更新会清除旧透视显示值并保留区域外单元格。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task Update_ShouldClearOldPivotBodyAndKeepOutsideCells(string provider,
        bool asynchronous)
    {
        var rangedTemplate = WithPivotTargetReference(CreatePivotTemplate(), "SalesByRegion", "C2:D4");
        using var template = new MemoryStream(WithWorksheetValues(rangedTemplate, "Summary",
            ("C2", "Old pivot label"), ("D2", "Old pivot value"),
            ("D4", "Old trailing value"), ("Z20", "Keep me")),
            writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .UpdatePivotTable(new ExcelPivotTableDefinition
            {
                Name = "SalesByRegion",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition { EndRow = 3, EndColumn = 1 },
                TargetSheet = "Summary",
                TargetRow = 1,
                TargetColumn = 2,
                RowField = "Region",
                ValueField = "Amount"
            }));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        var summary = workbook.Worksheet("Summary");
        Assert.False(WorksheetCellHasValue(output.ToArray(), "Summary", "C2"));
        Assert.False(WorksheetCellHasValue(output.ToArray(), "Summary", "D2"));
        Assert.False(WorksheetCellHasValue(output.ToArray(), "Summary", "D4"));
        Assert.Equal("Keep me", summary.Cell("Z20").GetString());
        Assert.Equal("SalesByRegion", Assert.Single(summary.PivotTables).Name);
    }

    /// <summary>
    /// 验证新透视区域覆盖模板非空单元格时拒绝更新且不写入目标流。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task Update_ShouldRejectOccupiedTargetAndKeepDestination(string provider,
        bool asynchronous)
    {
        using var template = new MemoryStream(WithWorksheetValues(CreatePivotTemplate(), "Summary",
            ("B1", "Do not overwrite")), writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .UpdatePivotTable(new ExcelPivotTableDefinition
            {
                Name = "SalesByRegion",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition { EndRow = 3, EndColumn = 1 },
                TargetSheet = "Summary",
                TargetRow = 0,
                TargetColumn = 1,
                RowField = "Region",
                ValueField = "Amount"
            }));
        using var output = new MemoryStream();
        output.Write(new byte[] { 9, 8, 7 });
        output.Position = 1;
        var original = output.ToArray();

        var error = asynchronous
            ? await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
                ExportAsync(provider, request, output)).ConfigureAwait(false)
            : Assert.Throws<BingOfficesConfigurationException>(() =>
                Export(provider, request, output));

        Assert.Equal(BingOfficesStage.Preflight, error.Stage);
        Assert.Equal(original, output.ToArray());
        Assert.Equal(1, output.Position);
    }

    /// <summary>
    /// 验证目标流写入失败后可复用同一更新请求重试。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdateDestinationFailure_ShouldAllowRetry(string provider, bool asynchronous)
    {
        using var template = new TemplatePivotTrackingStream(CreatePivotTemplate());
        template.Position = 7;
        var originalTemplate = template.ToArray();
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", new[]
            {
                new TemplatePivotRow { Region = "West", Amount = 99 },
                new TemplatePivotRow { Region = "East", Amount = 101 }
            })
            .UpdatePivotTable(new ExcelPivotTableDefinition
            {
                Name = "SalesByRegion",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition { EndRow = 2, EndColumn = 1 },
                TargetSheet = "Summary",
                TargetRow = 1,
                TargetColumn = 2,
                RowField = "Region",
                ValueField = "Amount"
            }));
        using var failedOutput = new FailOnceWriteStream();

        var error = asynchronous
            ? await Record.ExceptionAsync(() => ExportAsync(provider, request, failedOutput))
                .ConfigureAwait(false)
            : Record.Exception(() => Export(provider, request, failedOutput));

        Assert.NotNull(error);
        Assert.True(ContainsIOException(error));
        Assert.True(failedOutput.CanWrite);
        Assert.Empty(failedOutput.ToArray());
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(7, template.Position);

        using var retryOutput = new MemoryStream();
        if (asynchronous)
            await ExportAsync(provider, request, retryOutput).ConfigureAwait(false);
        else
            Export(provider, request, retryOutput);

        var retryBytes = retryOutput.ToArray();
        Assert.Contains("99", ReadPivotCacheXml(retryBytes), StringComparison.Ordinal);
        Assert.Contains("101", ReadPivotCacheXml(retryBytes), StringComparison.Ordinal);
        Assert.Contains(true, ReadPivotCacheRefreshFlags(retryBytes));
        using var workbook = new XLWorkbook(new MemoryStream(retryBytes));
        Assert.Equal("SalesByRegion", Assert.Single(workbook.Worksheet("Summary").PivotTables).Name);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(7, template.Position);
    }

    /// <summary>
    /// 验证多个独立缓存透视可以同时更新。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdateMultiplePivots_ShouldWriteAllIndependentCaches(string provider, bool asynchronous)
    {
        using var template = new TemplatePivotTrackingStream(CreateTwoPivotTemplate());
        template.Position = 7;
        var originalTemplate = template.ToArray();
        var enumerated = false;
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", EnumerateRows(() => enumerated = true))
            .AddSheet("Data2", EnumerateRows(() => enumerated = true))
            .UpdatePivotTable(CreateDefinition("SalesByRegion", "Summary", "Data"))
            .UpdatePivotTable(CreateDefinition("SalesByRegion2", "Summary2", "Data2")));
        using var output = new MemoryStream();
        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        Assert.True(enumerated);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(7, template.Position);
        var cacheXml = ReadPivotCacheXml(output.ToArray());
        Assert.Contains("99", cacheXml, StringComparison.Ordinal);
        Assert.Contains("101", cacheXml, StringComparison.Ordinal);
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal("SalesByRegion", Assert.Single(workbook.Worksheet("Summary").PivotTables).Name);
        Assert.Equal("SalesByRegion2", Assert.Single(workbook.Worksheet("Summary2").PivotTables).Name);
        Assert.DoesNotContain("__BingPivotUpdate_", string.Join("|",
            workbook.Worksheets.Select(sheet => sheet.Name)), StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证同一请求可以新增透视并更新模板中的全部已有透视。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdateAndCreate_ShouldWriteNewAndUpdatedPivots(string provider,
        bool asynchronous)
    {
        using var template = new MemoryStream(CreatePivotTemplate(), writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", EnumerateRows(() => { }))
            .UpdatePivotTable(CreateDefinition("SalesByRegion", "Summary", "Data"))
            .PivotTable(CreateDefinition("NewByRegion", "NewSummary", "Data")));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal("SalesByRegion", Assert.Single(workbook.Worksheet("Summary").PivotTables).Name);
        Assert.Equal("NewByRegion", Assert.Single(workbook.Worksheet("NewSummary").PivotTables).Name);
        Assert.DoesNotContain("__BingPivotUpdate_", string.Join("|",
            workbook.Worksheets.Select(sheet => sheet.Name)), StringComparison.Ordinal);
        var cacheXml = ReadPivotCacheXml(output.ToArray());
        Assert.Contains("99", cacheXml, StringComparison.Ordinal);
        Assert.Contains("101", cacheXml, StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证组合新增和更新时继续保留未更新的模板透视。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdateAndCreatePartial_ShouldKeepUnchangedTemplatePivot(string provider,
        bool asynchronous)
    {
        using var template = new MemoryStream(CreateTwoPivotTemplate(), writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", EnumerateRows(() => { }))
            .UpdatePivotTable(CreateDefinition("SalesByRegion", "Summary", "Data"))
            .PivotTable(CreateDefinition("NewByRegion", "NewSummary", "Data")));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal("SalesByRegion", Assert.Single(workbook.Worksheet("Summary").PivotTables).Name);
        Assert.Equal("SalesByRegion2", Assert.Single(workbook.Worksheet("Summary2").PivotTables).Name);
        Assert.Equal("NewByRegion", Assert.Single(workbook.Worksheet("NewSummary").PivotTables).Name);
        Assert.DoesNotContain("__BingPivotUpdate_", string.Join("|",
            workbook.Worksheets.Select(sheet => sheet.Name)), StringComparison.Ordinal);
        var cacheXml = ReadPivotCacheXml(output.ToArray());
        Assert.Contains("10", cacheXml, StringComparison.Ordinal);
        Assert.Contains("20", cacheXml, StringComparison.Ordinal);
        Assert.Contains("30", cacheXml, StringComparison.Ordinal);
        Assert.Contains("99", cacheXml, StringComparison.Ordinal);
        Assert.Contains("101", cacheXml, StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证仅保留模板并新建透视时继续保留原透视部件。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task PreserveAndCreate_ShouldKeepExistingPivotAndCreateNewPivot(string provider,
        bool asynchronous)
    {
        using var template = new TemplatePivotTrackingStream(CreatePivotTemplate());
        template.Position = 7;
        var originalTemplate = template.ToArray();
        var originalPivotParts = ReadPivotPackageParts(originalTemplate);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", new[]
            {
                new TemplatePivotRow { Region = "North", Amount = 10 },
                new TemplatePivotRow { Region = "South", Amount = 20 },
                new TemplatePivotRow { Region = "North", Amount = 30 }
            })
            .PreserveTemplatePivotTables()
            .PivotTable(CreateDefinition("NewByRegion", "NewSummary")));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        var outputBytes = output.ToArray();
        var outputPivotParts = ReadPivotPackageParts(outputBytes);
        foreach (var originalPart in originalPivotParts)
        {
            Assert.True(outputPivotParts.TryGetValue(originalPart.Key, out var preservedPart),
                $"缺少原始透视部件: {originalPart.Key}");
            Assert.Equal(originalPart.Value, preservedPart);
        }

        Assert.Equal(2, CountPivotCacheDefinitions(outputBytes));
        using (var archive = new ZipArchive(new MemoryStream(outputBytes), ZipArchiveMode.Read))
        using (var workbookStream = archive.GetEntry("xl/workbook.xml")?.Open()
            ?? throw new InvalidDataException("输出缺少 workbook.xml。"))
        {
            var workbookXml = XDocument.Load(workbookStream);
            Assert.Single(workbookXml.Root.Elements().Where(item => item.Name.LocalName == "pivotCaches"));
            var cacheIds = workbookXml.Descendants().Where(item => item.Name.LocalName == "pivotCache")
                .Select(item => (string)item.Attribute("cacheId")).ToArray();
            Assert.Equal(2, cacheIds.Length);
            Assert.Equal(cacheIds.Length, cacheIds.Distinct(StringComparer.Ordinal).Count());
        }

        var cacheXml = ReadPivotCacheXml(outputBytes);
        Assert.Contains("10", cacheXml, StringComparison.Ordinal);
        Assert.Contains("20", cacheXml, StringComparison.Ordinal);
        Assert.Contains("30", cacheXml, StringComparison.Ordinal);
        using var workbook = new XLWorkbook(new MemoryStream(outputBytes));
        Assert.Equal("SalesByRegion", Assert.Single(workbook.Worksheet("Summary").PivotTables).Name);
        Assert.Equal("NewByRegion", Assert.Single(workbook.Worksheet("NewSummary").PivotTables).Name);
        Assert.DoesNotContain("__BingPivotUpdate_", string.Join("|",
            workbook.Worksheets.Select(sheet => sheet.Name)), StringComparison.Ordinal);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(7, template.Position);
        Assert.False(template.WasDisposed);
    }

    /// <summary>
    /// 验证仅保留模板并新建透视时名称或目标冲突会在写入前拒绝。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    /// <param name="conflictName">是否使用已有透视名称。</param>
    [Theory]
    [InlineData("NPOI", false, true)]
    [InlineData("NPOI", false, false)]
    [InlineData("NPOI", true, true)]
    [InlineData("NPOI", true, false)]
    [InlineData("ClosedXML", false, true)]
    [InlineData("ClosedXML", false, false)]
    [InlineData("ClosedXML", true, true)]
    [InlineData("ClosedXML", true, false)]
    public async Task PreserveAndCreate_ShouldRejectNameOrTargetConflictBeforeWriting(string provider,
        bool asynchronous, bool conflictName)
    {
        using var template = new TemplatePivotTrackingStream(CreatePivotTemplate());
        template.Position = 7;
        var originalTemplate = template.ToArray();
        var definition = conflictName
            ? CreateDefinition("SalesByRegion", "NewSummary")
            : CreateDefinition("NewByRegion", "Summary");
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", new[]
            {
                new TemplatePivotRow { Region = "North", Amount = 10 },
                new TemplatePivotRow { Region = "South", Amount = 20 },
                new TemplatePivotRow { Region = "North", Amount = 30 }
            })
            .PreserveTemplatePivotTables()
            .PivotTable(definition));
        using var output = new MemoryStream();
        output.Write(new byte[] { 9, 8, 7 });
        output.Position = 1;
        var originalOutput = output.ToArray();

        var error = asynchronous
            ? await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
                ExportAsync(provider, request, output)).ConfigureAwait(false)
            : Assert.Throws<BingOfficesConfigurationException>(() =>
                Export(provider, request, output));

        Assert.Equal(BingOfficesStage.Preflight, error.Stage);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(7, template.Position);
        Assert.False(template.WasDisposed);
        Assert.Equal(originalOutput, output.ToArray());
        Assert.Equal(1, output.Position);
    }

    /// <summary>
    /// 验证只更新共享缓存的一个消费者时会拆分新缓存并保留原消费者。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdateSharedCache_ShouldSplitUnchangedConsumer(string provider, bool asynchronous)
    {
        using var template = new TemplatePivotTrackingStream(CreateSharedPivotTemplate());
        template.Position = 7;
        var originalTemplate = template.ToArray();
        var enumerated = false;
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", EnumerateRows(() => enumerated = true))
            .UpdatePivotTable(CreateDefinition("SalesByRegion", "Summary", "Data")));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        Assert.True(enumerated);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(7, template.Position);
        var cacheXml = ReadPivotCacheXml(output.ToArray());
        Assert.Contains("10", cacheXml, StringComparison.Ordinal);
        Assert.Contains("20", cacheXml, StringComparison.Ordinal);
        Assert.Contains("30", cacheXml, StringComparison.Ordinal);
        Assert.Contains("99", cacheXml, StringComparison.Ordinal);
        Assert.Contains("101", cacheXml, StringComparison.Ordinal);
        Assert.Equal(2, CountPivotCacheDefinitions(output.ToArray()));
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal("SalesByRegion", Assert.Single(workbook.Worksheet("Summary").PivotTables).Name);
        Assert.Equal("SalesByRegion2", Assert.Single(workbook.Worksheet("Summary2").PivotTables).Name);
    }

    /// <summary>
    /// 验证更新共享缓存的全部消费者时生成有效的新缓存。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdateSharedCache_WhenAllConsumersUpdated_ShouldCreateValidCaches(
        string provider, bool asynchronous)
    {
        using var template = new MemoryStream(CreateSharedPivotTemplate(), writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", EnumerateRows(() => { }))
            .UpdatePivotTable(CreateDefinition("SalesByRegion", "Summary", "Data"))
            .UpdatePivotTable(CreateDefinition("SalesByRegion2", "Summary2", "Data")));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        var cacheXml = ReadPivotCacheXml(output.ToArray());
        Assert.Contains("99", cacheXml, StringComparison.Ordinal);
        Assert.Contains("101", cacheXml, StringComparison.Ordinal);
        Assert.InRange(CountPivotCacheDefinitions(output.ToArray()), 1, 2);
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal("SalesByRegion", Assert.Single(workbook.Worksheet("Summary").PivotTables).Name);
        Assert.Equal("SalesByRegion2", Assert.Single(workbook.Worksheet("Summary2").PivotTables).Name);
    }

    /// <summary>
    /// 验证只更新部分模板透视时仍保留未更新透视及其缓存。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdatePartialTemplate_ShouldPreserveUnchangedPivot(string provider, bool asynchronous)
    {
        using var template = new TemplatePivotTrackingStream(CreateTwoPivotTemplate());
        template.Position = 7;
        var originalTemplate = template.ToArray();
        var enumerated = false;
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", EnumerateRows(() => enumerated = true))
            .UpdatePivotTable(CreateDefinition("SalesByRegion", "Summary", "Data")));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        Assert.True(enumerated);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(7, template.Position);
        var cacheXml = ReadPivotCacheXml(output.ToArray());
        Assert.Contains("10", cacheXml, StringComparison.Ordinal);
        Assert.Contains("20", cacheXml, StringComparison.Ordinal);
        Assert.Contains("30", cacheXml, StringComparison.Ordinal);
        Assert.Contains("99", cacheXml, StringComparison.Ordinal);
        Assert.Contains("101", cacheXml, StringComparison.Ordinal);
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal("SalesByRegion", Assert.Single(workbook.Worksheet("Summary").PivotTables).Name);
        Assert.Equal("SalesByRegion2", Assert.Single(workbook.Worksheet("Summary2").PivotTables).Name);
        Assert.DoesNotContain("__BingPivotUpdate_", string.Join("|",
            workbook.Worksheets.Select(sheet => sheet.Name)), StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证更新后置透视时仍能避开前置透视的部件路径。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task UpdatePartialTemplate_WhenSecondPivotUpdated_ShouldPreserveFirstPivot(
        string provider, bool asynchronous)
    {
        using var template = new MemoryStream(CreateTwoPivotTemplate(), writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data2", new[]
            {
                new TemplatePivotRow { Region = "West", Amount = 199 },
                new TemplatePivotRow { Region = "East", Amount = 201 }
            })
            .UpdatePivotTable(CreateDefinition("SalesByRegion2", "Summary2", "Data2")));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        var cacheXml = ReadPivotCacheXml(output.ToArray());
        Assert.Contains("10", cacheXml, StringComparison.Ordinal);
        Assert.Contains("20", cacheXml, StringComparison.Ordinal);
        Assert.Contains("30", cacheXml, StringComparison.Ordinal);
        Assert.Contains("199", cacheXml, StringComparison.Ordinal);
        Assert.Contains("201", cacheXml, StringComparison.Ordinal);
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal("SalesByRegion", Assert.Single(workbook.Worksheet("Summary").PivotTables).Name);
        Assert.Equal("SalesByRegion2", Assert.Single(workbook.Worksheet("Summary2").PivotTables).Name);
        Assert.DoesNotContain("__BingPivotUpdate_", string.Join("|",
            workbook.Worksheets.Select(sheet => sheet.Name)), StringComparison.Ordinal);
    }

    /// <summary>
    /// 使用 ClosedXML 创建包含原生透视部件的模板。
    /// </summary>
    /// <returns>完整 XLSX 模板字节。</returns>
    private static byte[] CreatePivotTemplate()
    {
        using var output = new MemoryStream();
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data",
            new[]
            {
                new TemplatePivotRow { Region = "North", Amount = 10 },
                new TemplatePivotRow { Region = "South", Amount = 20 },
                new TemplatePivotRow { Region = "North", Amount = 30 }
            }).PivotTable(new ExcelPivotTableDefinition
            {
                Name = "SalesByRegion",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition { EndRow = 3, EndColumn = 1 },
                TargetSheet = "Summary",
                TargetRow = 1,
                TargetColumn = 2,
                RowField = "Region",
                ValueField = "Amount"
            }));
        new ClosedXmlExcelExporter().Export(request, output);
        return output.ToArray();
    }

    /// <summary>
    /// 使用 ClosedXML 创建采购与仓储透视模板。
    /// </summary>
    /// <returns>包含基础及多聚合透视的 XLSX 模板。</returns>
    private static byte[] CreateBusinessPivotTemplate()
    {
        using var output = new MemoryStream();
        var request = ExcelExport.Workbook(builder => builder
            .AddSheet("PurchaseData", CreateBusinessRows(10))
            .AddSheet("WarehouseData", CreateBusinessRows(40))
            .PivotTable(CreateBusinessDefinition("PurchaseByRegion", "PurchaseData",
                "PurchaseSummary", includeMultipleMeasures: true))
            .PivotTable(CreateBusinessDefinition("WarehouseByRegion", "WarehouseData",
                "WarehouseSummary", includeMultipleMeasures: false)));
        new ClosedXmlExcelExporter().Export(request, output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建采购或仓储业务透视定义。
    /// </summary>
    /// <param name="name">透视名称。</param>
    /// <param name="sourceSheet">来源工作表。</param>
    /// <param name="targetSheet">目标工作表。</param>
    /// <param name="includeMultipleMeasures">是否配置五种聚合及多轴。</param>
    /// <returns>带状态筛选和数字格式的业务定义。</returns>
    private static ExcelPivotTableDefinition CreateBusinessDefinition(string name,
        string sourceSheet, string targetSheet, bool includeMultipleMeasures)
    {
        var values = includeMultipleMeasures
            ? new[]
            {
                new ExcelPivotValueDefinition
                {
                    Field = "Amount", Caption = "AmountSum", NumberFormat = "#,##0.000"
                },
                new ExcelPivotValueDefinition
                {
                    Field = "Amount", Caption = "AmountCount", Aggregation = ExcelPivotAggregation.Count
                },
                new ExcelPivotValueDefinition
                {
                    Field = "Amount", Caption = "AmountAverage", Aggregation = ExcelPivotAggregation.Average
                },
                new ExcelPivotValueDefinition
                {
                    Field = "Amount", Caption = "AmountMinimum", Aggregation = ExcelPivotAggregation.Minimum
                },
                new ExcelPivotValueDefinition
                {
                    Field = "Amount", Caption = "AmountMaximum", Aggregation = ExcelPivotAggregation.Maximum
                }
            }
            : new[]
            {
                new ExcelPivotValueDefinition
                {
                    Field = "Amount", Caption = "AmountSum", NumberFormat = "#,##0.000"
                }
            };
        return new ExcelPivotTableDefinition
        {
            Name = name,
            SourceSheet = sourceSheet,
            SourceRange = new ExcelRangeDefinition { EndRow = 3, EndColumn = 4 },
            TargetSheet = targetSheet,
            TargetRow = 1,
            TargetColumn = 2,
            RowFields = includeMultipleMeasures
                ? new[]
                {
                    new ExcelPivotAxisFieldDefinition { Field = "Region" },
                    new ExcelPivotAxisFieldDefinition
                    {
                        Field = "Category", SortOrder = ExcelPivotSortOrder.Descending,
                        SubtotalMode = ExcelPivotSubtotalMode.None
                    }
                }
                : new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } },
            ColumnFields = includeMultipleMeasures
                ? new[] { new ExcelPivotAxisFieldDefinition { Field = "Month" } }
                : Array.Empty<ExcelPivotAxisFieldDefinition>(),
            ValueFields = values,
            Filters = new[]
            {
                new ExcelPivotFilterDefinition
                {
                    Field = "Status", Values = new object[] { "Open" }
                }
            }
        };
    }

    /// <summary>
    /// 创建采购或仓储透视来源行。
    /// </summary>
    /// <param name="amountBase">用于区分模板与更新数据的金额基数。</param>
    /// <returns>包含开放和关闭状态的三条来源记录。</returns>
    private static TemplatePivotBusinessRow[] CreateBusinessRows(decimal amountBase) => new[]
    {
        new TemplatePivotBusinessRow
        {
            Region = "East", Category = "Hardware", Month = "Jan", Status = "Open",
            Amount = amountBase + 0.25m
        },
        new TemplatePivotBusinessRow
        {
            Region = "East", Category = "Hardware", Month = "Feb", Status = "Open",
            Amount = amountBase + 10.5m
        },
        new TemplatePivotBusinessRow
        {
            Region = "West", Category = "Software", Month = "Feb", Status = "Closed",
            Amount = amountBase + 20.75m
        }
    };

    /// <summary>
    /// 在指定工作表的 ZIP XML 中写入测试单元格。
    /// </summary>
    /// <param name="workbookBytes">原始工作簿。</param>
    /// <param name="sheetName">目标工作表名称。</param>
    /// <param name="values">单元格地址和值。</param>
    /// <returns>更新后的工作簿字节。</returns>
    private static byte[] WithWorksheetValues(byte[] workbookBytes, string sheetName,
        params (string Address, string Value)[] values)
    {
        using var input = new ZipArchive(new MemoryStream(workbookBytes, writable: false),
            ZipArchiveMode.Read);
        var workbook = ReadXml(input, "xl/workbook.xml");
        var relationships = ReadXml(input, "xl/_rels/workbook.xml.rels");
        var sheet = workbook.Descendants().Single(item => item.Name.LocalName == "sheet"
            && string.Equals((string)item.Attribute("name"), sheetName, StringComparison.Ordinal));
        var relationshipId = (string)sheet.Attributes().Single(item => item.Name.LocalName == "id");
        var relationship = relationships.Descendants().Single(item => item.Name.LocalName == "Relationship"
            && string.Equals((string)item.Attribute("Id"), relationshipId, StringComparison.Ordinal));
        var target = (string)relationship.Attribute("Target");
        var sheetPart = target.StartsWith("/", StringComparison.Ordinal)
            ? target.TrimStart('/') : "xl/" + target;
        var sheetXml = ReadXml(input, sheetPart);
        var sheetData = sheetXml.Descendants().Single(item => item.Name.LocalName == "sheetData");
        foreach (var (address, value) in values)
        {
            var rowNumber = new string(address.Where(char.IsDigit).ToArray());
            var row = sheetData.Elements().SingleOrDefault(item => item.Name.LocalName == "row"
                && string.Equals((string)item.Attribute("r"), rowNumber, StringComparison.Ordinal));
            if (row == null)
            {
                row = new XElement(sheetData.Name.Namespace + "row", new XAttribute("r", rowNumber));
                sheetData.Add(row);
            }
            var cell = row.Elements().SingleOrDefault(item => item.Name.LocalName == "c"
                && string.Equals((string)item.Attribute("r"), address, StringComparison.Ordinal));
            if (cell == null)
            {
                cell = new XElement(row.Name.Namespace + "c", new XAttribute("r", address));
                row.Add(cell);
            }
            cell.SetAttributeValue("t", "str");
            cell.Elements().Remove();
            cell.Add(new XElement(cell.Name.Namespace + "v", value));
        }

        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                var created = archive.CreateEntry(entry.FullName, CompressionLevel.Fastest);
                using var stream = created.Open();
                if (string.Equals(entry.FullName, sheetPart, StringComparison.OrdinalIgnoreCase))
                {
                    using var writer = new StreamWriter(stream, new UTF8Encoding(false), 1024,
                        leaveOpen: true);
                    sheetXml.Save(writer, System.Xml.Linq.SaveOptions.DisableFormatting);
                    writer.Flush();
                }
                else
                {
                    using var source = entry.Open();
                    source.CopyTo(stream);
                }
            }
        }
        return output.ToArray();
    }

    /// <summary>
    /// 设置测试透视部件声明的显示区域。
    /// </summary>
    private static byte[] WithPivotTargetReference(byte[] workbookBytes, string pivotName,
        string targetReference)
    {
        using var input = new ZipArchive(new MemoryStream(workbookBytes, writable: false),
            ZipArchiveMode.Read);
        var entry = input.Entries.Single(item => item.FullName.Contains("pivotTables/",
            StringComparison.OrdinalIgnoreCase) && item.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            && string.Equals((string)ReadXml(input, item.FullName).Root.Attribute("name"), pivotName,
                StringComparison.Ordinal));
        var document = ReadXml(input, entry.FullName);
        document.Descendants().Single(item => item.Name.LocalName == "location")
            .SetAttributeValue("ref", targetReference);

        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var sourceEntry in input.Entries)
            {
                var created = archive.CreateEntry(sourceEntry.FullName, CompressionLevel.Fastest);
                using var stream = created.Open();
                if (string.Equals(sourceEntry.FullName, entry.FullName, StringComparison.OrdinalIgnoreCase))
                {
                    using var writer = new StreamWriter(stream, new UTF8Encoding(false), 1024,
                        leaveOpen: true);
                    document.Save(writer, System.Xml.Linq.SaveOptions.DisableFormatting);
                    writer.Flush();
                }
                else
                {
                    using var source = sourceEntry.Open();
                    source.CopyTo(stream);
                }
            }
        }
        return output.ToArray();
    }

    /// <summary>
    /// 为模板透视设置自定义布局和样式描述。
    /// </summary>
    private static byte[] WithPivotLayout(byte[] workbookBytes, string pivotName)
    {
        using var input = new ZipArchive(new MemoryStream(workbookBytes, writable: false),
            ZipArchiveMode.Read);
        var entry = input.Entries.Single(item => item.FullName.Contains("pivotTables/",
            StringComparison.OrdinalIgnoreCase) && item.FullName.EndsWith(".xml",
            StringComparison.OrdinalIgnoreCase) && string.Equals((string)ReadXml(input,
                item.FullName).Root.Attribute("name"), pivotName, StringComparison.Ordinal));
        var document = ReadXml(input, entry.FullName);
        document.Root.SetAttributeValue("compact", "0");
        document.Root.SetAttributeValue("showEmptyCol", "0");
        document.Root.SetAttributeValue("showDataTips", "0");
        document.Root.SetAttributeValue("showHeaders", "0");
        document.Root.SetAttributeValue("rowGrandTotals", "0");
        var style = document.Root.Elements().SingleOrDefault(item =>
            item.Name.LocalName == "pivotTableStyleInfo");
        if (style == null)
        {
            style = new XElement(document.Root.Name.Namespace + "pivotTableStyleInfo");
            var anchor = document.Root.Elements().FirstOrDefault(item =>
                item.Name.LocalName is "filters" or "rowHierarchiesUsage"
                    or "colHierarchiesUsage" or "extLst");
            if (anchor == null)
                document.Root.Add(style);
            else
                anchor.AddBeforeSelf(style);
        }
        style.SetAttributeValue("name", "PivotStyleDark4");
        style.SetAttributeValue("showRowHeaders", "0");
        style.SetAttributeValue("showColHeaders", "1");
        style.SetAttributeValue("showRowStripes", "1");
        style.SetAttributeValue("showColStripes", "0");
        var formats = new XElement(document.Root.Name.Namespace + "formats",
            new XAttribute("count", "1"),
            new XElement(document.Root.Name.Namespace + "format",
                new XAttribute("action", "formatting"),
                new XElement(document.Root.Name.Namespace + "pivotArea",
                    new XAttribute("dataOnly", "0"))));
        var formatAnchor = document.Root.Elements().FirstOrDefault(item =>
            item.Name.LocalName is "conditionalFormats" or "chartFormats" or "pivotHierarchies"
                or "pivotTableStyleInfo" or "filters" or "rowHierarchiesUsage"
                or "colHierarchiesUsage" or "extLst");
        if (formatAnchor == null)
            document.Root.Add(formats);
        else
            formatAnchor.AddBeforeSelf(formats);

        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var sourceEntry in input.Entries)
            {
                var created = archive.CreateEntry(sourceEntry.FullName, CompressionLevel.Fastest);
                using var stream = created.Open();
                if (string.Equals(sourceEntry.FullName, entry.FullName, StringComparison.OrdinalIgnoreCase))
                {
                    using var writer = new StreamWriter(stream, new UTF8Encoding(false), 1024,
                        leaveOpen: true);
                    document.Save(writer, System.Xml.Linq.SaveOptions.DisableFormatting);
                    writer.Flush();
                }
                else
                {
                    using var source = sourceEntry.Open();
                    source.CopyTo(stream);
                }
            }
        }
        return output.ToArray();
    }

    /// <summary>
    /// 为透视定义注入无法识别的扩展节点。
    /// </summary>
    /// <param name="workbookBytes">原始工作簿。</param>
    /// <param name="pivotName">目标透视名称。</param>
    /// <returns>包含未知透视扩展的工作簿。</returns>
    private static byte[] WithPivotExtension(byte[] workbookBytes, string pivotName)
    {
        using var input = new ZipArchive(new MemoryStream(workbookBytes, writable: false),
            ZipArchiveMode.Read);
        var entry = input.Entries.Single(item => item.FullName.Contains("pivotTables/",
            StringComparison.OrdinalIgnoreCase) && item.FullName.EndsWith(".xml",
            StringComparison.OrdinalIgnoreCase) && string.Equals((string)ReadXml(input,
                item.FullName).Root.Attribute("name"), pivotName, StringComparison.Ordinal));
        var document = ReadXml(input, entry.FullName);
        var extensionList = document.Root.Elements().SingleOrDefault(item =>
            item.Name.LocalName == "extLst");
        if (extensionList == null)
        {
            extensionList = new XElement(document.Root.Name.Namespace + "extLst");
            document.Root.Add(extensionList);
        }
        extensionList.Add(new XElement(document.Root.Name.Namespace + "ext",
            new XAttribute("uri", "urn:bing-offices:test:unknown-pivot-extension"),
            new XElement(XName.Get("layoutMarker", "urn:bing-offices:test:pivot"),
                new XAttribute("fieldIndex", "0"))));

        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var sourceEntry in input.Entries)
            {
                var created = archive.CreateEntry(sourceEntry.FullName, CompressionLevel.Fastest);
                using var stream = created.Open();
                if (string.Equals(sourceEntry.FullName, entry.FullName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    using var writer = new StreamWriter(stream, new UTF8Encoding(false), 1024,
                        leaveOpen: true);
                    document.Save(writer, System.Xml.Linq.SaveOptions.DisableFormatting);
                    writer.Flush();
                }
                else
                {
                    using var source = sourceEntry.Open();
                    source.CopyTo(stream);
                }
            }
        }
        return output.ToArray();
    }

    /// <summary>
    /// 读取 ZIP 中的 XML 部件。
    /// </summary>
    private static XDocument ReadXml(ZipArchive archive, string partName)
    {
        using var stream = archive.GetEntry(partName).Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// 判断指定工作表单元格是否仍包含值节点。
    /// </summary>
    private static bool WorksheetCellHasValue(byte[] workbookBytes, string sheetName, string address)
    {
        using var archive = new ZipArchive(new MemoryStream(workbookBytes, writable: false));
        var workbook = ReadXml(archive, "xl/workbook.xml");
        var relationships = ReadXml(archive, "xl/_rels/workbook.xml.rels");
        var sheet = workbook.Descendants().Single(item => item.Name.LocalName == "sheet"
            && string.Equals((string)item.Attribute("name"), sheetName, StringComparison.Ordinal));
        var relationshipId = (string)sheet.Attributes().Single(item => item.Name.LocalName == "id");
        var relationship = relationships.Descendants().Single(item => item.Name.LocalName == "Relationship"
            && string.Equals((string)item.Attribute("Id"), relationshipId, StringComparison.Ordinal));
        var target = (string)relationship.Attribute("Target");
        var sheetPart = target.StartsWith("/", StringComparison.Ordinal)
            ? target.TrimStart('/') : "xl/" + target;
        var worksheet = ReadXml(archive, sheetPart);
        var cell = worksheet.Descendants().SingleOrDefault(item => item.Name.LocalName == "c"
            && string.Equals((string)item.Attribute("r"), address, StringComparison.Ordinal));
        return cell?.Elements().Any(item => item.Name.LocalName is "v" or "f" or "is") == true;
    }


    /// <summary>
    /// 使用 ClosedXML 创建包含两个透视部件的模板。
    /// </summary>
    /// <returns>完整 XLSX 模板字节。</returns>
    private static byte[] CreateTwoPivotTemplate()
    {
        using var output = new MemoryStream();
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data",
            new[]
            {
                new TemplatePivotRow { Region = "North", Amount = 10 },
                new TemplatePivotRow { Region = "South", Amount = 20 },
                new TemplatePivotRow { Region = "North", Amount = 30 }
            })
            .AddSheet("Data2",
            new[]
            {
                new TemplatePivotRow { Region = "North", Amount = 10 },
                new TemplatePivotRow { Region = "South", Amount = 20 },
                new TemplatePivotRow { Region = "North", Amount = 30 }
            })
            .PivotTable(CreateDefinition("SalesByRegion", "Summary", "Data"))
            .PivotTable(CreateDefinition("SalesByRegion2", "Summary2", "Data2")));
        new ClosedXmlExcelExporter().Export(request, output);
        return output.ToArray();
    }

    /// <summary>
    /// 使用同一来源创建共享缓存透视模板。
    /// </summary>
    /// <returns>完整 XLSX 模板字节。</returns>
    private static byte[] CreateSharedPivotTemplate()
    {
        using var output = new MemoryStream();
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data",
            new[]
            {
                new TemplatePivotRow { Region = "North", Amount = 10 },
                new TemplatePivotRow { Region = "South", Amount = 20 },
                new TemplatePivotRow { Region = "North", Amount = 30 }
            })
            .PivotTable(CreateDefinition("SalesByRegion", "Summary", "Data"))
            .PivotTable(CreateDefinition("SalesByRegion2", "Summary2", "Data")));
        new ClosedXmlExcelExporter().Export(request, output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建指定目标的基础透视定义。
    /// </summary>
    /// <param name="name">透视名称。</param>
    /// <param name="targetSheet">目标工作表。</param>
    /// <param name="sourceSheet">来源工作表。</param>
    /// <returns>透视定义。</returns>
    private static ExcelPivotTableDefinition CreateDefinition(string name, string targetSheet,
        string sourceSheet = "Data") => new()
    {
        Name = name,
        SourceSheet = sourceSheet,
        SourceRange = new ExcelRangeDefinition { EndRow = 3, EndColumn = 1 },
        TargetSheet = targetSheet,
        TargetRow = 1,
        TargetColumn = 2,
        RowField = "Region",
        ValueField = "Amount"
    };

    /// <summary>
    /// 创建一次枚举时执行回调的延迟数据源。
    /// </summary>
    /// <param name="onEnumeration">开始枚举时执行的回调。</param>
    /// <returns>透视来源行。</returns>
    private static System.Collections.Generic.IEnumerable<TemplatePivotRow> EnumerateRows(Action onEnumeration)
    {
        onEnumeration();
        yield return new TemplatePivotRow { Region = "West", Amount = 99 };
        yield return new TemplatePivotRow { Region = "East", Amount = 101 };
    }

    /// <summary>
    /// 读取输出包中的全部部件名称。
    /// </summary>
    /// <param name="bytes">完整 XLSX 字节。</param>
    /// <returns>按包顺序排列的部件名称。</returns>
    private static string[] ReadEntries(byte[] bytes)
    {
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        return archive.Entries.Select(entry => entry.FullName).ToArray();
    }

    /// <summary>
    /// 读取透视缓存定义和记录中的文本内容。
    /// </summary>
    /// <param name="bytes">完整 XLSX 字节。</param>
    /// <returns>缓存记录 XML 中的文本。</returns>
    private static string ReadPivotCacheXml(byte[] bytes)
    {
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        return string.Join("|", archive.Entries
            .Where(item => item.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .Select(item =>
            {
                using var stream = item.Open();
                var document = XDocument.Load(stream);
                return document.Root?.Name.LocalName is "pivotCacheDefinition" or "pivotCacheRecords"
                    ? document.ToString(System.Xml.Linq.SaveOptions.DisableFormatting) : string.Empty;
            }));
    }

    /// <summary>
    /// 读取指定名称的透视定义 XML。
    /// </summary>
    /// <param name="bytes">完整 XLSX 字节。</param>
    /// <param name="pivotName">透视名称。</param>
    /// <returns>透视定义文档。</returns>
    private static XDocument ReadPivotTableXml(byte[] bytes, string pivotName)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        var entry = archive.Entries.Single(item => item.FullName.Contains("pivotTables/",
            StringComparison.OrdinalIgnoreCase) && item.FullName.EndsWith(".xml",
            StringComparison.OrdinalIgnoreCase) && string.Equals(ReadPivotTableName(item), pivotName,
            StringComparison.Ordinal));
        return ReadXml(archive, entry.FullName);
    }

    /// <summary>
    /// 读取透视定义根元素的名称属性。
    /// </summary>
    /// <param name="entry">透视定义部件。</param>
    /// <returns>透视表名称。</returns>
    private static string ReadPivotTableName(System.IO.Compression.ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        var root = XDocument.Load(stream).Root;
        return root?.Name.LocalName == "pivotTableDefinition" ? (string)root.Attribute("name") : null;
    }

    /// <summary>
    /// 读取工作簿中声明的自定义数字格式。
    /// </summary>
    /// <param name="bytes">完整 XLSX 字节。</param>
    /// <returns>数字格式代码。</returns>
    private static string[] ReadCustomNumberFormats(byte[] bytes)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        return ReadXml(archive, "xl/styles.xml").Descendants()
            .Where(item => item.Name.LocalName == "numFmt")
            .Select(item => (string)item.Attribute("formatCode"))
            .ToArray();
    }

    /// <summary>
    /// 统计输出包中的透视缓存定义部件。
    /// </summary>
    /// <param name="bytes">完整 XLSX 字节。</param>
    /// <returns>透视缓存定义数量。</returns>
    private static int CountPivotCacheDefinitions(byte[] bytes)
    {
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        return archive.Entries.Count(entry => entry.FullName.EndsWith(".xml",
            StringComparison.OrdinalIgnoreCase) && ReadRootName(entry) == "pivotCacheDefinition");
    }

    /// <summary>
    /// 读取输出缓存的客户端打开刷新标记。
    /// </summary>
    /// <param name="bytes">完整 XLSX 字节。</param>
    /// <returns>按缓存定义部件顺序返回刷新标记。</returns>
    private static bool[] ReadPivotCacheRefreshFlags(byte[] bytes)
    {
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        return archive.Entries
            .Where(entry => entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .Select(entry =>
            {
                using var stream = entry.Open();
                var root = XDocument.Load(stream).Root;
                if (root?.Name.LocalName != "pivotCacheDefinition")
                    return (bool?)null;
                var value = (string)root.Attribute("refreshOnLoad");
                return value switch
                {
                    "1" or "true" or "True" or "TRUE" => true,
                    "0" or "false" or "False" or "FALSE" => false,
                    _ => (bool?)null
                };
            })
            .Where(value => value.HasValue)
            .Select(value => value.Value)
            .ToArray();
    }

    /// <summary>
    /// 获取透视桥接器使用的临时归档文件。
    /// </summary>
    /// <returns>按路径排序的临时归档文件列表。</returns>
    private static string[] GetArchiveTemporaryFiles() => Directory.GetFiles(Path.GetTempPath(),
        "bing-offices-pivot-template-archive-*.tmp").OrderBy(path => path,
        StringComparer.OrdinalIgnoreCase).ToArray();

    /// <summary>
    /// 获取透视部件存储使用的临时文件。
    /// </summary>
    /// <returns>按路径排序的临时部件文件列表。</returns>
    private static string[] GetPartsTemporaryFiles() => Directory.GetFiles(Path.GetTempPath(),
        "bing-offices-pivot-template-parts-*.tmp").OrderBy(path => path,
        StringComparer.OrdinalIgnoreCase).ToArray();

    /// <summary>
    /// 获取透视桥接器生成的中间模板临时文件。
    /// </summary>
    /// <returns>按路径排序的中间模板文件列表。</returns>
    private static string[] GetStrippedTemporaryFiles() => Directory.GetFiles(Path.GetTempPath(),
        "bing-offices-pivot-template-stripped-*.tmp").OrderBy(path => path,
        StringComparer.OrdinalIgnoreCase).ToArray();

    /// <summary>
    /// 获取 Provider 生成工作簿使用的临时文件。
    /// </summary>
    /// <returns>按路径排序的生成工作簿文件列表。</returns>
    private static string[] GetGeneratedTemporaryFiles() => Directory.GetFiles(Path.GetTempPath(),
        "bing-offices-pivot-template-generated-*.tmp").OrderBy(path => path,
        StringComparer.OrdinalIgnoreCase).ToArray();

    /// <summary>
    /// 获取 ClosedXML 异步桥接输出使用的临时文件。
    /// </summary>
    /// <returns>按路径排序的异步输出文件列表。</returns>
    private static string[] GetAsyncTemporaryFiles() => Directory.GetFiles(Path.GetTempPath(),
        "bing-offices-pivot-template-async-*.tmp").OrderBy(path => path,
        StringComparer.OrdinalIgnoreCase).ToArray();

    /// <summary>
    /// 读取 ZIP XML 部件根元素名称。
    /// </summary>
    /// <param name="entry">XML 部件。</param>
    /// <returns>根元素本地名称。</returns>
    private static string ReadRootName(System.IO.Compression.ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return XDocument.Load(stream).Root?.Name.LocalName;
    }

    /// <summary>
    /// 读取透视缓存记录的行数。
    /// </summary>
    /// <param name="bytes">完整 XLSX 字节。</param>
    /// <returns>缓存记录行数。</returns>
    private static int ReadPivotCacheRecordRows(byte[] bytes)
    {
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        foreach (var entry in archive.Entries.Where(item => item.FullName.EndsWith(".xml",
            StringComparison.OrdinalIgnoreCase)))
        {
            using var stream = entry.Open();
            var document = XDocument.Load(stream);
            if (document.Root?.Name.LocalName == "pivotCacheRecords")
                return document.Root.Elements().Count(item => item.Name.LocalName == "r");
        }
        throw new InvalidDataException("输出缺少透视缓存记录部件。");
    }

    /// <summary>
    /// 调用指定 Provider 的同步接口。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="request">工作簿请求。</param>
    /// <param name="output">目标流。</param>
    private static void Export(string provider, ExcelWorkbookExportRequest request, Stream output,
        CancellationToken cancellationToken = default)
    {
        if (provider == "NPOI")
            new NpoiExcelExporter().Export(request, output, cancellationToken);
        else
            new ClosedXmlExcelExporter().Export(request, output, cancellationToken);
    }

    /// <summary>
    /// 调用指定 Provider 的异步接口。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="request">工作簿请求。</param>
    /// <param name="output">目标流。</param>
    /// <returns>表示导出操作的异步任务。</returns>
    private static Task ExportAsync(string provider, ExcelWorkbookExportRequest request, Stream output,
        CancellationToken cancellationToken = default) =>
        provider == "NPOI"
            ? new NpoiExcelExporter().ExportAsync(request, output, cancellationToken)
            : new ClosedXmlExcelExporter().ExportAsync(request, output, cancellationToken);

    /// <summary>
    /// 检查异常链中是否包含流写入异常。
    /// </summary>
    /// <param name="exception">待检查的异常。</param>
    /// <returns>异常链包含 <see cref="IOException"/> 时返回 true。</returns>
    private static bool ContainsIOException(Exception exception) =>
        exception is IOException || exception.InnerException != null
        && ContainsIOException(exception.InnerException);

    /// <summary>
    /// 首次写入时抛出异常的测试流。
    /// </summary>
    private sealed class FailOnceWriteStream : MemoryStream
    {
        /// <summary>
        /// 指示是否仍需让下一次写入失败。
        /// </summary>
        private bool _failNextWrite = true;

        /// <summary>
        /// 判断当前写入是否应触发一次性故障。
        /// </summary>
        /// <returns>首次写入时返回 true。</returns>
        private bool ShouldFailWrite()
        {
            if (!_failNextWrite)
                return false;
            _failNextWrite = false;
            return true;
        }

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            if (ShouldFailWrite())
                throw new IOException("Injected template update destination failure.");
            base.Write(buffer, offset, count);
        }

        /// <inheritdoc />
        public override void Write(ReadOnlySpan<byte> buffer)
        {
            if (ShouldFailWrite())
                throw new IOException("Injected template update destination failure.");
            base.Write(buffer);
        }

        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            if (ShouldFailWrite())
                return Task.FromException(new IOException("Injected template update destination failure."));
            return base.WriteAsync(buffer, offset, count, cancellationToken);
        }

        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            if (ShouldFailWrite())
                return new ValueTask(Task.FromException(
                    new IOException("Injected template update destination failure.")));
            return base.WriteAsync(buffer, cancellationToken);
        }
    }

    /// <summary>
    /// 读取透视表及缓存相关原始部件。
    /// </summary>
    /// <param name="bytes">完整 XLSX 文件。</param>
    /// <returns>按部件名称索引的原始字节。</returns>
    private static System.Collections.Generic.Dictionary<string, byte[]> ReadPivotPackageParts(byte[] bytes)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes, writable: false),
            ZipArchiveMode.Read);
        return archive.Entries.Where(entry => entry.FullName.StartsWith("xl/pivotTables/",
                StringComparison.OrdinalIgnoreCase)
            || entry.FullName.StartsWith("xl/pivotCache/", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(entry => entry.FullName, entry =>
            {
                using var source = entry.Open();
                using var buffer = new MemoryStream();
                source.CopyTo(buffer);
                return buffer.ToArray();
            }, StringComparer.OrdinalIgnoreCase);
    }

    private sealed class TemplatePivotRow
    {
        /// <summary>
        /// 获取或设置区域。
        /// </summary>
        public string Region { get; set; }

        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public int Amount { get; set; }
    }

    /// <summary>
    /// 表示模板透视业务场景中的来源记录。
    /// </summary>
    private sealed class TemplatePivotBusinessRow
    {
        /// <summary>
        /// 获取或设置区域名称。
        /// </summary>
        public string Region { get; set; }

        /// <summary>
        /// 获取或设置商品类别。
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        /// 获取或设置月份标签。
        /// </summary>
        public string Month { get; set; }

        /// <summary>
        /// 获取或设置单据状态。
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// 获取或设置业务金额。
        /// </summary>
        public decimal Amount { get; set; }
    }
}
