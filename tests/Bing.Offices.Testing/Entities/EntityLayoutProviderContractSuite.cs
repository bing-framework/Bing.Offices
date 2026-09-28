using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;

namespace Bing.Offices.Testing.Entities;

/// <summary>
/// 通过公开接口验证第三方 Provider 的实体布局行为。
/// </summary>
/// <remarks>
/// 此源码可链接到独立测试工程；调用方应传入声明支持 XLSX Entity Layout 的导入器和导出器。
/// 失败时抛出包含场景名称的 <see cref="InvalidOperationException" />。
/// </remarks>
public static class EntityLayoutProviderContractSuite
{
    /// <summary>
    /// 验证动态列分组、最终尾部、名称和明细往返。
    /// </summary>
    /// <param name="exporter">待验证的实体导出器。</param>
    /// <param name="importer">待验证的实体导入器。</param>
    public static void VerifyCore(IExcelEntityExporter exporter, IExcelEntityImporter importer)
    {
        ValidateServices(exporter, importer);
        var layout = ExcelEntity.Layout<DynamicOrder>(builder => builder
            .ListRegion("Contract", "A1", order => order.Lines, region => region
                .DynamicColumnGroup("goods", line => line.Goods, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "goods-color", Title = "Color", DataType = typeof(string)
                    }
                })
                .DynamicColumnGroup("product", line => line.Product, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "product-weight", Title = "Weight", DataType = typeof(decimal)
                    }
                })
                .FooterNamed("ContractTotal", "TOTAL", footer => footer
                    .Cell("B1", lines => lines.Sum(line => line.Quantity)))));
        var source = new DynamicOrder
        {
            Lines = new List<DynamicLine>
            {
                new() { Name = "A", Quantity = 2,
                    Goods = new Dictionary<string, object> { ["goods-color"] = "red" },
                    Product = new Dictionary<string, object> { ["product-weight"] = 1.5m } },
                new() { Name = "B", Quantity = 3,
                    Goods = new Dictionary<string, object> { ["goods-color"] = "blue" },
                    Product = new Dictionary<string, object> { ["product-weight"] = 2.5m } }
            }
        };

        var (entity, workbook) = RoundTrip(exporter, importer, source, layout, "Core");
        Require(entity.Lines != null && entity.Lines.Count == 2,
            "Core", "尾部被导入为明细，或明细数量错误。");
        Require(entity.Lines[0].Name == "A" && entity.Lines[0].Quantity == 2
            && entity.Lines[1].Name == "B" && entity.Lines[1].Quantity == 3,
            "Core", "固定列没有完整读回。");
        var first = entity.Lines[0];
        var second = entity.Lines[1];
        Require(first.Goods != null && second.Goods != null
            && first.Product != null && second.Product != null
            && first.Goods.TryGetValue("goods-color", out var firstColor)
            && second.Goods.TryGetValue("goods-color", out var secondColor)
            && first.Product.TryGetValue("product-weight", out var firstWeight)
            && second.Product.TryGetValue("product-weight", out var secondWeight)
            && Equals(firstColor, "red") && Equals(secondColor, "blue")
            && Equals(firstWeight, 1.5m) && Equals(secondWeight, 2.5m),
            "Core", "两个动态字典没有分别读回。");
        AssertFooter(importer, workbook, "A4", "B4", "TOTAL", 5m, "Core");
        AssertName(workbook, "ContractTotal", "$A$4");
    }

    /// <summary>
    /// 验证属性式固定单元格及 CellNamed、ListRegionNamed 在 XLSX 模板中的导入导出。
    /// </summary>
    /// <param name="exporter">待验证的实体导出器。</param>
    /// <param name="importer">待验证的实体导入器。</param>
    public static void VerifyAttributeAndNamedAnchors(IExcelEntityExporter exporter,
        IExcelEntityImporter importer)
    {
        ValidateServices(exporter, importer);
        var source = new AnchoredOrder
        {
            Number = "PO-ANCHOR-1",
            Customer = "Anchor Customer",
            Lines = new List<AnchoredLine>
            {
                new() { Code = "A", Quantity = 2 },
                new() { Code = "B", Quantity = 3 }
            }
        };
        var layout = ExcelEntity.LayoutFromAttributes<AnchoredOrder>(builder => builder
            .CellNamed("Contract", "ContractCustomer", order => order.Customer)
            .ListRegionNamed("Contract", "ContractLines", order => order.Lines,
                region => region.Footer("TOTAL")));
        var template = CreateNamedAnchorTemplate(exporter);

        using var templateInput = new MemoryStream(template, writable: false);
        using var output = new MemoryStream();
        exporter.ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(templateInput, leaveOpen: true), output);
        Require(output.CanWrite && output.Length > 0,
            "AttributeAndNamedAnchors", "模板导出未产生工作簿或关闭了调用方流。");

        var workbook = output.ToArray();
        AssertName(workbook, "ContractCustomer", "$C$2", "AttributeAndNamedAnchors");
        AssertName(workbook, "ContractLines", "$D$5", "AttributeAndNamedAnchors");
        using var input = new MemoryStream(workbook, writable: false);
        using var importTemplate = new MemoryStream(template, writable: false);
        var result = importer.ImportForTemplate(input, layout,
            new ExcelEntityTemplateOptions(importTemplate, leaveOpen: true));
        Require(result != null && result.IsSuccess && result.Entity != null
            && result.Entity.Number == source.Number && result.Entity.Customer == source.Customer
            && result.Entity.Lines != null && result.Entity.Lines.Count == source.Lines.Count
            && result.Entity.Lines[0].Code == "A" && result.Entity.Lines[0].Quantity == 2
            && result.Entity.Lines[1].Code == "B" && result.Entity.Lines[1].Quantity == 3,
            "AttributeAndNamedAnchors", "属性固定单元格或命名锚点内容没有完整往返。");
        Require(input.CanRead, "AttributeAndNamedAnchors", "模板导入关闭了调用方来源流。");
    }

    /// <summary>
    /// 验证命名列表按模板实际位置与固定单元格重叠时，模板导入导出均报告配置错误。
    /// </summary>
    /// <param name="exporter">待验证的实体导出器。</param>
    /// <param name="importer">待验证的实体导入器。</param>
    public static void VerifyNamedListFixedCellCollision(IExcelEntityExporter exporter,
        IExcelEntityImporter importer)
    {
        ValidateServices(exporter, importer);
        var source = new AnchoredOrder
        {
            Number = "PO-COLLISION-1",
            Customer = "Collision Customer",
            Collision = "fixed value",
            Lines = new List<AnchoredLine> { new() { Code = "A", Quantity = 1 } }
        };
        var layout = ExcelEntity.LayoutFromAttributes<AnchoredOrder>(builder => builder
            .CellNamed("Contract", "ContractCustomer", order => order.Customer)
            .Cell("Contract", "D6", order => order.Collision)
            .ListRegionNamed("Contract", "ContractLines", order => order.Lines,
                region => region.Footer("TOTAL")));
        var template = CreateNamedAnchorTemplate(exporter);

        using var exportTemplate = new MemoryStream(template, writable: false);
        var existingOutput = new byte[] { 0x31, 0x32, 0x33 };
        using var output = new MemoryStream(existingOutput);
        var outputPosition = output.Position;
        AssertLayoutConfigurationFailure(() => exporter.ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(exportTemplate, leaveOpen: true), output),
            "NamedListFixedCellCollision export");
        Require(output.ToArray().SequenceEqual(existingOutput) && output.Position == outputPosition,
            "NamedListFixedCellCollision", "配置冲突破坏了已有目标流内容或位置。");

        var validLayout = ExcelEntity.LayoutFromAttributes<AnchoredOrder>(builder => builder
            .CellNamed("Contract", "ContractCustomer", order => order.Customer)
            .ListRegionNamed("Contract", "ContractLines", order => order.Lines,
                region => region.Footer("TOTAL")));
        using var validTemplate = new MemoryStream(template, writable: false);
        using var validOutput = new MemoryStream();
        exporter.ExportForTemplate(source, validLayout,
            new ExcelEntityTemplateOptions(validTemplate, leaveOpen: true), validOutput);

        using var importSource = new MemoryStream(validOutput.ToArray(), writable: false);
        using var importTemplate = new MemoryStream(template, writable: false);
        AssertLayoutConfigurationFailure(() => importer.ImportForTemplate(importSource, layout,
            new ExcelEntityTemplateOptions(importTemplate, leaveOpen: true)),
            "NamedListFixedCellCollision import");
        Require(importSource.CanRead, "NamedListFixedCellCollision", "配置冲突关闭了导入来源流。");
    }

    /// <summary>
    /// 验证中间分页小计与最终尾部的输出和导入边界。
    /// </summary>
    /// <param name="exporter">待验证的实体导出器。</param>
    /// <param name="importer">待验证的实体导入器。</param>
    public static void VerifyPageSubtotal(IExcelEntityExporter exporter, IExcelEntityImporter importer)
    {
        ValidateServices(exporter, importer);
        var layout = ExcelEntity.Layout<SubtotalOrder>(builder => builder
            .ListRegion("Contract", "A1", order => order.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("PAGE", footer => footer
                    .Cell("B1", lines => lines.Sum(line => line.Amount)))
                .Footer("TOTAL", footer => footer
                    .Cell("B1", lines => lines.Sum(line => line.Amount)))));
        var source = new SubtotalOrder
        {
            Lines = new List<SubtotalLine>
            {
                new() { Category = "A", Amount = 2m },
                new() { Category = "B", Amount = 3m },
                new() { Category = "C", Amount = 4m }
            }
        };

        var (entity, workbook) = RoundTrip(exporter, importer, source, layout, "PageSubtotal");
        AssertLines(entity, new[] { "A", "B", "C" }, new[] { 2m, 3m, 4m }, "PageSubtotal");
        AssertFooter(importer, workbook, "A4", "B4", "PAGE", 5m, "PageSubtotal");
        AssertFooter(importer, workbook, "A6", "B6", "TOTAL", 9m, "PageSubtotal");
    }

    /// <summary>
    /// 验证连续分组小计与最终尾部的输出和导入边界。
    /// </summary>
    /// <param name="exporter">待验证的实体导出器。</param>
    /// <param name="importer">待验证的实体导入器。</param>
    public static void VerifyGroupSubtotal(IExcelEntityExporter exporter, IExcelEntityImporter importer)
    {
        ValidateServices(exporter, importer);
        var layout = ExcelEntity.Layout<SubtotalOrder>(builder => builder
            .ListRegion("Contract", "A1", order => order.Lines, region => region
                .GroupSubtotal(line => line.Category, "GROUP", footer => footer
                    .Cell("B1", lines => lines.Sum(line => line.Amount)))
                .Footer("TOTAL", footer => footer
                    .Cell("B1", lines => lines.Sum(line => line.Amount)))));
        var source = new SubtotalOrder
        {
            Lines = new List<SubtotalLine>
            {
                new() { Category = "A", Amount = 2m },
                new() { Category = "A", Amount = 3m },
                new() { Category = "B", Amount = 4m }
            }
        };

        var (entity, workbook) = RoundTrip(exporter, importer, source, layout, "GroupSubtotal");
        AssertLines(entity, new[] { "A", "A", "B" }, new[] { 2m, 3m, 4m }, "GroupSubtotal");
        AssertFooter(importer, workbook, "A4", "B4", "GROUP", 5m, "GroupSubtotal");
        AssertFooter(importer, workbook, "A6", "B6", "GROUP", 4m, "GroupSubtotal");
        AssertFooter(importer, workbook, "A7", "B7", "TOTAL", 9m, "GroupSubtotal");
    }

    /// <summary>
    /// 验证尾部原生公式、连续明细求和及导入边界。
    /// </summary>
    /// <param name="exporter">待验证的实体导出器。</param>
    /// <param name="importer">待验证的实体导入器。</param>
    public static void VerifyFooterFormulas(IExcelEntityExporter exporter,
        IExcelEntityImporter importer)
    {
        ValidateServices(exporter, importer);
        var layout = ExcelEntity.Layout<SubtotalOrder>(builder => builder
            .ListRegion("Contract", "A1", order => order.Lines, region => region
                .Header(false)
                .Footer("TOTAL", footer => footer
                    .GapRows(1)
                    .Formula("C1", "=1+1")
                    .FormulaSumContiguousRowsAbove("C2", "b", numberFormat: "0.00"))));
        var source = new SubtotalOrder
        {
            Lines = new List<SubtotalLine>
            {
                new() { Category = "A", Amount = 2m },
                new() { Category = "B", Amount = 3m }
            }
        };

        var (entity, workbook) = RoundTrip(exporter, importer, source, layout, "FooterFormulas");
        AssertLines(entity, new[] { "A", "B" }, new[] { 2m, 3m }, "FooterFormulas");
        AssertFormula(workbook, "C4", "1+1", "FooterFormulas");
        AssertFormula(workbook, "C5",
            "SUM(INDEX(B:B,ROW()-4):INDEX(B:B,ROW()-3))", "FooterFormulas");
        AssertMarker(importer, workbook, "A4", "TOTAL", "FooterFormulas");

        var (emptyEntity, emptyWorkbook) = RoundTrip(exporter, importer,
            new SubtotalOrder(), layout, "FooterFormulas empty");
        Require(emptyEntity.Lines != null && emptyEntity.Lines.Count == 0,
            "FooterFormulas empty", "空明细被导入为尾部数据。");
        AssertFormula(emptyWorkbook, "C2", "1+1", "FooterFormulas empty");
        AssertFormula(emptyWorkbook, "C3", "0", "FooterFormulas empty");
        AssertMarker(importer, emptyWorkbook, "A2", "TOTAL", "FooterFormulas empty");

        var pagedLayout = ExcelEntity.Layout<SubtotalOrder>(builder => builder
            .ListRegion("Contract", "A1", order => order.Lines, region => region
                .Header(false)
                .PageBreak(1)
                .PageSubtotal("PAGE")
                .Footer("TOTAL", footer => footer
                    .FormulaSumDetailRowsAbove("C1", "B"))));
        var (pagedEntity, pagedWorkbook) = RoundTrip(exporter, importer, source,
            pagedLayout, "FooterFormulas paged");
        AssertLines(pagedEntity, new[] { "A", "B" }, new[] { 2m, 3m },
            "FooterFormulas paged");
        AssertFormula(pagedWorkbook, "C4", "SUM(B1:B1,B3:B3)",
            "FooterFormulas paged");
        AssertMarker(importer, pagedWorkbook, "A2", "PAGE", "FooterFormulas paged");
        AssertMarker(importer, pagedWorkbook, "A4", "TOTAL", "FooterFormulas paged");

        try
        {
            ExcelEntity.Layout<SubtotalOrder>(builder => builder
                .ListRegion("Contract", "A1", order => order.Lines, region => region
                    .PageBreak(1)
                    .PageSubtotal("PAGE")
                    .Footer("TOTAL", footer => footer
                        .FormulaSumContiguousRowsAbove("C1", "B"))));
            throw new InvalidOperationException(
                "Entity contract FooterFormulas: 跨分页小计的最终求和未在布局阶段拒绝。");
        }
        catch (ArgumentException)
        {
            return;
        }
    }

    /// <summary>
    /// 验证异步往返、预取消与调用方流所有权。
    /// </summary>
    /// <param name="exporter">待验证的实体导出器。</param>
    /// <param name="importer">待验证的实体导入器。</param>
    public static async Task VerifyAsync(IExcelEntityExporter exporter, IExcelEntityImporter importer)
    {
        ValidateServices(exporter, importer);
        var layout = ExcelEntity.Layout<SubtotalOrder>(builder => builder
            .ListRegion("Contract", "A1", order => order.Lines, region => region.Footer("TOTAL")));
        var source = new SubtotalOrder
        {
            Lines = new List<SubtotalLine> { new() { Category = "A", Amount = 2m } }
        };
        using var output = new MemoryStream();
        await exporter.ExportEntityAsync(source, layout, output).ConfigureAwait(false);
        Require(output.CanWrite && output.Length > 0, "Async", "异步导出未产生工作簿或关闭了调用方流。");
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = await importer.ImportEntityAsync(input, layout).ConfigureAwait(false);
        Require(result != null && result.IsSuccess && result.Entity?.Lines?.Count == 1
            && result.Entity.Lines[0].Category == "A" && result.Entity.Lines[0].Amount == 2m,
            "Async", "异步导入未完整读回明细。");
        Require(input.CanRead, "Async", "异步导入关闭了调用方流。");

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var canceledOutput = new MemoryStream(new byte[] { 1, 2, 3 });
        canceledOutput.Position = canceledOutput.Length;
        await AssertCanceled(() => exporter.ExportEntityAsync(source, layout, canceledOutput,
            cancellation.Token), "Async export").ConfigureAwait(false);
        Require(canceledOutput.ToArray().SequenceEqual(new byte[] { 1, 2, 3 })
            && canceledOutput.CanWrite, "Async", "预取消破坏了目标流。");
        using var canceledInput = new MemoryStream(output.ToArray(), writable: false);
        await AssertCanceled(() => importer.ImportEntityAsync(canceledInput, layout,
            cancellation.Token), "Async import").ConfigureAwait(false);
        Require(canceledInput.CanRead, "Async", "预取消关闭了来源流。");
    }

    /// <summary>
    /// 验证服务参数均已提供。
    /// </summary>
    private static void ValidateServices(IExcelEntityExporter exporter, IExcelEntityImporter importer)
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
    }

    /// <summary>
    /// 导出并重新导入一个实体布局工作簿。
    /// </summary>
    private static (TEntity Entity, byte[] Workbook) RoundTrip<TEntity>(IExcelEntityExporter exporter,
        IExcelEntityImporter importer, TEntity source, ExcelEntityLayout<TEntity> layout, string scenario)
        where TEntity : class, new()
    {
        using var output = new MemoryStream();
        exporter.ExportEntity(source, layout, output);
        Require(output.CanWrite && output.Length > 0, scenario, "未输出完整工作簿或关闭了调用方流。");
        var bytes = output.ToArray();
        using var input = new MemoryStream(bytes, writable: false);
        var result = importer.ImportEntity(input, layout);
        Require(result != null && result.IsSuccess && result.Entity != null, scenario,
            "导入失败: " + string.Join("; ", result?.Errors?.Select(error => error.Message)
                ?? Array.Empty<string>()));
        Require(input.CanRead, scenario, "导入关闭了调用方流。");
        return (result.Entity, bytes);
    }

    /// <summary>
    /// 检查导入的分组明细完整性。
    /// </summary>
    private static void AssertLines(SubtotalOrder entity, string[] categories, decimal[] amounts,
        string scenario)
    {
        Require(entity.Lines != null
            && entity.Lines.Select(line => line.Category).SequenceEqual(categories)
            && entity.Lines.Select(line => line.Amount).SequenceEqual(amounts),
            scenario, "小计或尾部被导入为明细，或明细内容不一致。");
    }

    /// <summary>
    /// 通过公开固定单元格导入接口检查尾部标记及聚合值。
    /// </summary>
    private static void AssertFooter(IExcelEntityImporter importer, byte[] workbook,
        string markerAddress, string totalAddress, string expectedMarker, decimal expectedTotal,
        string scenario)
    {
        var layout = ExcelEntity.Layout<FooterProbe>(builder => builder
            .Cell("Contract", markerAddress, probe => probe.Marker)
            .Cell("Contract", totalAddress, probe => probe.Total));
        using var input = new MemoryStream(workbook, writable: false);
        var result = importer.ImportEntity(input, layout);
        Require(result != null && result.IsSuccess && result.Entity != null
            && result.Entity.Marker == expectedMarker
            && result.Entity.Total == expectedTotal, scenario,
            $"尾部 {markerAddress}/{totalAddress} 内容错误或无法导入。");
    }

    /// <summary>
    /// 检查尾部标记的完整文本。
    /// </summary>
    private static void AssertMarker(IExcelEntityImporter importer, byte[] workbook,
        string address, string expectedMarker, string scenario)
    {
        var layout = ExcelEntity.Layout<FooterProbe>(builder => builder
            .Cell("Contract", address, probe => probe.Marker));
        using var input = new MemoryStream(workbook, writable: false);
        var result = importer.ImportEntity(input, layout);
        Require(result != null && result.IsSuccess && result.Entity?.Marker == expectedMarker,
            scenario, $"尾部标记 {address} 内容错误。");
    }

    /// <summary>
    /// 检查 XLSX 单元格存储的原生公式。
    /// </summary>
    private static void AssertFormula(byte[] workbook, string address,
        string expectedFormula, string scenario)
    {
        using var input = new MemoryStream(workbook, writable: false);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read);
        XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var workbookEntry = archive.GetEntry("xl/workbook.xml");
        var relationshipsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
        Require(workbookEntry != null && relationshipsEntry != null, scenario,
            "XLSX 缺少工作簿或工作表关系。");
        using var workbookXml = workbookEntry.Open();
        using var relationshipsXml = relationshipsEntry.Open();
        var workbookDocument = XDocument.Load(workbookXml);
        var relationshipsDocument = XDocument.Load(relationshipsXml);
        var sheet = workbookDocument.Descendants(main + "sheet")
            .SingleOrDefault(element => (string)element.Attribute("name") == "Contract");
        var relationshipId = (string)sheet?.Attribute(relationships + "id");
        var target = relationshipsDocument.Descendants()
            .Where(element => element.Name.LocalName == "Relationship")
            .Where(element => (string)element.Attribute("Id") == relationshipId)
            .Select(element => (string)element.Attribute("Target"))
            .SingleOrDefault();
        Require(!string.IsNullOrWhiteSpace(target), scenario, "XLSX 无法定位 Contract 工作表。");
        var path = new Uri(new Uri("http://package/xl/workbook.xml"), target)
            .AbsolutePath.TrimStart('/');
        var entry = archive.GetEntry(path);
        Require(entry != null, scenario, "XLSX 缺少工作表内容。");
        using var xml = entry.Open();
        var document = XDocument.Load(xml);
        var cells = document.Descendants(main + "c")
            .Where(cell => string.Equals((string)cell.Attribute("r"), address,
                StringComparison.OrdinalIgnoreCase)).ToArray();
        Require(cells.Length == 1 && (string)cells[0].Element(main + "f") == expectedFormula,
            scenario, $"单元格 {address} 没有保存预期原生公式 {expectedFormula}。");
    }

    /// <summary>
    /// 在 XLSX 工作簿元数据中检查最终尾部名称。
    /// </summary>
    private static void AssertName(byte[] workbook, string name, string expectedAddress,
        string scenario = "Core")
    {
        using var input = new MemoryStream(workbook, writable: false);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read);
        var entry = archive.GetEntry("xl/workbook.xml");
        Require(entry != null, "Core", "XLSX 缺少工作簿元数据。");
        using var xml = entry.Open();
        var document = XDocument.Load(xml);
        XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var matches = document.Descendants(main + "definedName")
            .Where(item => string.Equals((string)item.Attribute("name"), name,
                StringComparison.OrdinalIgnoreCase)).ToArray();
        var formula = matches.Length == 1 ? matches[0].Value : string.Empty;
        var separator = formula.LastIndexOf('!');
        var sheet = separator >= 0 ? formula.Substring(0, separator).Trim('\'') : string.Empty;
        var address = separator >= 0 ? formula.Substring(separator + 1) : string.Empty;
        var references = address.Split(':');
        Require(matches.Length == 1 && string.Equals(sheet, "Contract", StringComparison.OrdinalIgnoreCase)
            && references.Length is 1 or 2
            && references.All(reference => string.Equals(reference, expectedAddress,
                StringComparison.OrdinalIgnoreCase)), scenario, "命名区域未指向预期单元格: "
            + string.Join("; ", matches.Select(item => item.Value)));
    }

    /// <summary>
    /// 以 Provider 生成的 XLSX 为基底，加入模板布局需要的工作簿级单格名称。
    /// </summary>
    /// <param name="exporter">用于生成有效 XLSX 基底的实体导出器。</param>
    /// <returns>含有固定单元格和列表区域名称的 XLSX 模板。</returns>
    private static byte[] CreateNamedAnchorTemplate(IExcelEntityExporter exporter)
    {
        var seedLayout = ExcelEntity.Layout<AnchoredOrder>(builder => builder
            .Cell("Contract", "B2", order => order.Number)
            .Cell("Contract", "C2", order => order.Customer)
            .ListRegion("Contract", "D5", order => order.Lines));
        using var output = new MemoryStream();
        exporter.ExportEntity(new AnchoredOrder(), seedLayout, output);
        using var input = new MemoryStream(output.ToArray(), writable: false);
        using var source = new ZipArchive(input, ZipArchiveMode.Read);
        using var result = new MemoryStream();
        using (var destination = new ZipArchive(result, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in source.Entries)
            {
                var copy = destination.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                using var entryInput = entry.Open();
                using var entryOutput = copy.Open();
                if (!string.Equals(entry.FullName, "xl/workbook.xml", StringComparison.Ordinal))
                {
                    entryInput.CopyTo(entryOutput);
                    continue;
                }

                var document = XDocument.Load(entryInput);
                XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                var root = document.Root ?? throw new InvalidOperationException(
                    "Entity contract AttributeAndNamedAnchors: XLSX 工作簿元数据为空。");
                var names = root.Element(main + "definedNames");
                if (names == null)
                {
                    names = new XElement(main + "definedNames");
                    var sheets = root.Element(main + "sheets");
                    if (sheets == null)
                        root.Add(names);
                    else
                        sheets.AddAfterSelf(names);
                }
                names.Add(
                    new XElement(main + "definedName", new XAttribute("name", "ContractCustomer"),
                        "'Contract'!$C$2"),
                    new XElement(main + "definedName", new XAttribute("name", "ContractLines"),
                        "'Contract'!$D$5"));
                document.Save(entryOutput);
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 检查布局冲突以结构化 Plan 配置异常失败。
    /// </summary>
    /// <param name="operation">执行布局导入或导出的操作。</param>
    /// <param name="scenario">用于诊断的合同场景名称。</param>
    private static void AssertLayoutConfigurationFailure(Action operation, string scenario)
    {
        try
        {
            operation();
        }
        catch (BingOfficesConfigurationException exception)
        {
            Require(exception.Code == BingOfficesErrorCode.ConfigurationInvalid
                && exception.Operation == BingOfficesOperation.Configuration
                && exception.Stage == BingOfficesStage.Plan,
                scenario, "布局冲突未返回 Plan 阶段的结构化配置错误。");
            return;
        }
        throw new InvalidOperationException($"Entity contract {scenario}: 命名列表与固定单元格冲突未被拒绝。");
    }

    /// <summary>
    /// 确认操作直接观察预取消令牌。
    /// </summary>
    private static async Task AssertCanceled(Func<Task> operation, string scenario)
    {
        try
        {
            await operation().ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            return;
        }
        throw new InvalidOperationException($"Entity contract {scenario}: 未观察预取消令牌。");
    }

    /// <summary>
    /// 确认场景断言成立。
    /// </summary>
    private static void Require(bool condition, string scenario, string message)
    {
        if (!condition)
            throw new InvalidOperationException($"Entity contract {scenario}: {message}");
    }

    /// <summary>
    /// 动态列合同的根实体。
    /// </summary>
    private sealed class DynamicOrder
    {
        /// <summary>
        /// 获取或设置明细集合。
        /// </summary>
        public List<DynamicLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 动态列合同的明细实体。
    /// </summary>
    private sealed class DynamicLine
    {
        /// <summary>
        /// 获取或设置明细名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置明细数量。
        /// </summary>
        public int Quantity { get; set; }
        /// <summary>
        /// 获取或设置商品动态值。
        /// </summary>
        public IDictionary<string, object> Goods { get; set; } = new Dictionary<string, object>();
        /// <summary>
        /// 获取或设置产品动态值。
        /// </summary>
        public IDictionary<string, object> Product { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>
    /// 小计合同的根实体。
    /// </summary>
    private sealed class SubtotalOrder
    {
        /// <summary>
        /// 获取或设置明细集合。
        /// </summary>
        public List<SubtotalLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 小计合同的明细实体。
    /// </summary>
    private sealed class SubtotalLine
    {
        /// <summary>
        /// 获取或设置分组名称。
        /// </summary>
        public string Category { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }
    }

    /// <summary>
    /// 尾部单元格合同的读取实体。
    /// </summary>
    private sealed class FooterProbe
    {
        /// <summary>
        /// 获取或设置标记文本。
        /// </summary>
        public string Marker { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置聚合值。
        /// </summary>
        public decimal Total { get; set; }
    }

    /// <summary>
    /// 固定单元格与命名锚点合同的根实体。
    /// </summary>
    private sealed class AnchoredOrder
    {
        /// <summary>
        /// 获取或设置订单编号。
        /// </summary>
        [ExcelEntityCell("Contract", "B2")]
        public string Number { get; set; } = string.Empty;

        /// <summary>
        /// 获取或设置客户名称。
        /// </summary>
        public string Customer { get; set; } = string.Empty;

        /// <summary>
        /// 获取或设置冲突测试中的固定单元格值。
        /// </summary>
        public string Collision { get; set; } = string.Empty;

        /// <summary>
        /// 获取或设置订单明细。
        /// </summary>
        public List<AnchoredLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 固定单元格与命名锚点合同的明细实体。
    /// </summary>
    private sealed class AnchoredLine
    {
        /// <summary>
        /// 获取或设置明细编码。
        /// </summary>
        public string Code { get; set; } = string.Empty;

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }
}
