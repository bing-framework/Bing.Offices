using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Exports;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证模板透视未知扩展的更新安全边界。
/// </summary>
public sealed class TemplatePivotDependencyBoundaryTest
{
    /// <summary>
    /// 验证透视定义包含未知扩展时在枚举业务数据前拒绝更新。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task Update_WhenPivotContainsUnknownExtension_ShouldRejectBeforeEnumeration(
        string provider, bool asynchronous)
    {
        using var template = new MemoryStream(AddPivotExtension(CreateTemplate()), writable: true);
        template.Position = 9;
        var originalTemplate = template.ToArray();
        var enumerated = false;
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", EnumerateRows(() => enumerated = true))
            .UpdatePivotTable(CreateDefinition()));
        using var output = new MemoryStream();
        output.Write(new byte[] { 9, 8, 7, 6 });
        output.Position = 2;

        BingOfficesConfigurationException error;
        if (asynchronous)
            error = await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
                ExportAsync(provider, request, output)).ConfigureAwait(false);
        else
            error = Assert.Throws<BingOfficesConfigurationException>(() =>
                Export(provider, request, output));

        Assert.Contains("未验证的扩展或依赖", error.Message, StringComparison.Ordinal);
        Assert.False(enumerated);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(9, template.Position);
        Assert.Equal(new byte[] { 9, 8, 7, 6 }, output.ToArray());
        Assert.Equal(2, output.Position);
    }

    /// <summary>
    /// 验证缓存使用命名来源时在枚举业务数据前拒绝更新。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task Update_WhenCacheUsesNamedSource_ShouldRejectBeforeEnumeration(
        string provider, bool asynchronous)
    {
        using var template = new MemoryStream(AddCacheSourceName(CreateTemplate()), writable: true);
        template.Position = 9;
        var originalTemplate = template.ToArray();
        var enumerated = false;
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", EnumerateRows(() => enumerated = true))
            .UpdatePivotTable(CreateDefinition()));
        using var output = new MemoryStream();
        output.Write(new byte[] { 9, 8, 7, 6 });
        output.Position = 2;

        BingOfficesConfigurationException error;
        if (asynchronous)
            error = await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
                ExportAsync(provider, request, output)).ConfigureAwait(false);
        else
            error = Assert.Throws<BingOfficesConfigurationException>(() =>
                Export(provider, request, output));

        Assert.Contains("未验证的扩展或依赖", error.Message, StringComparison.Ordinal);
        Assert.False(enumerated);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(9, template.Position);
        Assert.Equal(new byte[] { 9, 8, 7, 6 }, output.ToArray());
        Assert.Equal(2, output.Position);
    }

    /// <summary>
    /// 验证透视部件包含未知关系时在枚举业务数据前拒绝更新。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task Update_WhenPivotContainsUnexpectedRelationship_ShouldRejectBeforeEnumeration(
        string provider, bool asynchronous)
    {
        using var template = new MemoryStream(AddUnexpectedRelationship(CreateTemplate()), writable: true);
        template.Position = 9;
        var originalTemplate = template.ToArray();
        var enumerated = false;
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", EnumerateRows(() => enumerated = true))
            .UpdatePivotTable(CreateDefinition()));
        using var output = new MemoryStream();
        output.Write(new byte[] { 9, 8, 7, 6 });
        output.Position = 2;

        BingOfficesConfigurationException error;
        if (asynchronous)
            error = await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
                ExportAsync(provider, request, output)).ConfigureAwait(false);
        else
            error = Assert.Throws<BingOfficesConfigurationException>(() =>
                Export(provider, request, output));

        Assert.Contains("未验证的扩展或依赖", error.Message, StringComparison.Ordinal);
        Assert.False(enumerated);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(9, template.Position);
        Assert.Equal(new byte[] { 9, 8, 7, 6 }, output.ToArray());
        Assert.Equal(2, output.Position);
    }

    /// <summary>
    /// 验证保留模式不会丢失模板透视扩展内容。
    /// </summary>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public void Preserve_WhenPivotContainsUnknownExtension_ShouldKeepMarkup(string provider)
    {
        var templateBytes = AddPivotExtension(CreateTemplate());
        using var template = new MemoryStream(templateBytes, writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .PreserveTemplatePivotTables());
        using var output = new MemoryStream();

        Export(provider, request, output);

        Assert.Contains("urn:bing-offices:unknown-pivot-extension",
            ReadPivotParts(output.ToArray()).SelectMany(document => document.Descendants())
                .Select(element => (string)element.Attribute("uri"))
                .Where(value => value != null), StringComparer.Ordinal);
    }

    /// <summary>
    /// 验证透视图、切片器和时间线关联在枚举业务数据前被明确拒绝更新。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    /// <param name="feature">模板关联功能。</param>
    /// <param name="displayName">预期的结构化拒绝说明。</param>
    [Theory]
    [InlineData("NPOI", false, "PivotChart", "透视图")]
    [InlineData("NPOI", true, "PivotChart", "透视图")]
    [InlineData("ClosedXML", false, "PivotChart", "透视图")]
    [InlineData("ClosedXML", true, "PivotChart", "透视图")]
    [InlineData("NPOI", false, "Slicer", "切片器")]
    [InlineData("NPOI", true, "Slicer", "切片器")]
    [InlineData("ClosedXML", false, "Slicer", "切片器")]
    [InlineData("ClosedXML", true, "Slicer", "切片器")]
    [InlineData("NPOI", false, "Timeline", "时间线")]
    [InlineData("NPOI", true, "Timeline", "时间线")]
    [InlineData("ClosedXML", false, "Timeline", "时间线")]
    [InlineData("ClosedXML", true, "Timeline", "时间线")]
    public async Task Update_WhenTemplateContainsRelatedFeature_ShouldRejectBeforeEnumeration(
        string provider, bool asynchronous, string feature, string displayName)
    {
        using var template = new MemoryStream(AddRelatedFeature(CreateTemplate(), feature), writable: true);
        template.Position = 9;
        var originalTemplate = template.ToArray();
        var enumerated = false;
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", EnumerateRows(() => enumerated = true))
            .UpdatePivotTable(CreateDefinition()));
        using var output = new MemoryStream();
        output.Write(new byte[] { 9, 8, 7, 6 });
        output.Position = 2;

        BingOfficesConfigurationException error;
        if (asynchronous)
            error = await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
                ExportAsync(provider, request, output)).ConfigureAwait(false);
        else
            error = Assert.Throws<BingOfficesConfigurationException>(() =>
                Export(provider, request, output));

        Assert.Contains("不支持的关联功能", error.Message, StringComparison.Ordinal);
        Assert.Contains(displayName, error.Message, StringComparison.Ordinal);
        Assert.False(enumerated);
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(9, template.Position);
        Assert.Equal(new byte[] { 9, 8, 7, 6 }, output.ToArray());
        Assert.Equal(2, output.Position);
    }

    /// <summary>
    /// 验证显式保留模式会保留关联功能的关系和部件。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="feature">模板关联功能。</param>
    /// <param name="partName">关联功能部件。</param>
    /// <param name="relationshipPart">声明关联关系的部件。</param>
    /// <param name="relationshipType">关联关系类型。</param>
    [Theory]
    [InlineData("NPOI", "PivotChart", "xl/charts/chart1.xml", "xl/drawings/_rels/drawing1.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")]
    [InlineData("ClosedXML", "PivotChart", "xl/charts/chart1.xml", "xl/drawings/_rels/drawing1.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")]
    [InlineData("NPOI", "Slicer", "xl/slicerCaches/slicerCache1.xml", "xl/_rels/workbook.xml.rels",
        "http://schemas.microsoft.com/office/2007/relationships/slicerCache")]
    [InlineData("ClosedXML", "Slicer", "xl/slicerCaches/slicerCache1.xml", "xl/_rels/workbook.xml.rels",
        "http://schemas.microsoft.com/office/2007/relationships/slicerCache")]
    [InlineData("NPOI", "Timeline", "xl/timelineCaches/timelineCache1.xml", "xl/_rels/workbook.xml.rels",
        "http://schemas.microsoft.com/office/2010/relationships/timelineCache")]
    [InlineData("ClosedXML", "Timeline", "xl/timelineCaches/timelineCache1.xml", "xl/_rels/workbook.xml.rels",
        "http://schemas.microsoft.com/office/2010/relationships/timelineCache")]
    public void Preserve_WhenTemplateContainsRelatedFeature_ShouldRetainItsPackageParts(
        string provider, string feature, string partName, string relationshipPart, string relationshipType)
    {
        var templateBytes = AddRelatedFeature(CreateTemplate(), feature);
        var originalParts = Unpack(templateBytes);
        using var template = new MemoryStream(templateBytes, writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .PreserveTemplatePivotTables());
        using var output = new MemoryStream();

        Export(provider, request, output);

        var outputParts = Unpack(output.ToArray());
        Assert.Equal(originalParts[partName], outputParts[partName]);
        var relationships = XDocument.Load(new MemoryStream(outputParts[relationshipPart]));
        Assert.Contains(relationships.Root.Elements(), element =>
            (string)element.Attribute("Type") == relationshipType);
    }

    /// <summary>
    /// 验证普通图表和表格切片器不会阻止无关透视更新。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="asynchronous">是否调用异步接口。</param>
    /// <param name="feature">与透视无关的工作簿功能。</param>
    [Theory]
    [InlineData("NPOI", false, "Chart")]
    [InlineData("NPOI", true, "Chart")]
    [InlineData("ClosedXML", false, "Chart")]
    [InlineData("ClosedXML", true, "Chart")]
    [InlineData("NPOI", false, "TableSlicer")]
    [InlineData("NPOI", true, "TableSlicer")]
    [InlineData("ClosedXML", false, "TableSlicer")]
    [InlineData("ClosedXML", true, "TableSlicer")]
    public async Task Update_WhenTemplateContainsUnrelatedFeature_ShouldProceed(
        string provider, bool asynchronous, string feature)
    {
        var templateBytes = AddRelatedFeature(CreateTemplate(), feature);
        var originalParts = Unpack(templateBytes);
        using var template = new MemoryStream(templateBytes, writable: true);
        template.Position = 7;
        var originalTemplate = template.ToArray();
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", new[]
            {
                new TemplatePivotRow { Region = "North", Amount = 99 },
                new TemplatePivotRow { Region = "South", Amount = 101 }
            })
            .UpdatePivotTable(CreateDefinition()));
        using var output = new MemoryStream();

        if (asynchronous)
            await ExportAsync(provider, request, output).ConfigureAwait(false);
        else
            Export(provider, request, output);

        var outputParts = Unpack(output.ToArray());
        var featurePart = feature == "Chart" ? "xl/charts/chart1.xml" : "xl/slicerCaches/slicerCache1.xml";
        var originalFeature = XDocument.Load(new MemoryStream(originalParts[featurePart]));
        var outputFeature = XDocument.Load(new MemoryStream(outputParts[featurePart]));
        Assert.Equal(originalFeature.Root.Name, outputFeature.Root.Name);
        Assert.Equal(originalFeature.Root.Descendants().Select(element => element.Name),
            outputFeature.Root.Descendants().Select(element => element.Name));
        var featureRelationships = XDocument.Load(new MemoryStream(outputParts[
            feature == "Chart" ? "xl/drawings/_rels/drawing1.xml.rels" : "xl/_rels/workbook.xml.rels"]));
        Assert.Contains(featureRelationships.Root.Elements(), element =>
            (string)element.Attribute("Type") == (feature == "Chart"
                ? "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
                : "http://schemas.microsoft.com/office/2007/relationships/slicerCache"));
        Assert.Equal(originalTemplate, template.ToArray());
        Assert.Equal(7, template.Position);
    }

    /// <summary>
    /// 创建包含单个透视的模板。
    /// </summary>
    /// <returns>完整 XLSX 模板字节。</returns>
    private static byte[] CreateTemplate()
    {
        using var output = new MemoryStream();
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data", new[]
        {
            new TemplatePivotRow { Region = "North", Amount = 10 },
            new TemplatePivotRow { Region = "South", Amount = 20 },
            new TemplatePivotRow { Region = "North", Amount = 30 }
        }).PivotTable(CreateDefinition()));
        new ClosedXmlExcelExporter().Export(request, output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建基础透视定义。
    /// </summary>
    /// <returns>指向模板来源的透视定义。</returns>
    private static ExcelPivotTableDefinition CreateDefinition() => new()
    {
        Name = "SalesByRegion",
        SourceSheet = "Data",
        SourceRange = new ExcelRangeDefinition { EndRow = 3, EndColumn = 1 },
        TargetSheet = "Summary",
        TargetRow = 1,
        TargetColumn = 2,
        RowField = "Region",
        ValueField = "Amount"
    };

    /// <summary>
    /// 创建枚举时通知测试的延迟数据源。
    /// </summary>
    /// <param name="onEnumeration">开始枚举时执行的回调。</param>
    /// <returns>两行透视来源。</returns>
    private static IEnumerable<TemplatePivotRow> EnumerateRows(Action onEnumeration)
    {
        onEnumeration();
        yield return new TemplatePivotRow { Region = "East", Amount = 99 };
        yield return new TemplatePivotRow { Region = "West", Amount = 101 };
    }

    /// <summary>
    /// 在透视定义中加入未验证的扩展标记。
    /// </summary>
    /// <param name="bytes">原始 XLSX 字节。</param>
    /// <returns>带扩展标记的 XLSX 字节。</returns>
    private static byte[] AddPivotExtension(byte[] bytes)
    {
        var parts = Unpack(bytes);
        var pivot = parts.First(pair => pair.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            && IsRoot(pair.Value, "pivotTableDefinition"));
        var document = XDocument.Load(new MemoryStream(pivot.Value));
        var root = document.Root;
        root.Add(new XElement(root.Name.Namespace + "extLst",
            new XElement(root.Name.Namespace + "ext",
                new XAttribute("uri", "urn:bing-offices:unknown-pivot-extension"),
                new XElement(XName.Get("customPivotState", "urn:bing-offices:custom"),
                    new XAttribute("fieldIndex", "0")))));
        parts[pivot.Key] = Save(document);
        return Pack(parts);
    }

    /// <summary>
    /// 为缓存来源加入命名区域标识。
    /// </summary>
    /// <param name="bytes">原始 XLSX 字节。</param>
    /// <returns>带命名来源的 XLSX 字节。</returns>
    private static byte[] AddCacheSourceName(byte[] bytes)
    {
        var parts = Unpack(bytes);
        var cache = parts.First(pair => pair.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            && IsRoot(pair.Value, "pivotCacheDefinition"));
        var document = XDocument.Load(new MemoryStream(cache.Value));
        document.Root.Descendants().Single(element => element.Name.LocalName == "worksheetSource")
            .SetAttributeValue("name", "SalesSource");
        parts[cache.Key] = Save(document);
        return Pack(parts);
    }

    /// <summary>
    /// 为透视部件加入未声明的内部关系。
    /// </summary>
    /// <param name="bytes">原始 XLSX 字节。</param>
    /// <returns>带未声明关系的 XLSX 字节。</returns>
    private static byte[] AddUnexpectedRelationship(byte[] bytes)
    {
        var parts = Unpack(bytes);
        var relationship = parts.First(pair => pair.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
            && IsPivotRelationshipDocument(pair.Value));
        var document = XDocument.Load(new MemoryStream(relationship.Value));
        document.Root.Add(new XElement(document.Root.Name.Namespace + "Relationship",
            new XAttribute("Id", "rIdUnverified"),
            new XAttribute("Type", "urn:bing-offices:relationships/customPivotFeature"),
            new XAttribute("Target", "../custom/custom1.xml")));
        parts[relationship.Key] = Save(document);
        return Pack(parts);
    }

    /// <summary>
    /// 在模板包中加入指定的透视关联功能。
    /// </summary>
    /// <param name="bytes">原始 XLSX 字节。</param>
    /// <param name="feature">关联功能名称。</param>
    /// <returns>包含关联部件和关系的 XLSX 字节。</returns>
    private static byte[] AddRelatedFeature(byte[] bytes, string feature)
    {
        var parts = Unpack(bytes);
        switch (feature)
        {
            case "PivotChart":
            case "Chart":
                parts["xl/drawings/drawing1.xml"] = Save(new XDocument(new XElement(
                    XName.Get("wsDr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"))));
                var chartSpace = new XElement(XName.Get("chartSpace",
                    "http://schemas.openxmlformats.org/drawingml/2006/chart"));
                if (feature == "PivotChart")
                    chartSpace.Add(new XElement(XName.Get("pivotSource",
                        "http://schemas.microsoft.com/office/drawing/2012/chart"),
                        new XElement(XName.Get("name", "http://schemas.openxmlformats.org/drawingml/2006/chart"),
                            "SalesByRegion")));
                parts["xl/charts/chart1.xml"] = Save(new XDocument(chartSpace));
                AddRelationship(parts, "xl/drawings/_rels/drawing1.xml.rels", "rIdPivotChart",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", "../charts/chart1.xml");
                var summarySheet = FindWorksheetPart(parts, "Summary");
                var summaryDocument = XDocument.Load(new MemoryStream(parts[summarySheet]));
                summaryDocument.Root.Add(new XElement(
                    XName.Get("drawing", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"),
                    new XAttribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
                        "rIdPivotDrawing")));
                parts[summarySheet] = Save(summaryDocument);
                var summaryFileName = summarySheet[(summarySheet.LastIndexOf('/') + 1)..];
                var summaryDirectory = summarySheet[..(summarySheet.LastIndexOf('/') + 1)];
                AddRelationship(parts, summaryDirectory + "_rels/" + summaryFileName + ".rels", "rIdPivotDrawing",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                    "../drawings/drawing1.xml");
                AddContentType(parts, "/xl/drawings/drawing1.xml",
                    "application/vnd.openxmlformats-officedocument.drawing+xml");
                AddContentType(parts, "/xl/charts/chart1.xml",
                    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
                break;
            case "Slicer":
            case "TableSlicer":
                var slicerCache = new XElement(
                    XName.Get("slicerCacheDefinition", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"));
                if (feature == "Slicer")
                    slicerCache.Add(new XElement(
                        XName.Get("pivotTables", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"),
                        new XElement(XName.Get("pivotTable", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"),
                            new XAttribute("name", "SalesByRegion"))));
                else
                    slicerCache.Add(new XElement(
                        XName.Get("sourceName", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"),
                        "Items"));
                parts["xl/slicerCaches/slicerCache1.xml"] = Save(new XDocument(slicerCache));
                AddRelationship(parts, "xl/_rels/workbook.xml.rels", "rIdSlicerCache",
                    "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
                    "slicerCaches/slicerCache1.xml");
                AddContentType(parts, "/xl/slicerCaches/slicerCache1.xml", "application/vnd.ms-excel.slicerCache+xml");
                break;
            case "Timeline":
                parts["xl/timelineCaches/timelineCache1.xml"] = Save(new XDocument(new XElement(
                    XName.Get("timelineCacheDefinition", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"),
                    new XElement(XName.Get("pivotTables", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"),
                        new XElement(XName.Get("pivotTable", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"),
                            new XAttribute("name", "SalesByRegion"))))));
                AddRelationship(parts, "xl/_rels/workbook.xml.rels", "rIdTimelineCache",
                    "http://schemas.microsoft.com/office/2010/relationships/timelineCache",
                    "timelineCaches/timelineCache1.xml");
                AddContentType(parts, "/xl/timelineCaches/timelineCache1.xml",
                    "application/vnd.ms-excel.TimelineCache+xml");
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(feature));
        }

        return Pack(parts);
    }

    /// <summary>
    /// 根据工作表名称解析工作表部件。
    /// </summary>
    /// <param name="parts">模板部件集合。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <returns>工作表部件名称。</returns>
    private static string FindWorksheetPart(Dictionary<string, byte[]> parts, string sheetName)
    {
        var workbook = XDocument.Load(new MemoryStream(parts["xl/workbook.xml"]));
        var relationshipId = (string)workbook.Descendants().Single(element => element.Name.LocalName == "sheet"
            && (string)element.Attribute("name") == sheetName).Attribute(
                XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
        var relationships = XDocument.Load(new MemoryStream(parts["xl/_rels/workbook.xml.rels"]));
        var target = (string)relationships.Root.Elements().Single(element =>
            (string)element.Attribute("Id") == relationshipId).Attribute("Target");
        return target.StartsWith("/", StringComparison.Ordinal)
            ? target.TrimStart('/')
            : "xl/" + target;
    }

    /// <summary>
    /// 在模板包中加入关系声明。
    /// </summary>
    /// <param name="parts">模板部件集合。</param>
    /// <param name="relationshipPart">关系部件名称。</param>
    /// <param name="id">关系标识。</param>
    /// <param name="type">关系类型。</param>
    /// <param name="target">关系目标。</param>
    private static void AddRelationship(Dictionary<string, byte[]> parts, string relationshipPart,
        string id, string type, string target)
    {
        XDocument document;
        if (parts.TryGetValue(relationshipPart, out var bytes))
        {
            document = XDocument.Load(new MemoryStream(bytes));
        }
        else
        {
            document = new XDocument(new XElement(
                XName.Get("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships")));
        }

        document.Root.Add(new XElement(document.Root.Name.Namespace + "Relationship",
            new XAttribute("Id", id), new XAttribute("Type", type), new XAttribute("Target", target)));
        parts[relationshipPart] = Save(document);
    }

    /// <summary>
    /// 在模板包中加入部件内容类型。
    /// </summary>
    /// <param name="parts">模板部件集合。</param>
    /// <param name="partName">部件名称。</param>
    /// <param name="contentType">部件内容类型。</param>
    private static void AddContentType(Dictionary<string, byte[]> parts, string partName, string contentType)
    {
        var document = XDocument.Load(new MemoryStream(parts["[Content_Types].xml"]));
        document.Root.Add(new XElement(document.Root.Name.Namespace + "Override",
            new XAttribute("PartName", partName), new XAttribute("ContentType", contentType)));
        parts["[Content_Types].xml"] = Save(document);
    }

    /// <summary>
    /// 读取输出中的透视 XML 部件。
    /// </summary>
    /// <param name="bytes">完整 XLSX 字节。</param>
    /// <returns>透视 XML 文档。</returns>
    private static IEnumerable<XDocument> ReadPivotParts(byte[] bytes)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes));
        foreach (var entry in archive.Entries.Where(item => item.FullName.EndsWith(".xml",
                     StringComparison.OrdinalIgnoreCase)))
        {
            using var stream = entry.Open();
            var document = XDocument.Load(stream);
            if (document.Root?.Name.LocalName is "pivotTableDefinition" or "pivotCacheDefinition")
                yield return document;
        }
    }

    /// <summary>
    /// 按 Provider 调用同步导出接口。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="request">工作簿请求。</param>
    /// <param name="output">目标流。</param>
    private static void Export(string provider, ExcelWorkbookExportRequest request, Stream output)
    {
        if (provider == "NPOI")
            new NpoiExcelExporter().Export(request, output);
        else
            new ClosedXmlExcelExporter().Export(request, output);
    }

    /// <summary>
    /// 按 Provider 调用异步导出接口。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="request">工作簿请求。</param>
    /// <param name="output">目标流。</param>
    /// <returns>表示异步导出操作的任务。</returns>
    private static Task ExportAsync(string provider, ExcelWorkbookExportRequest request, Stream output) =>
        provider == "NPOI"
            ? new NpoiExcelExporter().ExportAsync(request, output)
            : new ClosedXmlExcelExporter().ExportAsync(request, output);

    /// <summary>
    /// 判断 XML 根元素名称。
    /// </summary>
    /// <param name="bytes">XML 部件字节。</param>
    /// <param name="rootName">期望的根元素本地名称。</param>
    /// <returns>根元素匹配时返回 <see langword="true" />。</returns>
    private static bool IsRoot(byte[] bytes, string rootName)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        return XDocument.Load(stream).Root?.Name.LocalName == rootName;
    }

    /// <summary>
    /// 判断关系部件是否属于透视定义。
    /// </summary>
    /// <param name="bytes">关系 XML 部件字节。</param>
    /// <returns>包含透视缓存关系时返回 <see langword="true" />。</returns>
    private static bool IsPivotRelationshipDocument(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        var document = XDocument.Load(stream);
        return document.Root?.Elements().Any(element =>
            ((string)element.Attribute("Type"))?.EndsWith("/pivotCacheDefinition",
                StringComparison.Ordinal) == true) == true;
    }

    /// <summary>
    /// 读取 ZIP 部件。
    /// </summary>
    /// <param name="bytes">完整 ZIP 字节。</param>
    /// <returns>按部件名称索引的内容。</returns>
    private static Dictionary<string, byte[]> Unpack(byte[] bytes)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes));
        return archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var input = entry.Open();
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        });
    }

    /// <summary>
    /// 写出 ZIP 部件。
    /// </summary>
    /// <param name="parts">ZIP 部件集合。</param>
    /// <returns>完整 ZIP 字节。</returns>
    private static byte[] Pack(IReadOnlyDictionary<string, byte[]> parts)
    {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var part in parts)
            {
                using var stream = archive.CreateEntry(part.Key).Open();
                stream.Write(part.Value, 0, part.Value.Length);
            }
        return output.ToArray();
    }

    /// <summary>
    /// 保存 XML 部件。
    /// </summary>
    /// <param name="document">XML 文档。</param>
    /// <returns>XML 字节。</returns>
    private static byte[] Save(XDocument document)
    {
        using var output = new MemoryStream();
        document.Save(output, SaveOptions.DisableFormatting);
        return output.ToArray();
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
}
