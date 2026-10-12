using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Internals;
using Bing.Offices.Tests;
using Xunit;

namespace Bing.Offices.Npoi.Tests.Exports.PivotTables;

/// <summary>
/// 验证 NPOI 透视模板部件发现及公开导出的拒绝路径。
/// </summary>
[Collection("NpoiNativePivot")]
public sealed class ExcelPivotTemplatePartsTest
{
    private const string ContentTypePrefix =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.";
    private const string RelationshipPrefix =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string StrictRelationshipPrefix =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/";

    /// <summary>
    /// 验证真实 NPOI 工作簿清单包含完整透视表和缓存部件。
    /// </summary>
    [Fact]
    public void Find_ShouldReturnCompletePartsFromNpoiWorkbook()
    {
        var bytes = CreatePivotWorkbook();
        using var stream = new MemoryStream(bytes);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        var parts = ExcelPivotTemplateParts.Find(archive);

        Assert.Equal(new[]
        {
            "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
            "xl/pivotCache/pivotCacheDefinition1.xml",
            "xl/pivotCache/pivotCacheRecords1.xml",
            "xl/pivotTables/_rels/pivotTable1.xml.rels",
            "xl/pivotTables/pivotTable1.xml"
        }, parts);
        Assert.All(parts, part => Assert.NotNull(archive.GetEntry(part)));
        Assert.True(stream.CanRead);
        Assert.True(stream.CanWrite);
    }

    /// <summary>
    /// 验证发现器跟随非标准路径、关系目标和内容类型。
    /// </summary>
    [Fact]
    public void Find_ShouldResolveRenamedPartsFromRelationshipsAndContentTypes()
    {
        var bytes = CreatePivotWorkbook();
        using var originalStream = new MemoryStream(bytes);
        using var originalArchive = new ZipArchive(originalStream, ZipArchiveMode.Read, leaveOpen: true);
        var originalParts = ExcelPivotTemplateParts.Find(originalArchive);
        var mapping = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["xl/pivotTables/pivotTable1.xml"] = "customPivot/Report.xml",
            ["xl/pivotCache/pivotCacheDefinition1.xml"] = "customPivot/CacheDefinition.xml",
            ["xl/pivotCache/pivotCacheRecords1.xml"] = "customPivot/CacheRecords.xml"
        };

        var renamed = RenameParts(bytes, mapping, out var relationshipParts);
        using var renamedStream = new MemoryStream(renamed);
        using var renamedArchive = new ZipArchive(renamedStream, ZipArchiveMode.Read, leaveOpen: true);
        var actual = ExcelPivotTemplateParts.Find(renamedArchive);
        var expected = originalParts.Where(part => !relationshipParts.Contains(part))
            .Select(part => mapping.GetValueOrDefault(part, part))
            .OrderBy(part => part, StringComparer.Ordinal).ToArray();

        Assert.Equal(expected, actual);
        Assert.Contains("customPivot/Report.xml", actual);
        Assert.Contains("customPivot/CacheDefinition.xml", actual);
        Assert.Contains("customPivot/CacheRecords.xml", actual);
        Assert.All(actual, part => Assert.NotNull(renamedArchive.GetEntry(part)));
    }

    /// <summary>
    /// 验证仅由 OPC 关系声明的非标准部件也会被发现。
    /// </summary>
    [Fact]
    public void Find_ShouldReturnRelationshipOnlyNonstandardPart()
    {
        var bytes = CreateRelationshipPackage("../../customPart/Records.bin", includeTarget: true);
        using var stream = new MemoryStream(bytes);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        var parts = ExcelPivotTemplateParts.Find(archive);

        Assert.Equal(new[] { "customPart/Records.bin" }, parts);
        Assert.NotNull(archive.GetEntry("customPart/Records.bin"));
    }

    /// <summary>
    /// 验证严格 OOXML 透视关系可解析到完整非标准部件。
    /// </summary>
    [Fact]
    public void Find_ShouldFollowStrictOoxmlRelationshipUri()
    {
        var bytes = CreateRelationshipPackage("../../customPart/Records.bin", includeTarget: true, strict: true);
        using var stream = new MemoryStream(bytes);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        var parts = ExcelPivotTemplateParts.Find(archive);

        Assert.Equal(new[] { "customPart/Records.bin" }, parts);
        Assert.NotNull(archive.GetEntry("customPart/Records.bin"));
        Assert.True(stream.CanRead);
    }

    /// <summary>
    /// 验证只有三个定义的透视内容类型会被识别。
    /// </summary>
    /// <param name="suffix">待检查的内容类型后缀。</param>
    /// <param name="expected">是否应识别为透视部件。</param>
    [Theory]
    [InlineData("pivotTable+xml", true)]
    [InlineData("pivotCacheDefinition+xml", true)]
    [InlineData("pivotCacheRecords+xml", true)]
    [InlineData("pivotTable+json", false)]
    [InlineData("pivotCacheDefinition+xml-extra", false)]
    public void IsPivotPart_ShouldMatchExactPivotContentTypes(string suffix, bool expected)
    {
        Assert.Equal(expected, ExcelPivotTemplateParts.IsPivotPart("custom/item.bin", ContentTypePrefix + suffix));
    }

    /// <summary>
    /// 验证不存在的关系目标作为无效包拒绝。
    /// </summary>
    [Fact]
    public void Find_MissingRelationshipTarget_ShouldThrowInvalidDataException()
    {
        var bytes = CreateRelationshipPackage("../../customPart/Missing.xml", includeTarget: false);
        using var stream = new MemoryStream(bytes);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        Assert.Throws<InvalidDataException>(() => ExcelPivotTemplateParts.Find(archive));
        Assert.True(stream.CanRead);
    }

    /// <summary>
    /// 验证外部透视关系作为无效包拒绝。
    /// </summary>
    [Fact]
    public void Find_ExternalRelationship_ShouldThrowInvalidDataException()
    {
        var bytes = CreateRelationshipPackage("https://example.invalid/records.xml", includeTarget: false,
            targetMode: "External");
        using var stream = new MemoryStream(bytes);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        Assert.Throws<InvalidDataException>(() => ExcelPivotTemplateParts.Find(archive));
    }

    /// <summary>
    /// 验证 ZIP 中重复部件名称作为无效包拒绝。
    /// </summary>
    [Fact]
    public void Find_DuplicateZipEntry_ShouldThrowInvalidDataException()
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "custom/duplicate.bin", "first");
            WriteEntry(archive, "custom/duplicate.bin", "second");
        }
        stream.Position = 0;
        using var input = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        Assert.Throws<InvalidDataException>(() => ExcelPivotTemplateParts.Find(input));
        Assert.True(stream.CanRead);
    }

    /// <summary>
    /// 验证 DTD 元数据拒绝并遵守调用方资源所有权。
    /// </summary>
    [Fact]
    public void Find_DtdMetadata_ShouldThrowInvalidDataExceptionAndLeaveArchiveOpen()
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "[Content_Types].xml",
                "<!DOCTYPE Types [<!ENTITY ct SYSTEM 'file:///pivot.dtd'>]>" +
                "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>" +
                "<Override PartName='/custom/part.bin' ContentType='&ct;'/></Types>");
        }
        stream.Position = 0;
        using var input = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        Assert.Throws<InvalidDataException>(() => ExcelPivotTemplateParts.Find(input));
        Assert.NotNull(input.GetEntry("[Content_Types].xml"));
        Assert.True(stream.CanRead);
        Assert.True(stream.CanWrite);
    }

    /// <summary>
    /// 验证预取消不关闭或移动调用方包流。
    /// </summary>
    [Fact]
    public void Find_PreCanceled_ShouldLeaveArchiveAndStreamOpen()
    {
        var bytes = CreatePivotWorkbook();
        using var stream = new MemoryStream(bytes);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => ExcelPivotTemplateParts.Find(archive, cancellation.Token));
        Assert.NotNull(archive.GetEntry("xl/workbook.xml"));
        Assert.True(stream.CanRead);
        Assert.True(stream.CanWrite);
    }

    /// <summary>
    /// 验证公开同步和异步导出仍拒绝含透视的模板并保全输入输出。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导出。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Export_PivotTemplate_ShouldRejectAndPreserveStreams(bool asynchronous)
    {
        var original = CreatePivotWorkbook();
        using var template = new MemoryStream(original);
        using var output = new MemoryStream();
        var sentinel = new byte[] { 13, 29, 47, 83 };
        output.Write(sentinel);
        output.Position = 0;
        var request = TemplateRequest(template);
        var exporter = new NpoiExcelExporter();

        var error = asynchronous
            ? await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() => exporter.ExportAsync(request, output))
            : Assert.Throws<BingOfficesUnsupportedFeatureException>(() => exporter.Export(request, output));

        Assert.Equal(BingOfficesStage.Preflight, error.Stage);
        Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, error.Code);
        Assert.Equal("NPOI", error.Provider);
        Assert.Equal(original, template.ToArray());
        Assert.Equal(sentinel, output.ToArray());
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证非标准路径仍通过精确内容类型触发 NPOI 模板拒绝。
    /// </summary>
    [Fact]
    public void Export_RenamedPivotTemplate_ShouldRejectByContentType()
    {
        var original = CreatePivotWorkbook();
        var renamed = RenameParts(original, new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["xl/pivotTables/pivotTable1.xml"] = "vendor/report.xml",
            ["xl/pivotCache/pivotCacheDefinition1.xml"] = "vendor/cache.xml",
            ["xl/pivotCache/pivotCacheRecords1.xml"] = "vendor/records.xml"
        }, out _);
        using var template = new MemoryStream(renamed);
        using var output = new MemoryStream();
        var sentinel = new byte[] { 2, 3, 5, 7 };
        output.Write(sentinel);
        output.Position = 0;

        var error = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new NpoiExcelExporter().Export(TemplateRequest(template), output));

        Assert.Equal(BingOfficesStage.Preflight, error.Stage);
        Assert.Equal("NPOI", error.Provider);
        Assert.Equal(renamed, template.ToArray());
        Assert.Equal(sentinel, output.ToArray());
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证 NPOI XLS/HSSF 透视导出仍在写入前拒绝且保全目标。
    /// </summary>
    [Fact]
    public void Export_XlsPivot_ShouldRejectBeforeWriting()
    {
        var request = ExcelExport.Workbook(builder => builder.Format(ExcelFormat.Xls)
            .AddSheet("Data", Rows()).PivotTable(Definition("Summary")));
        using var output = new MemoryStream();
        var sentinel = new byte[] { 31, 37, 41 };
        output.Write(sentinel);
        output.Position = 0;

        var error = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new NpoiExcelExporter().Export(request, output));

        Assert.Equal(BingOfficesStage.Preflight, error.Stage);
        Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, error.Code);
        Assert.Equal("NPOI", error.Provider);
        Assert.Equal(sentinel, output.ToArray());
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 创建包含完整基础透视表的真实 NPOI XLSX 工作簿。
    /// </summary>
    /// <returns>公开 Provider 导出的完整工作簿。</returns>
    private static byte[] CreatePivotWorkbook()
    {
        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(NpoiPivotPublicExportTest.Request(false), output);
        NpoiPivotPublicExportTest.AssertWorkbook(output.ToArray(), advanced: false);
        return output.ToArray();
    }

    /// <summary>
    /// 创建模板中已有透视表时用于验证拒绝的请求。
    /// </summary>
    /// <param name="template">调用方拥有的模板流。</param>
    /// <returns>包含一个新透视定义的公开请求。</returns>
    private static ExcelWorkbookExportRequest TemplateRequest(Stream template) =>
        ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", Rows()).PivotTable(Definition("NewSummary")));

    /// <summary>
    /// 创建基础透视定义。
    /// </summary>
    /// <param name="target">透视目标工作表名称。</param>
    /// <returns>包含完整来源范围的透视定义。</returns>
    private static ExcelPivotTableDefinition Definition(string target) => new()
    {
        Name = "TemplatePartsPivot", SourceSheet = "Data", TargetSheet = target,
        SourceRange = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 4 },
        RowField = "Region", ValueField = "Amount"
    };

    /// <summary>
    /// 创建完整采购来源记录。
    /// </summary>
    /// <returns>三条业务来源行。</returns>
    private static NpoiPivotExportRow[] Rows() => new[]
    {
        new NpoiPivotExportRow { Region = "North", Amount = 10, State = "Open", Channel = "Retail", Code = "A" },
        new NpoiPivotExportRow { Region = "North", Amount = 30, State = "Closed", Channel = "Retail", Code = "B" },
        new NpoiPivotExportRow { Region = "South", Amount = 20, State = "Open", Channel = "Wholesale", Code = "C" }
    };

    /// <summary>
    /// 将透视主部件移动到非标准路径并同步更新其 OPC 引用。
    /// </summary>
    /// <param name="source">真实 Provider 输出。</param>
    /// <param name="requestedMapping">主部件的新路径映射。</param>
    /// <param name="renamedRelationships">随主部件迁移的关系部件名称。</param>
    /// <returns>路径和关系均已更新的工作簿。</returns>
    private static byte[] RenameParts(byte[] source, IReadOnlyDictionary<string, string> requestedMapping,
        out HashSet<string> renamedRelationships)
    {
        var entries = ReadEntries(source);
        var mapping = requestedMapping.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
        renamedRelationships = new HashSet<string>(StringComparer.Ordinal);
        foreach (var pair in requestedMapping)
        {
            var oldRelationship = RelationshipPart(pair.Key);
            var newRelationship = RelationshipPart(pair.Value);
            if (entries.ContainsKey(oldRelationship))
            {
                mapping.Add(oldRelationship, newRelationship);
                renamedRelationships.Add(oldRelationship);
            }
        }

        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var pair in entries)
            {
                var name = mapping.GetValueOrDefault(pair.Key, pair.Key);
                var contents = pair.Value;
                if (pair.Key == "[Content_Types].xml")
                    contents = RewriteContentTypes(contents, requestedMapping);
                else if (pair.Key.EndsWith(".rels", StringComparison.Ordinal))
                    contents = RewriteRelationships(pair.Key, contents, mapping);
                var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
                using var stream = entry.Open();
                stream.Write(contents);
            }
        }
        return output.ToArray();
    }

    /// <summary>
    /// 读取 ZIP 中每个文件的原始内容。
    /// </summary>
    /// <param name="bytes">源 XLSX 字节。</param>
    /// <returns>按部件路径索引的内容快照。</returns>
    private static Dictionary<string, byte[]> ReadEntries(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entries = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in archive.Entries)
        {
            using var content = entry.Open();
            using var copy = new MemoryStream();
            content.CopyTo(copy);
            entries.Add(entry.FullName, copy.ToArray());
        }
        return entries;
    }

    /// <summary>
    /// 计算 OPC 部件对应的关系部件路径。
    /// </summary>
    /// <param name="partName">拥有关系的部件路径。</param>
    /// <returns>关系部件路径。</returns>
    private static string RelationshipPart(string partName)
    {
        var separator = partName.LastIndexOf('/');
        return separator < 0
            ? $"_rels/{partName}.rels"
            : $"{partName[..(separator + 1)]}_rels/{partName[(separator + 1)..]}.rels";
    }

    /// <summary>
    /// 更新内容类型覆盖项的部件名称。
    /// </summary>
    /// <param name="bytes">内容类型 XML。</param>
    /// <param name="mapping">旧路径到新路径的映射。</param>
    /// <returns>修改后的 XML 字节。</returns>
    private static byte[] RewriteContentTypes(byte[] bytes, IReadOnlyDictionary<string, string> mapping)
    {
        using var stream = new MemoryStream(bytes);
        var document = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";
        foreach (var item in document.Root.Elements(ns + "Override"))
        {
            var partName = ((string)item.Attribute("PartName"))?.TrimStart('/');
            if (partName != null && mapping.TryGetValue(partName, out var renamed))
                item.SetAttributeValue("PartName", "/" + renamed);
        }
        return Serialize(document);
    }

    /// <summary>
    /// 重写关系目标，使其继续指向迁移后的实际部件。
    /// </summary>
    /// <param name="relationshipPart">当前关系部件路径。</param>
    /// <param name="bytes">关系 XML。</param>
    /// <param name="mapping">旧路径到新路径的映射。</param>
    /// <returns>修改后的 XML 字节。</returns>
    private static byte[] RewriteRelationships(string relationshipPart, byte[] bytes,
        IReadOnlyDictionary<string, string> mapping)
    {
        using var stream = new MemoryStream(bytes);
        var document = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";
        var owner = RelationshipOwner(relationshipPart);
        var ownerDirectory = owner.Contains('/') ? owner[..(owner.LastIndexOf('/') + 1)] : string.Empty;
        var baseUri = new Uri("http://package.invalid/" + ownerDirectory, UriKind.Absolute);
        foreach (var item in document.Root.Elements(ns + "Relationship"))
        {
            if (string.Equals((string)item.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
                continue;
            var target = (string)item.Attribute("Target");
            if (string.IsNullOrWhiteSpace(target))
                continue;
            var resolved = new Uri(baseUri, target).AbsolutePath.TrimStart('/');
            if (mapping.TryGetValue(resolved, out var renamed))
                item.SetAttributeValue("Target", baseUri.MakeRelativeUri(
                    new Uri("http://package.invalid/" + renamed, UriKind.Absolute)).ToString());
        }
        return Serialize(document);
    }

    /// <summary>
    /// 从 OPC 关系部件路径还原其拥有者路径。
    /// </summary>
    /// <param name="relationshipPart">关系部件名称。</param>
    /// <returns>拥有关系的源部件路径。</returns>
    private static string RelationshipOwner(string relationshipPart)
    {
        if (relationshipPart == "_rels/.rels")
            return string.Empty;
        var separator = relationshipPart.LastIndexOf("/_rels/", StringComparison.Ordinal);
        if (separator < 0)
            throw new InvalidDataException("关系部件路径无效。");
        return relationshipPart[..(separator + 1)] + relationshipPart[(separator + 7)..^5];
    }

    /// <summary>
    /// 将 XML 文档序列化为 UTF-8 字节。
    /// </summary>
    /// <param name="document">待序列化文档。</param>
    /// <returns>XML 字节。</returns>
    private static byte[] Serialize(XDocument document)
    {
        using var output = new MemoryStream();
        using (var writer = System.Xml.XmlWriter.Create(output, new System.Xml.XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = document.Declaration == null
        }))
            document.Save(writer);
        return output.ToArray();
    }

    /// <summary>
    /// 创建仅包含一条透视关系的最小 OPC 包。
    /// </summary>
    /// <param name="target">关系目标。</param>
    /// <param name="includeTarget">是否写入目标部件。</param>
    /// <param name="targetMode">可选目标模式。</param>
    /// <returns>最小关系包字节。</returns>
    private static byte[] CreateRelationshipPackage(string target, bool includeTarget, string targetMode = null,
        bool strict = false)
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            var mode = targetMode == null ? string.Empty : $" TargetMode='{targetMode}'";
            var relationshipPrefix = strict ? StrictRelationshipPrefix : RelationshipPrefix;
            WriteEntry(archive, "xl/worksheets/_rels/sheet1.xml.rels",
                $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
                $"<Relationship Id='rId1' Type='{relationshipPrefix}pivotCacheRecords' Target='{target}'{mode}/></Relationships>");
            if (includeTarget)
                WriteEntry(archive, "customPart/Records.bin", "records");
        }
        return stream.ToArray();
    }

    /// <summary>
    /// 写入 UTF-8 ZIP 条目。
    /// </summary>
    /// <param name="archive">目标 ZIP 包。</param>
    /// <param name="name">部件名称。</param>
    /// <param name="contents">文本内容。</param>
    private static void WriteEntry(ZipArchive archive, string name, string contents)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(contents);
    }
}
