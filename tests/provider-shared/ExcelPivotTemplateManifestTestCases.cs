using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Internals;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Exports;
using Xunit;

namespace Bing.Offices.Tests.PivotTemplates;

/// <summary>
/// 验证真实透视模板的依赖清单。
/// </summary>
internal static class ExcelPivotTemplateManifestTestCases
{
    /// <summary>
    /// 验证磁盘快照、真实外围异步读取和失败清理。
    /// </summary>
    /// <param name="scenario">资源场景。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    /// <param name="bytes">完整模板。</param>
    /// <param name="fields">来源字段。</param>
    /// <param name="provider">提供程序名称。</param>
    internal static async Task VerifySnapshotAsync(string scenario, bool asynchronous, byte[] bytes,
        string[] fields, string provider)
    {
        if (scenario == "malformed")
            bytes = new byte[] { 1, 2, 3 };
        var existing = Directory.GetFiles(Path.GetTempPath(), "bing-offices-pivot-template-*.tmp").OrderBy(path => path).ToArray();
        var inputPath = Path.Combine(Path.GetTempPath(), "bing-tests-pivot-source-" + Guid.NewGuid().ToString("N") + ".tmp");
        using var file = new FileStream(inputPath, FileMode.CreateNew, FileAccess.ReadWrite,
            FileShare.None, 81920, FileOptions.Asynchronous | FileOptions.DeleteOnClose);
        await file.WriteAsync(bytes.AsMemory()).ConfigureAwait(false);
        await file.FlushAsync().ConfigureAwait(false);
        using var cancellation = new CancellationTokenSource();
        using var source = new ExcelPivotTemplateReadStream(file,
            scenario == "copy-cancel" ? cancellation : null, scenario == "copy-fail");
        source.Position = 7;
        if (scenario == "pre-cancel")
            cancellation.Cancel();
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(source, leaveOpen: true)
            .UpdatePivotTable(new ExcelPivotTableDefinition
            {
                Name = scenario == "unknown-update" ? "Missing" : "ByGroupOne",
                SourceSheet = "Data", TargetSheet = "Summary", RowField = fields[0], ValueField = fields[1],
                SourceRange = new ExcelRangeDefinition { EndRow = 2, EndColumn = fields.Length - 1 }
            }));
        var limits = scenario == "budget" ? new ExcelResourceLimits { MaxZipEntries = 1 } : new ExcelResourceLimits();
        async Task<ExcelPivotTemplateSnapshot> Create() => asynchronous
            ? await ExcelPivotTemplateSnapshot.CreateAsync(request, limits, provider, cancellation.Token).ConfigureAwait(false)
            : ExcelPivotTemplateSnapshot.Create(request, limits, provider, cancellation.Token);
        if (scenario == "success")
        {
            string stagedPath;
            using (var snapshot = await Create().ConfigureAwait(false))
            {
                var staged = Assert.IsType<FileStream>(snapshot.Source);
                stagedPath = staged.Name;
                Assert.True(staged.IsAsync);
                Assert.Equal(0, staged.Position);
                using var copied = new MemoryStream();
                await staged.CopyToAsync(copied).ConfigureAwait(false);
                Assert.Equal(bytes, copied.ToArray());
                Assert.Equal("ByGroupOne", Assert.Single(snapshot.Plan.Updates).Entry.Name);
                Assert.Equal("ByGroupTwo", Assert.Single(snapshot.Plan.PreservedEntries).Name);
                Assert.Equal(new[] { "Data", "Summary", "Summary2" }, snapshot.Plan.Manifest.SheetNames);
            }
            Assert.False(File.Exists(stagedPath));
        }
        else if (scenario is "pre-cancel" or "copy-cancel")
        {
            var error = await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
                await Create().ConfigureAwait(false));
            Assert.Equal(cancellation.Token, error.CancellationToken);
        }
        else if (scenario == "copy-fail")
            await Assert.ThrowsAsync<IOException>(async () => await Create().ConfigureAwait(false));
        else if (scenario == "unknown-update")
        {
            var error = await Assert.ThrowsAsync<BingOfficesConfigurationException>(async () =>
                await Create().ConfigureAwait(false));
            Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, error.Code);
            Assert.Equal(BingOfficesStage.Preflight, error.Stage);
        }
        else if (scenario == "budget")
        {
            var error = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(async () =>
                await Create().ConfigureAwait(false));
            Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, error.Code);
            Assert.Equal(provider, error.Provider);
            Assert.Equal(BingOfficesOperation.Export, error.Operation);
            Assert.Equal(BingOfficesStage.Preflight, error.Stage);
        }
        else
        {
            var error = await Assert.ThrowsAsync<BingOfficesExportException>(async () =>
                await Create().ConfigureAwait(false));
            Assert.Equal(provider, error.Provider);
            Assert.Equal(BingOfficesOperation.Export, error.Operation);
            Assert.Equal(BingOfficesStage.Preflight, error.Stage);
        }
        Assert.Equal(7, source.Position);
        Assert.False(source.WasDisposed);
        Assert.True(file.CanRead);
        if (scenario != "pre-cancel")
        {
            Assert.True(asynchronous ? source.AsyncReads > 0 : source.SyncReads > 0);
            Assert.Equal(0, asynchronous ? source.SyncReads : source.AsyncReads);
        }
        file.Position = 0;
        using var unchanged = new MemoryStream();
        await file.CopyToAsync(unchanged).ConfigureAwait(false);
        Assert.Equal(bytes, unchanged.ToArray());
        Assert.Equal(existing, Directory.GetFiles(Path.GetTempPath(), "bing-offices-pivot-template-*.tmp").OrderBy(path => path));
    }

    /// <summary>
    /// 验证已有透视的精确定位与更新边界。
    /// </summary>
    /// <param name="scenario">预检场景。</param>
    /// <param name="bytes">完整模板。</param>
    /// <param name="fields">来源字段。</param>
    internal static void VerifyPlan(string scenario, byte[] bytes, string[] fields)
    {
        using var source = new MemoryStream(bytes);
        source.Position = 7;
        var manifest = ExcelPivotTemplateManifest.Read(source);
        var definition = new ExcelPivotTableDefinition
        {
            Name = scenario == "unknown-name" ? "Missing" : scenario == "new-target-conflict" ? "NewPivot" : "ByGroupOne",
            SourceSheet = scenario == "unknown-source" ? "Missing" : "Data",
            TargetSheet = scenario is "wrong-sheet" or "new-target-conflict" ? "Summary2"
                : scenario == "new-name-conflict" ? "NewSummary" : "Summary",
            SourceRange = new ExcelRangeDefinition
            {
                StartRow = 0, StartColumn = 0,
                EndRow = scenario == "source-overflow" ? 1048576 : 2,
                EndColumn = fields.Length - 1
            },
            TargetRow = scenario == "row-overflow" ? 1048576 : 1,
            TargetColumn = scenario == "column-overflow" ? 16384 : 2,
            RowField = fields[0], ValueField = fields[1]
        };
        var request = ExcelExport.Workbook(builder =>
        {
            builder.UseTemplate(source, leaveOpen: true).PreserveTemplatePivotTables();
            if (scenario is "new-name-conflict" or "new-target-conflict")
                builder.AddSheet("Data", Array.Empty<object>()).PivotTable(definition);
            else if (scenario != "preserve-only")
                builder.UpdatePivotTable(definition);
        });
        if (scenario is "valid" or "preserve-only")
        {
            var plan = ExcelPivotTemplatePlan.Create(request, manifest, "TestProvider");
            Assert.Same(manifest, plan.Manifest);
            if (scenario == "valid")
            {
                var update = Assert.Single(plan.Updates);
                Assert.Same(manifest.Entries[0], update.Entry);
                Assert.Same(Assert.Single(request.PivotTableUpdates), update.Definition);
                Assert.Equal("ByGroupTwo", Assert.Single(plan.PreservedEntries).Name);
            }
            else
            {
                Assert.Empty(plan.Updates);
                Assert.Equal(new[] { "ByGroupOne", "ByGroupTwo" }, plan.PreservedEntries.Select(item => item.Name));
            }
            Assert.Equal(new[] { "Data", "Summary", "Summary2" }, manifest.SheetNames);
            Assert.Throws<NotSupportedException>(() => ((IList<ExcelPivotTemplateUpdate>)plan.Updates).Clear());
        }
        else if (scenario == "canceled")
        {
            var token = new CancellationToken(true);
            var error = Assert.Throws<OperationCanceledException>(() =>
                ExcelPivotTemplatePlan.Create(request, manifest, "TestProvider", token));
            Assert.Equal(token, error.CancellationToken);
        }
        else
        {
            var error = Assert.Throws<BingOfficesConfigurationException>(() =>
                ExcelPivotTemplatePlan.Create(request, manifest, "TestProvider"));
            Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, error.Code);
            Assert.Equal(BingOfficesStage.Preflight, error.Stage);
            Assert.Equal(BingOfficesOperation.Configuration, error.Operation);
            Assert.Equal("Core", error.Provider);
            Assert.Equal(2, manifest.Entries.Count);
            Assert.Equal("ByGroupOne", manifest.Entries[0].Name);
        }
        Assert.Equal(7, source.Position);
        Assert.Equal(bytes, source.ToArray());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证模板资源预检和输入生命周期。
    /// </summary>
    /// <param name="scenario">资源场景。</param>
    /// <param name="bytes">真实透视模板。</param>
    /// <param name="provider">提供程序名称。</param>
    internal static void VerifyResource(string scenario, byte[] bytes, string provider)
    {
        if (scenario == "renamed")
            bytes = Pack(Rename(Unpack(bytes), true));
        else if (scenario is "relationships" or "relationship-depth")
        {
            var parts = Unpack(bytes);
            var relationships = Xml(parts, "_rels/.rels");
            XNamespace relationshipNamespace =
                "http://schemas.openxmlformats.org/package/2006/relationships";
            if (scenario == "relationships")
            {
                relationships.Root.Add(new XElement(relationshipNamespace + "Relationship",
                    new XAttribute("Id", "rIdResourceBudget"),
                    new XAttribute("Type", "urn:bing-offices:resource-budget"),
                    new XAttribute("Target", new string('x', 65536))));
            }
            else
            {
                var nested = new XElement(relationshipNamespace + "Nested");
                var current = nested;
                for (var index = 0; index < 64; index++)
                {
                    var child = new XElement(relationshipNamespace + "Nested");
                    current.Add(child);
                    current = child;
                }
                relationships.Root.Add(new XElement(relationshipNamespace + "Extension",
                    nested));
            }
            Save(parts, "_rels/.rels", relationships);
            bytes = Pack(parts);
        }
        var limits = scenario switch
        {
            "input" => new ExcelResourceLimits { MaxInputBytes = 1 },
            "entries" => new ExcelResourceLimits { MaxZipEntries = 1 },
            "entry" => new ExcelResourceLimits { MaxZipEntryUncompressedBytes = 1 },
            "total" => new ExcelResourceLimits { MaxZipTotalUncompressedBytes = 1 },
            "ratio" => new ExcelResourceLimits { MaxZipCompressionRatio = 0.0001 },
            "characters" => new ExcelResourceLimits { MaxXmlCharacters = 1 },
            "depth" => new ExcelResourceLimits { MaxXmlDepth = 1 },
            "relationships" => new ExcelResourceLimits { MaxXmlCharacters = 32768 },
            "relationship-depth" => new ExcelResourceLimits { MaxXmlDepth = 64 },
            _ => new ExcelResourceLimits()
        };
        using var source = new MemoryStream(bytes);
        source.Position = 7;
        if (scenario == "renamed")
        {
            var manifest = ExcelPivotTemplateManifest.Read(source, limits, provider);
            Assert.Equal("book.xml", manifest.WorkbookPart);
            Assert.Equal(new[] { "ByGroupOne", "ByGroupTwo" }, manifest.Entries.Select(item => item.Name));
            Assert.All(manifest.Entries, item => Assert.Equal("Data", item.Cache.SourceSheet));
            Assert.All(manifest.Caches, item => Assert.StartsWith("business/storage/", item.DefinitionPart));
        }
        else
        {
            var error = Assert.Throws<BingOfficesResourceLimitException>(() =>
                ExcelPivotTemplateManifest.Read(source, limits, provider));
            Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, error.Code);
            Assert.Equal(BingOfficesStage.Preflight, error.Stage);
            Assert.Equal(BingOfficesOperation.Import, error.Operation);
            Assert.Equal(provider, error.Provider);
            if (scenario is "relationships" or "relationship-depth")
                Assert.Contains("_rels/.rels", error.Message);
        }
        Assert.Equal(7, source.Position);
        Assert.Equal(bytes, source.ToArray());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 执行实际模板清单职责。
    /// </summary>
    /// <param name="scenario">职责场景。</param>
    /// <param name="bytes">公开导出的两个透视表模板。</param>
    /// <param name="fields">完整来源字段。</param>
    /// <param name="sourceReference">完整来源区域。</param>
    internal static void Verify(string scenario, byte[] bytes, string[] fields, string sourceReference)
    {
        var parts = Unpack(bytes);
        var workbook = Xml(parts, "xl/workbook.xml");
        var cacheReferences = workbook.Root.Elements().Single(item => item.Name.LocalName == "pivotCaches").Elements().ToArray();
        var pivotNames = parts.Keys.Where(name => name.StartsWith("xl/pivotTables/", StringComparison.Ordinal)
            && name.EndsWith(".xml", StringComparison.Ordinal)).OrderBy(name => name).ToArray();
        Assert.Equal(2, pivotNames.Length);
        if (scenario == "shared")
            ShareCaches(parts, workbook, cacheReferences, pivotNames);
        else if (scenario is "renamed" or "root-owner")
            parts = Rename(parts, scenario == "root-owner");
        else if (scenario == "unknown-cache")
            Edit(parts, pivotNames[0], xml => xml.Root.SetAttributeValue("cacheId", uint.MaxValue));
        else if (scenario == "duplicate-name")
            Edit(parts, pivotNames[1], xml => xml.Root.SetAttributeValue("name", "ByGroupOne"));
        else if (scenario == "duplicate-cache")
        {
            cacheReferences[0].Parent.Add(new XElement(cacheReferences[0]));
            Save(parts, "xl/workbook.xml", workbook);
        }
        else if (scenario == "duplicate-relationship")
            Edit(parts, "xl/_rels/workbook.xml.rels", xml => xml.Root.Add(new XElement(xml.Root.Elements().First())));
        else if (scenario == "orphan-pivot")
        {
            foreach (var name in parts.Keys.Where(name => name.StartsWith("xl/worksheets/", StringComparison.Ordinal)
                && name.EndsWith(".xml", StringComparison.Ordinal)).ToArray())
                Edit(parts, name, xml => xml.Descendants().Where(item => item.Name.LocalName == "pivotTablePart").Remove());
            foreach (var name in parts.Keys.Where(name => name.StartsWith("xl/worksheets/", StringComparison.Ordinal)
                && name.EndsWith(".rels", StringComparison.Ordinal)).ToArray())
                Edit(parts, name, xml => xml.Root.Elements().Where(item =>
                    ((string)item.Attribute("Type")).EndsWith("/pivotTable", StringComparison.Ordinal)).Remove());
        }
        else if (scenario == "missing-records")
            parts.Remove(parts.Keys.First(name => name.Contains("pivotCacheRecords") && name.EndsWith(".xml", StringComparison.Ordinal)));
        else if (scenario == "missing-cache-link")
            Edit(parts, RelationshipPart(pivotNames[0]), xml =>
                xml.Root.Elements().Where(item => ((string)item.Attribute("Type")).EndsWith("/pivotCacheDefinition", StringComparison.Ordinal)).Remove());
        else if (scenario == "external")
            Edit(parts, "xl/_rels/workbook.xml.rels", xml =>
            {
                var item = xml.Root.Elements().First(value => ((string)value.Attribute("Type")).EndsWith("/pivotCacheDefinition", StringComparison.Ordinal));
                item.SetAttributeValue("TargetMode", "External");
                item.SetAttributeValue("Target", "https://example.invalid/cache.xml");
            });
        else if (scenario == "dtd")
            parts["xl/workbook.xml"] = Encoding.UTF8.GetBytes("<!DOCTYPE workbook [<!ENTITY x 'invalid'>]>" + workbook.ToString(SaveOptions.DisableFormatting));
        else if (scenario is "unknown-source" or "invalid-range" or "reversed-range" or "overflow-range" or "missing-range")
            Edit(parts, parts.Keys.First(name => name.Contains("pivotCacheDefinition") && name.EndsWith(".xml", StringComparison.Ordinal)), xml =>
            {
                var source = xml.Descendants().Single(item => item.Name.LocalName == "worksheetSource");
                if (scenario == "unknown-source")
                    source.SetAttributeValue("sheet", "Missing");
                else
                    source.SetAttributeValue("ref", scenario == "invalid-range" ? "A:A"
                        : scenario == "reversed-range" ? "B3:A1" : scenario == "overflow-range" ? "A1:XFE3" : null);
            });
        else if (scenario == "wrong-sheet-root")
            Edit(parts, "xl/worksheets/sheet1.xml", xml => xml.Root.Name = xml.Root.Name.Namespace + "workbook");
        else if (scenario == "wrong-worksheet-type")
            Edit(parts, "xl/_rels/workbook.xml.rels", xml => xml.Root.Elements()
                .First(item => ((string)item.Attribute("Type")).EndsWith("/worksheet", StringComparison.Ordinal))
                .SetAttributeValue("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"));
        else if (scenario == "duplicate-owner")
        {
            var link = Xml(parts, "xl/worksheets/_rels/sheet2.xml.rels").Root.Elements()
                .First(item => ((string)item.Attribute("Type")).EndsWith("/pivotTable", StringComparison.Ordinal));
            const string ownerRelationships = "xl/worksheets/_rels/sheet1.xml.rels";
            if (parts.ContainsKey(ownerRelationships))
                Edit(parts, ownerRelationships, xml => xml.Root.Add(new XElement(link)));
            else
                Save(parts, ownerRelationships, new XDocument(new XElement(link.Name.Namespace + "Relationships", new XElement(link))));
        }
        else if (scenario == "mismatched-cache-link")
        {
            var wbLinks = Xml(parts, "xl/_rels/workbook.xml.rels");
            var oldLink = wbLinks.Root.Elements().First(item => ((string)item.Attribute("Type")).EndsWith("/pivotCacheDefinition", StringComparison.Ordinal));
            var oldPart = new Uri(new Uri("http://fixture/xl/workbook.xml"), (string)oldLink.Attribute("Target")).AbsolutePath.TrimStart('/');
            var newPart = oldPart[..(oldPart.LastIndexOf('/') + 1)] + "unrelated-cache.xml";
            parts.Add(newPart, parts[oldPart]);
            parts.Add(RelationshipPart(newPart), parts[RelationshipPart(oldPart)]);
            var id = "rIdUnrelatedCache";
            var newLink = new XElement(oldLink);
            newLink.SetAttributeValue("Id", id);
            newLink.SetAttributeValue("Target", new Uri("http://fixture/xl/workbook.xml").MakeRelativeUri(new Uri("http://fixture/" + newPart)).ToString());
            wbLinks.Root.Add(newLink);
            Save(parts, "xl/_rels/workbook.xml.rels", wbLinks);
            var newCache = new XElement(cacheReferences[0]);
            newCache.SetAttributeValue("cacheId", uint.MaxValue - 1);
            newCache.Attributes().Single(attribute => attribute.Name.LocalName == "id").Value = id;
            cacheReferences[0].Parent.Add(newCache);
            Save(parts, "xl/workbook.xml", workbook);
            Edit(parts, "[Content_Types].xml", xml =>
            {
                var definition = new XElement(xml.Root.Elements().First(item => (string)item.Attribute("PartName") == "/" + oldPart));
                definition.SetAttributeValue("PartName", "/" + newPart);
                xml.Root.Add(definition);
            });
            Edit(parts, RelationshipPart(pivotNames[0]), xml => xml.Root.Elements().Single().SetAttributeValue("Target",
                new Uri("http://fixture/" + pivotNames[0]).MakeRelativeUri(new Uri("http://fixture/" + newPart)).ToString()));
        }
        else if (scenario == "strict")
            foreach (var name in parts.Keys.Where(name => name.EndsWith(".xml", StringComparison.Ordinal)
                || name.EndsWith(".rels", StringComparison.Ordinal)).ToArray())
                parts[name] = Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(parts[name])
                    .Replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        "http://purl.oclc.org/ooxml/officeDocument/relationships", StringComparison.Ordinal)
                    .Replace("http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                        "http://purl.oclc.org/ooxml/spreadsheetml/main", StringComparison.Ordinal));

        using var stream = new MemoryStream(Pack(parts));
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        stream.Position = 7;
        if (scenario == "canceled")
        {
            var token = new CancellationToken(true);
            var error = Assert.Throws<OperationCanceledException>(() => ExcelPivotTemplateManifest.Read(stream, token));
            Assert.Equal(token, error.CancellationToken);
            Assert.Equal(7, stream.Position);
        }
        else if (scenario is not ("normal" or "shared" or "renamed" or "root-owner" or "strict"))
        {
            Assert.Throws<InvalidDataException>(() => ExcelPivotTemplateManifest.Read(stream));
            Assert.Equal(7, stream.Position);
        }
        else
        {
            var manifest = ExcelPivotTemplateManifest.Read(stream);
            Assert.Equal(7, stream.Position);
            Assert.Equal(scenario == "root-owner" ? "book.xml" : scenario == "renamed" ? "business/book.xml" : "xl/workbook.xml", manifest.WorkbookPart);
            Assert.Equal(new[] { "ByGroupOne", "ByGroupTwo" }, manifest.Entries.Select(item => item.Name));
            Assert.Equal(new[] { "Summary", "Summary2" }, manifest.Entries.Select(item => item.SheetName));
            Assert.All(manifest.Entries, item =>
            {
                Assert.Equal("Data", item.Cache.SourceSheet);
                Assert.Equal(sourceReference, item.Cache.SourceReference);
                Assert.Equal("worksheet", item.Cache.SourceType);
                Assert.Null(item.Cache.SourceName);
                Assert.Equal(fields, item.Cache.Fields);
                Assert.Contains(item.PartName, item.Cache.Consumers);
                Assert.Contains(item.Cache, manifest.Caches);
                Assert.Equal(item.PartName, manifest.Relationships[item.SheetPart].Single(link => link.Id == item.RelationshipId).PartName);
                Assert.Equal(item.Cache.DefinitionPart, manifest.Relationships[manifest.WorkbookPart].Single(link => link.Id == item.Cache.RelationshipId).PartName);
                Assert.Equal(item.Cache.RecordsPart, manifest.Relationships[item.Cache.DefinitionPart].Single(link => link.Id == item.Cache.RecordsRelationshipId).PartName);
                Assert.Equal((string)Xml(parts, item.PartName).Root.Elements().Single(element => element.Name.LocalName == "location").Attribute("ref"), item.TargetReference);
                Assert.StartsWith("C2", item.TargetReference);
            });
            if (scenario == "shared")
            {
                var cache = Assert.Single(manifest.Caches);
                Assert.Equal(manifest.Entries.Select(item => item.PartName), cache.Consumers);
                Assert.All(manifest.Entries, item => Assert.Same(cache, item.Cache));
            }
            if (scenario is "renamed" or "root-owner")
            {
                Assert.All(manifest.Entries, item => Assert.StartsWith("business/native/", item.PartName));
                Assert.All(manifest.Caches, item => Assert.StartsWith("business/storage/", item.DefinitionPart));
                Assert.All(manifest.Caches, item => Assert.StartsWith("business/storage/", item.RecordsPart));
            }
            Assert.Throws<NotSupportedException>(() => ((IList<ExcelPivotTemplateEntry>)manifest.Entries).Clear());
            Assert.Throws<NotSupportedException>(() => ((IList<string>)manifest.Caches[0].Fields).Clear());
            Assert.Throws<NotSupportedException>(() => ((IList<string>)manifest.Caches[0].Consumers).Clear());
        }
        Assert.NotEmpty(archive.Entries);
        Assert.True(stream.CanRead);
    }

    /// <summary>
    /// 构造两个原生透视共享缓存的模板。
    /// </summary>
    /// <param name="parts">实际包部件。</param>
    /// <param name="workbook">实际工作簿描述。</param>
    /// <param name="references">缓存引用。</param>
    /// <param name="pivots">两个透视部件。</param>
    private static void ShareCaches(Dictionary<string, byte[]> parts, XDocument workbook,
        XElement[] references, string[] pivots)
    {
        var id = (string)references[0].Attribute("cacheId");
        var firstLink = Xml(parts, RelationshipPart(pivots[0])).Root.Elements().Single();
        Edit(parts, pivots[1], xml => xml.Root.SetAttributeValue("cacheId", id));
        Edit(parts, RelationshipPart(pivots[1]), xml => xml.Root.Elements().Single()
            .SetAttributeValue("Target", (string)firstLink.Attribute("Target")));
        if (references.Length == 1)
            return;
        var firstDefinition = new Uri(new Uri("http://fixture/" + pivots[0]), (string)firstLink.Attribute("Target")).AbsolutePath.TrimStart('/');
        var keep = new HashSet<string>(StringComparer.Ordinal) { firstDefinition, RelationshipPart(firstDefinition) };
        var recordsLink = Xml(parts, RelationshipPart(firstDefinition)).Root.Elements().Single();
        keep.Add(new Uri(new Uri("http://fixture/" + firstDefinition), (string)recordsLink.Attribute("Target")).AbsolutePath.TrimStart('/'));
        var removed = parts.Keys.Where(name => name.StartsWith("xl/pivotCache/", StringComparison.Ordinal) && !keep.Contains(name)).ToArray();
        foreach (var name in removed)
            parts.Remove(name);
        references.Skip(1).Remove();
        Save(parts, "xl/workbook.xml", workbook);
        Edit(parts, "xl/_rels/workbook.xml.rels", xml => xml.Root.Elements()
            .Where(item => ((string)item.Attribute("Type")).EndsWith("/pivotCacheDefinition", StringComparison.Ordinal)
                && (string)item.Attribute("Id") != (string)references[0].Attributes().Single(attribute => attribute.Name.LocalName == "id")).Remove());
        Edit(parts, "[Content_Types].xml", xml => xml.Root.Elements()
            .Where(item => removed.Contains(((string)item.Attribute("PartName"))?.TrimStart('/'))).Remove());
    }

    /// <summary>
    /// 将模板关系与部件迁到非标准目录。
    /// </summary>
    /// <param name="parts">实际包部件。</param>
    /// <param name="rootWorkbook">是否将工作簿放在包根目录。</param>
    /// <returns>保留完整关系的非标准模板。</returns>
    private static Dictionary<string, byte[]> Rename(Dictionary<string, byte[]> parts, bool rootWorkbook)
    {
        var payloads = Xml(parts, "[Content_Types].xml").Root.Elements()
            .Where(item => ((string)item.Attribute("ContentType")) is
                "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
                or "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
                or "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml")
            .Select(item => ((string)item.Attribute("PartName")).TrimStart('/')).ToHashSet(StringComparer.Ordinal);
        var map = parts.Keys.Where(name => name == "xl/workbook.xml" || name.StartsWith("xl/worksheets/", StringComparison.Ordinal)
            || payloads.Contains(name))
            .Where(name => !name.EndsWith(".rels", StringComparison.Ordinal))
            .ToDictionary(name => name, name => name == "xl/workbook.xml" ? rootWorkbook ? "book.xml" : "business/book.xml"
                : payloads.Contains(name) && Path.GetFileName(name).StartsWith("pivotTable", StringComparison.Ordinal) ? "business/native/" + Path.GetFileName(name)
                : payloads.Contains(name) ? "business/storage/" + Path.GetFileName(name)
                : "business/sheets/" + Path.GetFileName(name));
        string NewName(string name) => map.TryGetValue(name, out var value) ? value : name;
        var renamed = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var pair in parts)
        {
            var name = NewName(pair.Key);
            var bytes = pair.Value;
            if (pair.Key.EndsWith(".rels", StringComparison.Ordinal))
            {
                var owner = pair.Key == "_rels/.rels" ? string.Empty : pair.Key.Replace("/_rels/", "/")[..^5];
                var newOwner = NewName(owner);
                name = newOwner.Length == 0 ? pair.Key : RelationshipPart(newOwner);
                var xml = Xml(parts, pair.Key);
                foreach (var item in xml.Root.Elements())
                {
                    var target = new Uri(new Uri("http://fixture/" + owner), (string)item.Attribute("Target")).AbsolutePath.TrimStart('/');
                    item.SetAttributeValue("Target", new Uri("http://fixture/" + newOwner)
                        .MakeRelativeUri(new Uri("http://fixture/" + NewName(target))).ToString());
                }
                bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
            }
            else if (name == "[Content_Types].xml")
            {
                var xml = Xml(parts, pair.Key);
                foreach (var item in xml.Root.Elements().Where(item => item.Attribute("PartName") != null))
                    item.SetAttributeValue("PartName", "/" + NewName(((string)item.Attribute("PartName")).TrimStart('/')));
                bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
            }
            renamed.Add(name, bytes);
        }
        return renamed;
    }

    /// <summary>
    /// 获取部件的关系路径。
    /// </summary>
    /// <param name="owner">关系拥有者。</param>
    /// <returns>关系部件路径。</returns>
    private static string RelationshipPart(string owner)
    {
        var index = owner.LastIndexOf('/');
        return owner[..(index + 1)] + "_rels/" + owner[(index + 1)..] + ".rels";
    }

    /// <summary>
    /// 读取全部实际包部件。
    /// </summary>
    /// <param name="bytes">测试工作簿。</param>
    /// <returns>按名称索引的部件。</returns>
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
    /// 写出完整测试包。
    /// </summary>
    /// <param name="parts">包部件。</param>
    /// <returns>完整 ZIP 内容。</returns>
    private static byte[] Pack(Dictionary<string, byte[]> parts)
    {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var part in parts)
            {
                using var stream = archive.CreateEntry(part.Key).Open();
                stream.Write(part.Value);
            }
        return output.ToArray();
    }

    /// <summary>
    /// 读取测试元数据。
    /// </summary>
    /// <param name="parts">包部件。</param>
    /// <param name="name">元数据名称。</param>
    /// <returns>实际 XML 描述。</returns>
    private static XDocument Xml(Dictionary<string, byte[]> parts, string name) =>
        XDocument.Load(new MemoryStream(parts[name]));

    /// <summary>
    /// 修改测试元数据。
    /// </summary>
    /// <param name="parts">包部件。</param>
    /// <param name="name">元数据名称。</param>
    /// <param name="edit">修改委托。</param>
    private static void Edit(Dictionary<string, byte[]> parts, string name, Action<XDocument> edit)
    {
        var xml = Xml(parts, name);
        edit(xml);
        Save(parts, name, xml);
    }

    /// <summary>
    /// 保存测试元数据。
    /// </summary>
    /// <param name="parts">包部件。</param>
    /// <param name="name">元数据名称。</param>
    /// <param name="xml">修改后的 XML。</param>
    private static void Save(Dictionary<string, byte[]> parts, string name, XDocument xml) =>
        parts[name] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
}
