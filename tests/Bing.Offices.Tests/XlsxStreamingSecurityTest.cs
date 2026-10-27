using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Streaming;
using Xunit;

namespace Bing.Offices.Tests;

public class XlsxStreamingSecurityTest
{
    [Fact]
    public void Import_ShouldReturnInvalidHeaderForMissingSheetNameOrIndex()
    {
        var importer = new NpoiExcelStreamingImporter();
        using var byNameSource = new MemoryStream(NpoiExcelStreamingImporterTest.Package(
            NpoiExcelStreamingImporterTest.Header("Count")));
        var byName = importer.Import(ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.NumberRow>("missing"),
            byNameSource, _ => throw new Xunit.Sdk.XunitException("缺失工作表不得调用消费者。"));

        using var byIndexSource = new MemoryStream(NpoiExcelStreamingImporterTest.Package(
            NpoiExcelStreamingImporterTest.Header("Count")));
        var byIndex = importer.Import(ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.NumberRow>(1),
            byIndexSource, _ => throw new Xunit.Sdk.XunitException("越界索引不得调用消费者。"));

        Assert.Equal(ImportStopReason.InvalidHeader, byName.StopReason);
        Assert.Equal(ExcelImportErrorCode.InvalidHeader, Assert.Single(byName.Errors).Code);
        Assert.Equal(ImportStopReason.InvalidHeader, byIndex.StopReason);
        Assert.Equal(ExcelImportErrorCode.InvalidHeader, Assert.Single(byIndex.Errors).Code);
    }

    [Fact]
    public void Import_ShouldRejectDuplicateZipPaths()
    {
        var package = RewritePackage(
            NpoiExcelStreamingImporterTest.Package(NpoiExcelStreamingImporterTest.Header("Count")),
            duplicatePath: "book/main.xml");
        using var source = new MemoryStream(package);

        Assert.Throws<BingOfficesResourceLimitException>(() => new NpoiExcelStreamingImporter().Import(
            ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.NumberRow>(0), source,
            _ => ExcelRowReadDecision.Continue));
    }

    [Fact]
    public void Import_ShouldProhibitDtdInWorksheetXml()
    {
        var package = RewritePackage(
            NpoiExcelStreamingImporterTest.Package(NpoiExcelStreamingImporterTest.Header("Count")),
            (_, content) => content.Replace("<worksheet ",
                "<!DOCTYPE worksheet [<!ENTITY xxe SYSTEM 'file:///unavailable'>]><worksheet ",
                StringComparison.Ordinal));
        using var source = new MemoryStream(package);
        var callbacks = 0;

        Assert.Throws<XmlException>(() => new NpoiExcelStreamingImporter().Import(
            ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.NumberRow>(0), source,
            _ => { callbacks++; return ExcelRowReadDecision.Continue; }));
        Assert.Equal(0, callbacks);
    }

    [Fact]
    public void Import_ShouldRejectRelationshipPathEscapingPackageRoot()
    {
        var relationshipRewritten = false;
        var package = RewritePackage(
            NpoiExcelStreamingImporterTest.Package(NpoiExcelStreamingImporterTest.Header("Count")),
            (name, content) =>
            {
                if (name != "book/_rels/main.xml.rels") return content;
                var rewritten = content.Replace("Target=\"../data/rows.xml\"",
                    "Target=\"../../outside.xml\"", StringComparison.Ordinal);
                relationshipRewritten = rewritten != content;
                return rewritten;
            });
        Assert.True(relationshipRewritten, "测试 fixture 必须实际替换工作表关系路径。");
        using var source = new MemoryStream(package);

        Assert.Throws<InvalidDataException>(() => new NpoiExcelStreamingImporter().Import(
            ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.NumberRow>(0), source,
            _ => ExcelRowReadDecision.Continue));
    }

    [Fact]
    public void Import_ShouldEnforceMaxSheetsAndMaxFonts()
    {
        var workbookRewritten = false;
        var relationshipsRewritten = false;
        var twoSheets = RewritePackage(
            NpoiExcelStreamingImporterTest.Package(NpoiExcelStreamingImporterTest.Header("Count")),
            (name, content) =>
            {
                if (name == "book/main.xml")
                {
                    var rewritten = content.Replace("</sheets>",
                        "<sheet name=\"第二\" sheetId=\"2\" r:id=\"sheet2\"/></sheets>", StringComparison.Ordinal);
                    workbookRewritten = rewritten != content;
                    return rewritten;
                }
                if (name == "book/_rels/main.xml.rels")
                {
                    var rewritten = content.Replace("</Relationships>",
                        "<Relationship Id=\"sheet2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"../data/rows.xml\"/></Relationships>",
                        StringComparison.Ordinal);
                    relationshipsRewritten = rewritten != content;
                    return rewritten;
                }
                return content;
            });
        Assert.True(workbookRewritten, "测试 fixture 必须实际增加第二个工作表。");
        Assert.True(relationshipsRewritten, "测试 fixture 必须实际增加第二个工作表关系。");
        using (var source = new MemoryStream(twoSheets))
        {
            Assert.Throws<StreamingResourceLimitException>(() => new NpoiExcelStreamingImporter().Import(
                ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.NumberRow>(0), source,
                _ => ExcelRowReadDecision.Continue, new ExcelStreamingImportOptions { MaxSheets = 1 }));
        }

        var fontsRewritten = false;
        var tooManyFonts = RewritePackage(
            NpoiExcelStreamingImporterTest.Package(NpoiExcelStreamingImporterTest.Header("Count")),
            (name, content) =>
            {
                if (name != "book/style.xml") return content;
                var rewritten = content.Replace("><cellXfs", "><fonts><font/><font/></fonts><cellXfs",
                    StringComparison.Ordinal);
                fontsRewritten = rewritten != content;
                return rewritten;
            });
        Assert.True(fontsRewritten, "测试 fixture 必须实际增加字体定义。");
        using (var source = new MemoryStream(tooManyFonts))
        {
            Assert.Throws<StreamingResourceLimitException>(() => new NpoiExcelStreamingImporter().Import(
                ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.NumberRow>(0), source,
                _ => ExcelRowReadDecision.Continue, new ExcelStreamingImportOptions { MaxFonts = 1 }));
        }
    }

    [Fact]
    public void Import_ShouldEnforceZipAndXmlBudgets()
    {
        var package = NpoiExcelStreamingImporterTest.Package(
            NpoiExcelStreamingImporterTest.Header("Count"));
        using (var source = new MemoryStream(package))
        {
            Assert.Throws<BingOfficesResourceLimitException>(() => new NpoiExcelStreamingImporter().Import(
                ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.NumberRow>(0), source,
                _ => ExcelRowReadDecision.Continue,
                new ExcelStreamingImportOptions { ResourceLimits = new ExcelResourceLimits { MaxZipEntryUncompressedBytes = 1 } }));
        }

        using (var source = new MemoryStream(package))
        {
            Assert.Throws<BingOfficesResourceLimitException>(() => new NpoiExcelStreamingImporter().Import(
                ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.NumberRow>(0), source,
                _ => ExcelRowReadDecision.Continue,
                new ExcelStreamingImportOptions { ResourceLimits = new ExcelResourceLimits { MaxZipTotalUncompressedBytes = 1 } }));
        }

        using (var source = new MemoryStream(package))
        {
            Assert.Throws<XmlException>(() => new NpoiExcelStreamingImporter().Import(
                ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.NumberRow>(0), source,
                _ => ExcelRowReadDecision.Continue,
                new ExcelStreamingImportOptions { ResourceLimits = new ExcelResourceLimits { MaxXmlCharacters = 32 } }));
        }

        using (var source = new MemoryStream(package))
        {
            Assert.Throws<StreamingResourceLimitException>(() => new NpoiExcelStreamingImporter().Import(
                ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.NumberRow>(0), source,
                _ => ExcelRowReadDecision.Continue,
                new ExcelStreamingImportOptions { ResourceLimits = new ExcelResourceLimits { MaxWorksheetBytes = 1 } }));
        }
    }

    [Fact]
    public void Import_ShouldRejectDynamicColumnInferenceWithoutFallback()
    {
        var rows = NpoiExcelStreamingImporterTest.Header("Code", "未声明列") +
            "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>A</t></is></c>" +
            "<c r=\"B2\" t=\"inlineStr\"><is><t>B</t></is></c></row>";
        using var source = new MemoryStream(NpoiExcelStreamingImporterTest.Package(rows));
        var callbacks = 0;

        Assert.Throws<NotSupportedException>(() => new NpoiExcelStreamingImporter().Import(
            ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.DynamicRow>(0), source,
            _ => { callbacks++; return ExcelRowReadDecision.Continue; }));
        Assert.Equal(0, callbacks);
    }

    private static byte[] RewritePackage(byte[] package,
        Func<string, string, string> transform = null, string duplicatePath = null)
    {
        using var input = new MemoryStream(package);
        using var source = new ZipArchive(input, ZipArchiveMode.Read, true);
        using var output = new MemoryStream();
        using (var destination = new ZipArchive(output, ZipArchiveMode.Create, true))
        {
            foreach (var entry in source.Entries)
            {
                using var entryStream = entry.Open();
                using var content = new MemoryStream();
                entryStream.CopyTo(content);
                var bytes = content.ToArray();
                if ((entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                    || entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) && transform != null)
                {
                    var text = Encoding.UTF8.GetString(bytes);
                    bytes = Encoding.UTF8.GetBytes(transform(entry.FullName, text));
                }
                WriteEntry(destination, entry.FullName, bytes);
            }

            if (duplicatePath != null)
                WriteEntry(destination, duplicatePath, Encoding.UTF8.GetBytes("<duplicate/>"));
        }
        return output.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string path, byte[] bytes)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        stream.Write(bytes, 0, bytes.Length);
    }
}
