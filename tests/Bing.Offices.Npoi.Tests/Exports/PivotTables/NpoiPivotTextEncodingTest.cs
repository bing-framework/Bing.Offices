using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Exports.PivotTables;
using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证透视缓存文本遵循 SpreadsheetML Xstring 编码规则。
/// </summary>
[Collection("NpoiNativePivot")]
public sealed class NpoiPivotTextEncodingTest
{
    /// <summary>
    /// 验证控制字符、字面转义序列和 Unicode 文本经过一次编码后可还原。
    /// </summary>
    [Fact]
    public void Encode_ShouldEscapeXstringControlsAndLiteralSequencesOnce()
    {
        var original = "start\r\n\t_x000A_ _xFFFF_ __ 普通文本 中文\u0001 end";
        var encoded = NpoiPivotTextEncoding.Encode(original);

        Assert.Equal("start_x000D__x000A__x0009__x005F_x000A_ _x005F_xFFFF_ __ 普通文本 中文_x0001_ end",
            encoded);
        Assert.Equal(original, DecodeXstringOnce(encoded));
        Assert.Equal("_x005F_x000A_", NpoiPivotTextEncoding.Encode("_x000A_"));
        Assert.Equal("__x005F_x000A_", NpoiPivotTextEncoding.Encode("__x000A_"));
        Assert.Equal("bad_xFFFF_", NpoiPivotTextEncoding.Encode("bad\uFFFF"));
        Assert.Equal("emoji \U0001F642", NpoiPivotTextEncoding.Encode("emoji \U0001F642"));
    }

    /// <summary>
    /// 验证实际缓存定义对字段名和文本值应用 Xstring 编码并保留来源关系。
    /// </summary>
    [Fact]
    public void CacheDefinition_ShouldSerializeEncodedHeadersAndSharedText()
    {
        var header = "Item_x000A_\n列";
        var value = "quote \"& 中文\rline\n_x000A_\t" + new string('界', 300);
        using var workbook = new XSSFWorkbook();
        var data = workbook.CreateSheet("Data");
        var summary = (XSSFSheet)workbook.CreateSheet("Summary");
        data.CreateRow(0).CreateCell(0).SetCellValue(header);
        data.CreateRow(1).CreateCell(0).SetCellValue(value);
        var source = new AreaReference(new CellReference("A1"), new CellReference(1, 0),
            SpreadsheetVersion.EXCEL2007);
        var pivot = summary.CreatePivotTable(source, new CellReference("A1"), data);
        pivot.AddRowLabel(0);
        IReadOnlyList<IReadOnlyList<object>> rows = new IReadOnlyList<object>[] { new object[] { value } };
        var records = NpoiPivotCacheWriter.Complete(pivot, new[] { header }, rows, default);
        using var output = new MemoryStream();

        NpoiSheetContentWriter.Write(workbook,
            new Dictionary<string, IReadOnlyList<ExcelDataValidationDefinition>>(),
            new[] { records }, output, default);

        using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        var mainNs = (XNamespace)"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var relationshipNs = (XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var packageRelationshipNs = (XNamespace)"http://schemas.openxmlformats.org/package/2006/relationships";
        var workbookRelationships = ReadXml(archive, "xl/_rels/workbook.xml.rels");
        var definitionRelationship = Assert.Single(workbookRelationships.Root
            .Elements(packageRelationshipNs + "Relationship")
            .Where(element => ((string)element.Attribute("Type"))?.EndsWith("/pivotCacheDefinition",
                StringComparison.Ordinal) == true));
        var definitionPath = ResolvePart("xl/workbook.xml", (string)definitionRelationship.Attribute("Target"));
        var definition = ReadXml(archive, definitionPath);
        var definitionRelationshipId = (string)definition.Root.Attribute(relationshipNs + "id");
        Assert.False(string.IsNullOrWhiteSpace(definitionRelationshipId));
        var sourceNode = definition.Root.Element(mainNs + "cacheSource")
            .Element(mainNs + "worksheetSource");
        Assert.Equal("Data", (string)sourceNode.Attribute("sheet"));
        Assert.Equal("A1:A2", (string)sourceNode.Attribute("ref"));
        var field = Assert.Single(definition.Root.Element(mainNs + "cacheFields").Elements(mainNs + "cacheField"));
        Assert.Equal(Encode(header), (string)field.Attribute("name"));
        Assert.Equal(header, DecodeXstringOnce((string)field.Attribute("name")));
        var sharedItems = field.Element(mainNs + "sharedItems");
        var sharedText = Assert.Single(sharedItems.Elements(mainNs + "s"));
        Assert.Equal(Encode(value), (string)sharedText.Attribute("v"));
        Assert.Equal(value, DecodeXstringOnce((string)sharedText.Attribute("v")));
        Assert.True(ReadSharedFlag(sharedItems, "containsString"));
        Assert.True(ReadSharedFlag(sharedItems, "longText"));
        Assert.Equal(1, (int)sharedItems.Attribute("count"));
        Assert.NotNull(archive.GetEntry(records.PartName));
        var definitionRelationships = ReadXml(archive, RelationshipPath(definitionPath));
        var recordsRelationship = Assert.Single(definitionRelationships.Root
            .Elements(packageRelationshipNs + "Relationship")
            .Where(element => (string)element.Attribute("Id") == definitionRelationshipId));
        Assert.EndsWith("/pivotCacheRecords", (string)recordsRelationship.Attribute("Type"));
        Assert.Equal(records.PartName, ResolvePart(definitionPath, (string)recordsRelationship.Attribute("Target")));
    }

    /// <summary>
    /// 编码测试输入并确保测试值使用一次 Xstring 转义。
    /// </summary>
    /// <param name="value">原始文本。</param>
    /// <returns>编码后的文本属性值。</returns>
    private static string Encode(string value) => NpoiPivotTextEncoding.Encode(value);

    /// <summary>
    /// 解码一遍 SpreadsheetML Xstring 转义序列。
    /// </summary>
    /// <param name="value">编码后的文本。</param>
    /// <returns>一次解码后的原文。</returns>
    private static string DecodeXstringOnce(string value)
    {
        var result = new StringBuilder(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            if (index + 6 < value.Length && value[index] == '_' && value[index + 1] == 'x'
                && value[index + 6] == '_' && Uri.IsHexDigit(value[index + 2])
                && Uri.IsHexDigit(value[index + 3]) && Uri.IsHexDigit(value[index + 4])
                && Uri.IsHexDigit(value[index + 5]))
            {
                result.Append((char)int.Parse(value.Substring(index + 2, 4), NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture));
                index += 6;
            }
            else
                result.Append(value[index]);
        }
        return result.ToString();
    }

    /// <summary>
    /// 读取真实缓存定义中的共享项标志及其架构默认值。
    /// </summary>
    /// <param name="element">包含共享项标志的 XML 元素。</param>
    /// <param name="name">标志属性名称。</param>
    /// <returns>架构解析后的标志值。</returns>
    private static bool ReadSharedFlag(XElement element, string name)
    {
        var document = new System.Xml.XmlDocument();
        document.LoadXml(element.ToString(SaveOptions.DisableFormatting));
        var namespaces = new System.Xml.XmlNamespaceManager(document.NameTable);
        namespaces.AddNamespace("x", element.Name.NamespaceName);
        var parsed = NPOI.OpenXmlFormats.Spreadsheet.CT_SharedItems.Parse(document.DocumentElement, namespaces);
        return name switch
        {
            "containsString" => parsed.containsString,
            "longText" => parsed.longText,
            _ => throw new ArgumentOutOfRangeException(nameof(name))
        };
    }

    /// <summary>
    /// 读取压缩包中的 XML 部件。
    /// </summary>
    /// <param name="archive">待读取的工作簿包。</param>
    /// <param name="partName">部件路径。</param>
    /// <returns>解析后的 XML。</returns>
    private static XDocument ReadXml(ZipArchive archive, string partName)
    {
        using var stream = archive.GetEntry(partName).Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// 根据来源部件和关系目标解析包内绝对路径。
    /// </summary>
    /// <param name="sourcePart">关系来源部件。</param>
    /// <param name="target">相对关系目标。</param>
    /// <returns>包内标准部件路径。</returns>
    private static string ResolvePart(string sourcePart, string target)
    {
        var segments = sourcePart.Split('/').SkipLast(1).ToList();
        foreach (var segment in target.Split('/'))
        {
            if (segment == "..")
                segments.RemoveAt(segments.Count - 1);
            else if (segment != ".")
                segments.Add(segment);
        }
        return string.Join("/", segments);
    }

    /// <summary>
    /// 获取指定部件的 OPC 关系 XML 路径。
    /// </summary>
    /// <param name="partName">关系来源部件。</param>
    /// <returns>关系部件路径。</returns>
    private static string RelationshipPath(string partName)
    {
        var slash = partName.LastIndexOf('/');
        return partName.Substring(0, slash + 1) + "_rels/" + partName.Substring(slash + 1) + ".rels";
    }
}
