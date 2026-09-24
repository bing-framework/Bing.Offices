using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Xml;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.ClosedXml.Internals;

/// <summary>
/// 在创建 ClosedXML DOM 前计算选定工作表的物理行预算。
/// </summary>
internal static class ClosedXmlRowBudgetPreflight
{
    /// <summary>
    /// Open XML 工作簿关系命名空间。
    /// </summary>
    private const string RelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// 查找第一个超过工作簿总行预算的 Sheet。
    /// </summary>
    /// <param name="source">已完成 ZIP/XML 基础预检的可定位 XLSX 流。</param>
    /// <param name="requests">按执行顺序排列的 Sheet 请求。</param>
    /// <param name="comparison">Sheet 名称比较策略。</param>
    /// <param name="maximumRows">整个 Workbook 允许处理的最大数据行数。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>超限描述；未超限或无法解析未选择的 Sheet 时返回 null。</returns>
    internal static ClosedXmlRowBudgetViolation FindViolation(Stream source,
        IReadOnlyList<ExcelSheetImportRequest> requests, ExcelNameComparison comparison,
        int? maximumRows, CancellationToken cancellationToken = default)
    {
        if (!maximumRows.HasValue || requests == null || requests.Count == 0)
            return null;
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanSeek)
            return null;

        var originalPosition = source.Position;
        source.Position = 0;
        try
        {
            using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
            var sheetParts = ReadSheetParts(archive, cancellationToken);
            long processedRows = 0;
            foreach (var request in requests)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var part = ResolvePart(sheetParts, request, comparison);
                if (part == null)
                    continue;
                var entry = archive.GetEntry(part.Target);
                if (entry == null)
                    continue;
                var lastRow = ReadLastRow(entry, cancellationToken);
                var potentialRows = Math.Max(0, lastRow - request.DataRowStartIndex);
                if (processedRows + potentialRows <= maximumRows.Value)
                {
                    processedRows += potentialRows;
                    continue;
                }

                var remaining = maximumRows.Value - processedRows;
                var rowIndex = remaining > 0
                    ? (long)request.DataRowStartIndex + remaining + 1
                    : (long)request.DataRowStartIndex + 1;
                return new ClosedXmlRowBudgetViolation(part.Name, request.ItemType,
                    (int)Math.Min(int.MaxValue, rowIndex),
                    maximumRows.Value);
            }
            return null;
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (InvalidDataException exception)
        {
            throw new BingOfficesImportException("ClosedXML 行预算预检失败。", exception,
                "ClosedXML", BingOfficesStage.Preflight);
        }
        catch (XmlException exception)
        {
            throw new BingOfficesImportException("ClosedXML 行预算 XML 预检失败。", exception,
                "ClosedXML", BingOfficesStage.Preflight);
        }
        finally
        {
            source.Position = originalPosition;
        }
    }

    /// <summary>
    /// 读取工作簿中 Sheet 与 XML 部件的对应关系。
    /// </summary>
    /// <param name="archive">已打开的 XLSX ZIP 存档。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>按工作簿顺序排列的 Sheet 部件。</returns>
    private static IReadOnlyList<SheetPart> ReadSheetParts(ZipArchive archive,
        CancellationToken cancellationToken)
    {
        var relationships = new Dictionary<string, string>(StringComparer.Ordinal);
        var relationshipEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
        if (relationshipEntry != null)
        {
            using var relationshipReader = CreateReader(relationshipEntry.Open());
            while (relationshipReader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (relationshipReader.NodeType != XmlNodeType.Element
                    || !string.Equals(relationshipReader.LocalName, "Relationship",
                        StringComparison.Ordinal))
                    continue;
                var id = relationshipReader.GetAttribute("Id");
                var target = relationshipReader.GetAttribute("Target");
                if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(target))
                    relationships[id] = NormalizeTarget(target);
            }
        }

        var workbookEntry = archive.GetEntry("xl/workbook.xml");
        if (workbookEntry == null)
            return Array.Empty<SheetPart>();
        var result = new List<SheetPart>();
        using var workbookReader = CreateReader(workbookEntry.Open());
        while (workbookReader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (workbookReader.NodeType != XmlNodeType.Element
                || !string.Equals(workbookReader.LocalName, "sheet", StringComparison.Ordinal))
                continue;
            var name = workbookReader.GetAttribute("name") ?? string.Empty;
            var relationshipId = workbookReader.GetAttribute("id", RelationshipNamespace);
            if (!string.IsNullOrWhiteSpace(relationshipId)
                && relationships.TryGetValue(relationshipId, out var target))
                result.Add(new SheetPart(name, target));
        }
        return result;
    }

    /// <summary>
    /// 按名称或索引解析请求对应的 Sheet 部件。
    /// </summary>
    /// <param name="parts">工作簿中的 Sheet 部件。</param>
    /// <param name="request">Sheet 导入请求。</param>
    /// <param name="comparison">Sheet 名称比较策略。</param>
    /// <returns>匹配的 Sheet 部件；没有匹配项时返回 <see langword="null" />。</returns>
    private static SheetPart ResolvePart(IReadOnlyList<SheetPart> parts,
        ExcelSheetImportRequest request, ExcelNameComparison comparison)
    {
        if (request.Selector.Kind == ExcelSheetSelectorKind.ByIndex)
        {
            var index = request.Selector.Index.GetValueOrDefault();
            return index >= 0 && index < parts.Count ? parts[index] : null;
        }
        var comparisonMode = comparison == ExcelNameComparison.Ordinal
            ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        return parts.FirstOrDefault(part => string.Equals(part.Name, request.Selector.Name,
            comparisonMode));
    }

    /// <summary>
    /// 扫描 Worksheet XML 获取实际最后一行号。
    /// </summary>
    /// <param name="entry">Worksheet ZIP 部件。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>实际最后一行号。</returns>
    private static int ReadLastRow(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        var sequentialRow = 0;
        var lastRow = 0;
        using var reader = CreateReader(entry.Open());
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element)
                continue;
            if (string.Equals(reader.LocalName, "row", StringComparison.Ordinal))
            {
                sequentialRow++;
                lastRow = Math.Max(lastRow, ParsePositiveInt(reader.GetAttribute("r"))
                    ?? sequentialRow);
            }
            else if (string.Equals(reader.LocalName, "c", StringComparison.Ordinal))
            {
                lastRow = Math.Max(lastRow, ParseCellRow(reader.GetAttribute("r")));
            }
        }
        return lastRow;
    }

    /// <summary>
    /// 从单元格引用中解析行号。
    /// </summary>
    /// <param name="reference">A1 格式单元格引用。</param>
    /// <returns>解析出的行号；无法解析时返回零。</returns>
    private static int ParseCellRow(string reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
            return 0;
        var firstDigit = 0;
        while (firstDigit < reference.Length && !char.IsDigit(reference[firstDigit]))
            firstDigit++;
        return firstDigit < reference.Length
            ? ParsePositiveInt(reference.Substring(firstDigit)) ?? 0 : 0;
    }

    /// <summary>
    /// 解析正整数文本。
    /// </summary>
    /// <param name="value">待解析文本。</param>
    /// <returns>正整数；输入无效时返回 <see langword="null" />。</returns>
    private static int? ParsePositiveInt(string value) =>
        int.TryParse(value, out var number) && number > 0 ? number : null;

    /// <summary>
    /// 将关系目标规范化为 XLSX 内部部件路径。
    /// </summary>
    /// <param name="target">关系目标路径。</param>
    /// <returns>规范化后的部件路径。</returns>
    private static string NormalizeTarget(string target)
    {
        var normalized = target.Replace('\\', '/').TrimStart('/');
        while (normalized.StartsWith("../", StringComparison.Ordinal))
            normalized = normalized.Substring(3);
        return normalized.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)
            ? normalized : "xl/" + normalized;
    }

    /// <summary>
    /// 创建禁止 DTD 和外部解析器的 XML 读取器。
    /// </summary>
    /// <param name="stream">XML 输入流。</param>
    /// <returns>安全配置的 XML 读取器。</returns>
    private static XmlReader CreateReader(Stream stream) => XmlReader.Create(stream,
        new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true
        });

    /// <summary>
    /// 表示工作簿中的一个 Sheet 与其 XML 部件路径。
    /// </summary>
    private sealed class SheetPart
    {
        /// <summary>
        /// 初始化一个 <see cref="SheetPart" /> 类型的实例。
        /// </summary>
        /// <param name="name">工作表名称。</param>
        /// <param name="target">Worksheet 部件路径。</param>
        internal SheetPart(string name, string target)
        {
            Name = name;
            Target = target;
        }

        /// <summary>
        /// 获取工作表名称。
        /// </summary>
        internal string Name { get; }

        /// <summary>
        /// 获取 Worksheet 部件路径。
        /// </summary>
        internal string Target { get; }
    }
}

/// <summary>
/// ClosedXML 行预算预检命中的结构化描述。
/// </summary>
internal sealed class ClosedXmlRowBudgetViolation
{
    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlRowBudgetViolation" /> 类型的实例。
    /// </summary>
    /// <param name="sheetName">超限工作表名称。</param>
    /// <param name="itemType">工作表项目类型。</param>
    /// <param name="rowIndex">首次可能超限的行号。</param>
    /// <param name="maximumRows">允许的最大行数。</param>
    internal ClosedXmlRowBudgetViolation(string sheetName, Type itemType, int rowIndex,
        int maximumRows)
    {
        SheetName = sheetName;
        ItemType = itemType;
        RowIndex = rowIndex;
        MaximumRows = maximumRows;
    }

    /// <summary>
    /// 获取超限工作表名称。
    /// </summary>
    internal string SheetName { get; }

    /// <summary>
    /// 获取工作表项目类型。
    /// </summary>
    internal Type ItemType { get; }

    /// <summary>
    /// 获取首次可能超限的行号。
    /// </summary>
    internal int RowIndex { get; }

    /// <summary>
    /// 获取允许的最大行数。
    /// </summary>
    internal int MaximumRows { get; }
}
