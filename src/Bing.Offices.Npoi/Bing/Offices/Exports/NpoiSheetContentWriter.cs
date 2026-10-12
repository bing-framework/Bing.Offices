using Bing.Offices.Npoi.Exports.PivotTables;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using System.Text;

using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.Exports;

/// <summary>
/// 写入原生内容、校验规则及完整透视缓存。
/// </summary>
internal static class NpoiSheetContentWriter
{
    /// <summary>
    /// 将工作簿写入目标流。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="request">当前操作请求。</param>
    /// <param name="destination">接收工作簿内容的可写流。</param>
    /// <param name="token">取消令牌。</param>
    internal static void Write(IWorkbook workbook, ExcelWorkbookExportRequest request, Stream destination, CancellationToken token)
        => Write(workbook, request.Sheets.ToDictionary(sheet => sheet.Name,
            sheet => sheet.DataValidations, StringComparer.OrdinalIgnoreCase), destination, token);

    /// <summary>
    /// 将工作簿写入目标流。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="validations">按工作表名称分组的校验定义。</param>
    /// <param name="destination">接收工作簿内容的流。</param>
    /// <param name="token">取消令牌。</param>
    internal static void Write(IWorkbook workbook,
        IReadOnlyDictionary<string, IReadOnlyList<ExcelDataValidationDefinition>> validations,
        Stream destination, CancellationToken token)
        => Write(workbook, validations, Array.Empty<NpoiPivotCacheRecords>(), destination, token);

    /// <summary>
    /// 将工作簿写入目标流。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="validations">按工作表名称分组的校验定义。</param>
    /// <param name="pivotRecords">待补齐的透视缓存记录快照。</param>
    /// <param name="destination">接收工作簿内容的流，所有权归调用方。</param>
    /// <param name="token">取消令牌。</param>
    /// <remarks>
    /// XLSX 内联列表和透视记录共用一次磁盘暂存与包重写，未覆盖部件按原内容复制。
    /// </remarks>
    internal static void Write(IWorkbook workbook,
        IReadOnlyDictionary<string, IReadOnlyList<ExcelDataValidationDefinition>> validations,
        IReadOnlyList<NpoiPivotCacheRecords> pivotRecords, Stream destination, CancellationToken token)
    {
        if (pivotRecords == null)
            throw new ArgumentNullException(nameof(pivotRecords));
        token.ThrowIfCancellationRequested();
        var recordsByPart = pivotRecords.ToDictionary(record => record?.PartName
            ?? throw new ArgumentException("透视缓存记录快照不能为空。", nameof(pivotRecords)),
            StringComparer.Ordinal);
        if (recordsByPart.Count > 0 && workbook is not NPOI.XSSF.UserModel.XSSFWorkbook)
            throw new ArgumentException("透视缓存记录仅支持 XLSX 工作簿。", nameof(pivotRecords));
        var rewriteValidations = workbook is NPOI.XSSF.UserModel.XSSFWorkbook
            && validations.Values.Any(definitions => definitions.Any(rule => rule.Type == ExcelDataValidationType.List
                && rule.ListSourceName == null && rule.CascadeSourceName == null));
        if (!rewriteValidations && recordsByPart.Count == 0)
        {
            workbook.Write(destination, false);
            return;
        }
        var zeroSelections = recordsByPart.Count == 0
            ? new Dictionary<string, IReadOnlyDictionary<int, uint>>()
            : NpoiPivotPageFieldWriter.Capture((NPOI.XSSF.UserModel.XSSFWorkbook)workbook, token);
        // NPOI 2.7.4 的内联列表、完整缓存和单选索引零共用一次包转写。
        // 只加载校验节点和透视定义，缓存记录直接流式写出，不加载业务数据行 XML。
        var path = Path.Combine(Path.GetTempPath(), "bing-npoi-validations-" + Guid.NewGuid().ToString("N") + ".tmp");
        using var staged = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 81920, FileOptions.DeleteOnClose);
        workbook.Write(staged, true);
        token.ThrowIfCancellationRequested();
        staged.Position = 0;
        using var input = new ZipArchive(staged, ZipArchiveMode.Read, true);
        foreach (var part in recordsByPart.Keys.Concat(zeroSelections.Keys))
            if (input.GetEntry(part) == null)
                throw new BingOfficesConfigurationException($"工作簿缺少透视缓存记录部件: {part}",
                    stage: BingOfficesStage.Write);
        var cacheDefinitionParts = FindPivotCacheDefinitionParts(input, recordsByPart.Keys, token);
        var rulesByPart = validations.Where(pair => pair.Value.Any(rule => rule.Type == ExcelDataValidationType.List
                && rule.ListSourceName == null && rule.CascadeSourceName == null))
            .ToDictionary(pair => ((NPOI.XSSF.UserModel.XSSFSheet)workbook.GetSheet(pair.Key)).GetPackagePart().PartName.Name.TrimStart('/'),
                pair => pair.Value.Where(rule => rule.Type == ExcelDataValidationType.List
                    && rule.ListSourceName == null && rule.CascadeSourceName == null).ToArray());
        using var output = new ZipArchive(destination, ZipArchiveMode.Create, true);
        foreach (var entry in input.Entries)
        {
            token.ThrowIfCancellationRequested();
            using var source = entry.Open();
            using var target = output.CreateEntry(entry.FullName).Open();
            if (recordsByPart.TryGetValue(entry.FullName, out var records))
                NpoiPivotCacheRecordWriter.Write(target, records, token);
            else if (cacheDefinitionParts.Contains(entry.FullName))
                RewritePivotCacheDefinitionXml(source, target, token);
            else if (zeroSelections.TryGetValue(entry.FullName, out var selections))
                NpoiPivotPageFieldWriter.Write(source, target, selections, token);
            else if (rulesByPart.TryGetValue(entry.FullName, out var definitions))
                RewriteValidationXml(source, target, definitions, token);
            else
            {
                var buffer = new byte[81920];
                int count;
                while ((count = source.Read(buffer, 0, buffer.Length)) > 0)
                {
                    token.ThrowIfCancellationRequested();
                    target.Write(buffer, 0, count);
                }
            }
        }
        token.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 根据缓存定义与记录部件的关系定位需要补写标志的缓存定义。
    /// </summary>
    /// <param name="archive">已暂存的工作簿包。</param>
    /// <param name="recordParts">本次写入的缓存记录部件。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>与本次记录部件对应的缓存定义部件。</returns>
    private static HashSet<string> FindPivotCacheDefinitionParts(ZipArchive archive,
        IEnumerable<string> recordParts, CancellationToken token)
    {
        var records = new HashSet<string>(recordParts, StringComparer.Ordinal);
        var result = new HashSet<string>(StringComparer.Ordinal);
        const string packageNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        foreach (var entry in archive.Entries)
        {
            token.ThrowIfCancellationRequested();
            var sourcePart = RelationshipSourcePart(entry.FullName);
            if (sourcePart == null || !sourcePart.Contains("/pivotCache/", StringComparison.OrdinalIgnoreCase))
                continue;
            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream,
                new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            var document = XDocument.Load(reader);
            foreach (var relationship in document.Root?.Elements(XName.Get("Relationship", packageNamespace))
                ?? Enumerable.Empty<XElement>())
            {
                token.ThrowIfCancellationRequested();
                var type = (string)relationship.Attribute("Type");
                var target = (string)relationship.Attribute("Target");
                if (target != null && type != null
                    && type.EndsWith("/pivotCacheRecords", StringComparison.Ordinal)
                    && records.Contains(ResolvePackagePart(sourcePart, target)))
                    result.Add(sourcePart);
            }
        }
        return result;
    }

    /// <summary>
    /// 将关系部件路径还原为其对应的源部件路径。
    /// </summary>
    /// <param name="relationshipPart">关系部件路径。</param>
    /// <returns>关系所属的源部件路径；无法解析时返回空值。</returns>
    private static string RelationshipSourcePart(string relationshipPart)
    {
        const string marker = "/_rels/";
        if (relationshipPart == null || !relationshipPart.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            return null;
        var index = relationshipPart.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (index < 0)
            return null;
        var directory = relationshipPart.Substring(0, index);
        var file = relationshipPart.Substring(index + marker.Length,
            relationshipPart.Length - index - marker.Length - ".rels".Length);
        return directory + "/" + file;
    }

    /// <summary>
    /// 解析 OOXML 关系中的相对部件目标。
    /// </summary>
    /// <param name="sourcePart">关系所属的源部件。</param>
    /// <param name="target">关系中的目标路径。</param>
    /// <returns>规范化的包内部件路径。</returns>
    private static string ResolvePackagePart(string sourcePart, string target)
    {
        if (target.StartsWith("/", StringComparison.Ordinal))
            return target.TrimStart('/');
        var source = new Uri("https://bing-offices.invalid/" + sourcePart);
        return Uri.UnescapeDataString(new Uri(source, target).AbsolutePath.TrimStart('/'));
    }

    /// <summary>
    /// 补写文本共享项需要的显式类型标志。
    /// </summary>
    /// <param name="source">原缓存定义 XML。</param>
    /// <param name="target">改写后的缓存定义 XML。</param>
    /// <param name="token">取消令牌。</param>
    private static void RewritePivotCacheDefinitionXml(Stream source, Stream target, CancellationToken token)
    {
        using var reader = XmlReader.Create(source,
            new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        foreach (var sharedItems in document.Descendants()
            .Where(element => element.Name.LocalName == "sharedItems"))
        {
            token.ThrowIfCancellationRequested();
            var items = sharedItems.Elements().ToArray();
            var hasString = false;
            var hasNumber = false;
            var hasBoolean = false;
            var hasDate = false;
            var hasError = false;
            var hasBlank = false;
            var hasInteger = true;
            var hasLongText = false;
            foreach (var element in items)
            {
                token.ThrowIfCancellationRequested();
                switch (element.Name.LocalName)
                {
                    case "s":
                        hasString = true;
                        hasLongText |= DecodePivotText((string)element.Attribute("v") ?? string.Empty).Length > 255;
                        break;
                    case "n":
                        hasNumber = true;
                        if (!double.TryParse((string)element.Attribute("v"), NumberStyles.Float,
                                CultureInfo.InvariantCulture, out var number)
                            || !double.IsFinite(number) || Math.Truncate(number) != number)
                            hasInteger = false;
                        break;
                    case "b":
                        hasBoolean = true;
                        break;
                    case "d":
                        hasDate = true;
                        break;
                    case "e":
                        hasError = true;
                        break;
                    case "m":
                        hasBlank = true;
                        break;
                }
            }
            var typeCount = new[] { hasString, hasNumber, hasBoolean, hasDate, hasError }
                .Count(value => value);
            var hasMixedTypes = typeCount > 1;
            hasInteger = hasNumber && !hasMixedTypes && hasInteger;

            SetSharedItemsFlag(sharedItems, "containsString", hasString);
            // 保留现有缓存生成约定：含文本项时标记为半混合类型。
            SetSharedItemsFlag(sharedItems, "containsSemiMixedTypes", hasString);
            SetSharedItemsFlag(sharedItems, "containsNonDate",
                hasString || hasNumber || hasBoolean || hasError);
            SetSharedItemsFlag(sharedItems, "containsDate", hasDate);
            SetSharedItemsFlag(sharedItems, "containsBlank", hasBlank);
            SetSharedItemsFlag(sharedItems, "containsMixedTypes", hasMixedTypes);
            SetSharedItemsFlag(sharedItems, "containsNumber", hasNumber);
            SetSharedItemsFlag(sharedItems, "containsInteger", hasInteger);
            SetSharedItemsFlag(sharedItems, "longText", hasLongText);
        }
        token.ThrowIfCancellationRequested();
        using var writer = XmlWriter.Create(target,
            new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false });
        document.Save(writer);
        writer.Flush();
        token.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 将共享项类型标志写成显式的 0 或 1。
    /// </summary>
    /// <param name="sharedItems">共享项 XML 节点。</param>
    /// <param name="attributeName">标志属性名称。</param>
    /// <param name="value">标志值。</param>
    private static void SetSharedItemsFlag(XElement sharedItems, string attributeName, bool value) =>
        sharedItems.SetAttributeValue(attributeName, value ? "1" : "0");

    /// <summary>
    /// 解码 SpreadsheetML 文本转义以计算原始字符长度。
    /// </summary>
    /// <param name="value">缓存 XML 中的文本值。</param>
    /// <returns>解码后的文本。</returns>
    private static string DecodePivotText(string value)
    {
        var decoded = new StringBuilder(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            if (index + 6 < value.Length && value[index] == '_' && value[index + 1] == 'x'
                && value[index + 6] == '_' && Uri.IsHexDigit(value[index + 2])
                && Uri.IsHexDigit(value[index + 3]) && Uri.IsHexDigit(value[index + 4])
                && Uri.IsHexDigit(value[index + 5]))
            {
                decoded.Append((char)int.Parse(value.Substring(index + 2, 4), NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture));
                index += 6;
            }
            else
                decoded.Append(value[index]);
        }
        return decoded.ToString();
    }

    /// <summary>
    /// 修正工作表 XML 中的显式列表校验值。
    /// </summary>
    /// <param name="source">包含原工作表 XML 的可读流。</param>
    /// <param name="target">写入修正后工作表 XML 的流。</param>
    /// <param name="definitions">需要修正的显式列表校验定义。</param>
    /// <param name="token">取消令牌。</param>
    private static void RewriteValidationXml(Stream source, Stream target, ExcelDataValidationDefinition[] definitions, CancellationToken token)
    {
        using var reader = XmlReader.Create(source, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        using var writer = XmlWriter.Create(target, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false });
        while (reader.Read())
        {
            token.ThrowIfCancellationRequested();
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "dataValidations")
            {
                using var subtree = reader.ReadSubtree();
                var rules = XElement.Load(subtree);
                var rewritten = new HashSet<XElement>();
                foreach (var definition in definitions)
                {
                    var address = ExcelSheetContent.Address(definition.Range);
                    var expected = "\"" + string.Join(",", definition.Values) + "\"";
                    var rule = rules.Elements().First(element => !rewritten.Contains(element)
                        && (string)element.Attribute("type") == "list"
                        && ((string)element.Attribute("sqref") == address || (string)element.Attribute("sqref") == address.Split(':')[0] && definition.Range.StartRow == definition.Range.EndRow && definition.Range.StartColumn == definition.Range.EndColumn)
                        && string.Equals(element.Elements().Single(item => item.Name.LocalName == "formula1")
                            .Value.Replace("&#34;", "\"").Replace("&quot;", "\""), expected,
                            StringComparison.Ordinal));
                    rule.Elements().Single(element => element.Name.LocalName == "formula1").Value = "\"" + string.Join(",", definition.Values) + "\"";
                    rewritten.Add(rule);
                }
                rules.WriteTo(writer);
                continue;
            }
            switch (reader.NodeType)
            {
                case XmlNodeType.Element:
                    writer.WriteStartElement(reader.Prefix, reader.LocalName, reader.NamespaceURI);
                    writer.WriteAttributes(reader, false);
                    if (reader.IsEmptyElement) writer.WriteEndElement();
                    break;
                case XmlNodeType.EndElement: writer.WriteFullEndElement(); break;
                case XmlNodeType.Text: writer.WriteString(reader.Value); break;
                case XmlNodeType.CDATA: writer.WriteCData(reader.Value); break;
                case XmlNodeType.Whitespace:
                case XmlNodeType.SignificantWhitespace: writer.WriteWhitespace(reader.Value); break;
                case XmlNodeType.Comment: writer.WriteComment(reader.Value); break;
                case XmlNodeType.ProcessingInstruction: writer.WriteProcessingInstruction(reader.Name, reader.Value); break;
            }
        }
    }

    /// <summary>
    /// 将图片和数据校验定义应用到工作表。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="request">当前操作请求。</param>
    /// <param name="token">取消令牌。</param>
    internal static void Apply(IWorkbook workbook, ISheet sheet, ExcelSheetExportRequest request, CancellationToken token)
    {
        foreach (var definition in request.Images)
        {
            token.ThrowIfCancellationRequested();
            var index = workbook.AddPicture(definition.Content,
                definition.Content[0] == 137 ? PictureType.PNG : PictureType.JPEG);
            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            var legacy = workbook is NPOI.HSSF.UserModel.HSSFWorkbook;
            double ColumnSize(int column) => Math.Max(1, sheet.GetColumnWidthInPixels(column));
            double RowSize(int row) => Math.Max(1, (sheet.GetRow(row)?.HeightInPoints ?? sheet.DefaultRowHeightInPoints) * 96d / 72d);
            var startX = Position(definition.Column, definition.OffsetX, ColumnSize);
            var startY = Position(definition.Row, definition.OffsetY, RowSize);
            var endX = Position(definition.Column, (double)definition.OffsetX + definition.Width, ColumnSize);
            var endY = Position(definition.Row, (double)definition.OffsetY + definition.Height, RowSize);
            anchor.Col1 = startX.Index;
            anchor.Row1 = startY.Index;
            anchor.Col2 = endX.Index;
            anchor.Row2 = endY.Index;
            anchor.Dx1 = (int)Math.Round(legacy ? startX.Pixels / ColumnSize(startX.Index) * 1024 : startX.Pixels * 9525);
            anchor.Dy1 = (int)Math.Round(legacy ? startY.Pixels / RowSize(startY.Index) * 256 : startY.Pixels * 9525);
            anchor.Dx2 = (int)Math.Round(legacy ? endX.Pixels / ColumnSize(endX.Index) * 1024 : endX.Pixels * 9525);
            anchor.Dy2 = (int)Math.Round(legacy ? endY.Pixels / RowSize(endY.Index) * 256 : endY.Pixels * 9525);
            anchor.AnchorType = AnchorType.MoveDontResize;
            sheet.CreateDrawingPatriarch().CreatePicture(anchor, index);
        }
    }

    /// <summary>
    /// 将原生校验定义应用到工作表。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="definitions">校验定义。</param>
    /// <param name="token">取消令牌。</param>
    /// <param name="cascadeMaps">级联字典的工作簿映射名称。</param>
    internal static void ApplyValidations(IWorkbook workbook, ISheet sheet,
        IReadOnlyList<ExcelDataValidationDefinition> definitions, CancellationToken token,
        IReadOnlyDictionary<string, string> cascadeMaps = null)
    {
        var helper = sheet.GetDataValidationHelper();
        foreach (var definition in definitions)
        {
            token.ThrowIfCancellationRequested();
            if (definition.ListSourceName != null
                && !NpoiWorkbookValidationPipeline.IsSupportedNamedListSource(sheet,
                    definition.ListSourceName, token))
                throw new BingOfficesConfigurationException(
                    $"列表校验名称缺失或来源不是受支持的一维有界区域: {definition.ListSourceName}",
                    stage: BingOfficesStage.Plan);
            string cascadeFormula = null;
            if (definition.CascadeSourceName != null)
            {
                if (cascadeMaps == null || !cascadeMaps.TryGetValue(definition.CascadeSourceName, out var map))
                    throw new BingOfficesConfigurationException(
                        $"级联列表名称缺失: {definition.CascadeSourceName}", stage: BingOfficesStage.Plan);
                var maxColumns = workbook is NPOI.HSSF.UserModel.HSSFWorkbook ? 256 : 16384;
                var maxRows = workbook is NPOI.HSSF.UserModel.HSSFWorkbook ? 65536 : 1048576;
                if (definition.ParentColumnIndex >= maxColumns || definition.Range.EndColumn >= maxColumns
                    || definition.Range.EndRow >= maxRows)
                    throw new BingOfficesConfigurationException("级联列表区域超过工作表格式边界。",
                        stage: BingOfficesStage.Plan);
                var column = NPOI.SS.Util.CellReference.ConvertNumToColString(definition.ParentColumnIndex.Value);
                cascadeFormula = "INDIRECT(VLOOKUP($" + column + (definition.Range.StartRow + 1)
                    + "," + map + ",2,FALSE))";
                if (cascadeFormula.Length > 255)
                    throw new BingOfficesConfigurationException("级联列表公式超过长度上限。",
                        stage: BingOfficesStage.Plan);
            }
            var op = definition.Operator switch
            {
                ExcelConditionalComparisonOperator.Equal => OperatorType.EQUAL,
                ExcelConditionalComparisonOperator.NotEqual => OperatorType.NOT_EQUAL,
                ExcelConditionalComparisonOperator.GreaterThan => OperatorType.GREATER_THAN,
                ExcelConditionalComparisonOperator.GreaterThanOrEqual => OperatorType.GREATER_OR_EQUAL,
                ExcelConditionalComparisonOperator.LessThan => OperatorType.LESS_THAN,
                ExcelConditionalComparisonOperator.LessThanOrEqual => OperatorType.LESS_OR_EQUAL,
                ExcelConditionalComparisonOperator.Between => OperatorType.BETWEEN,
                _ => OperatorType.NOT_BETWEEN
            };
            var first = definition.Type == ExcelDataValidationType.List ? null : ExcelSheetContent.Value(definition, false, workbook.IsDate1904());
            var second = definition.Operator == ExcelConditionalComparisonOperator.Between
                || definition.Operator == ExcelConditionalComparisonOperator.NotBetween
                ? ExcelSheetContent.Value(definition, true, workbook.IsDate1904()) : null;
            var constraint = definition.Type switch
            {
                ExcelDataValidationType.Integer => helper.CreateintConstraint(op, first, second),
                ExcelDataValidationType.Decimal => helper.CreateDecimalConstraint(op, first, second),
                ExcelDataValidationType.Date => helper.CreateDateConstraint(op, "=" + first, second == null ? null : "=" + second, null),
                ExcelDataValidationType.TextLength => helper.CreateTextLengthConstraint(op, first, second),
                _ => cascadeFormula != null
                    ? helper.CreateFormulaListConstraint(cascadeFormula)
                    : definition.ListSourceName == null
                        ? helper.CreateExplicitListConstraint(definition.Values.ToArray())
                        : helper.CreateFormulaListConstraint(definition.ListSourceName)
            };
            var range = definition.Range;
            var rule = helper.CreateValidation(constraint, new CellRangeAddressList(range.StartRow, range.EndRow, range.StartColumn, range.EndColumn));
            rule.EmptyCellAllowed = definition.IgnoreBlanks;
            rule.ShowErrorBox = definition.ShowErrorMessage;
            rule.ErrorStyle = 0;
            rule.CreateErrorBox(definition.ErrorTitle ?? string.Empty, definition.ErrorMessage ?? string.Empty);
            rule.CreatePromptBox(definition.InputTitle ?? string.Empty, definition.InputMessage ?? string.Empty);
            rule.ShowPromptBox = !string.IsNullOrEmpty(definition.InputTitle) || !string.IsNullOrEmpty(definition.InputMessage);
            sheet.AddValidationData(rule);
        }
    }

    /// <summary>
    /// 将像素偏移换算为单元格索引及单元格内偏移。
    /// </summary>
    /// <param name="index">起始索引，从 0 开始。</param>
    /// <param name="pixels">相对起始单元格的像素偏移。</param>
    /// <param name="size">获取指定行或列像素尺寸的函数。</param>
    /// <returns>偏移落入的单元格索引及该单元格内的像素偏移。</returns>
    private static (int Index, double Pixels) Position(int index, double pixels, Func<int, double> size)
    {
        while (pixels >= size(index)) pixels -= size(index++);
        return (index, pixels);
    }
}
