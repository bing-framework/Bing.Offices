using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using Bing.Offices.Imports;

namespace Bing.Offices.ExcelDataReader;

/// <summary>
/// 读取并执行 XLSX OOXML 中的有限 Workbook Data Validation 规则。
/// </summary>
/// <remarks>
/// ExcelDataReader 本身只暴露前向单元格读取，因此这里建立独立的只读校验快照。
/// 仅支持有界的一维区域和字面量边界；公式、外部引用和二维区域由调用方按不支持策略处理。
/// </remarks>
internal sealed class ExcelXlsxWorkbookValidation
{
    /// <summary>
    /// 不支持的工作簿原生校验规则消息。
    /// </summary>
    private const string UnsupportedMessage = "Workbook Data Validation 规则类型或公式暂不支持。";
    /// <summary>
    /// 按工作表名称保存解析后的工作表快照。
    /// </summary>
    private readonly IReadOnlyDictionary<string, XlsxSheet> _sheets;
    /// <summary>
    /// 保存工作簿级和工作表级名称定义。
    /// </summary>
    private readonly IReadOnlyList<XlsxDefinedName> _definedNames;

    /// <summary>
    /// 初始化一个 <see cref="ExcelXlsxWorkbookValidation" /> 类型的实例。
    /// </summary>
    /// <param name="sheets">按工作表名称索引的校验快照。</param>
    /// <param name="definedNames">工作簿名称定义。</param>
    private ExcelXlsxWorkbookValidation(IReadOnlyDictionary<string, XlsxSheet> sheets,
        IReadOnlyList<XlsxDefinedName> definedNames)
    {
        _sheets = sheets;
        _definedNames = definedNames;
    }

    /// <summary>
    /// 从已打开的 XLSX ZIP 读取可安全重映射的失败工作簿原生校验元数据。
    /// </summary>
    /// <param name="archive">已打开的 XLSX ZIP。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>按源工作表排列的校验元数据。</returns>
    /// <remarks>
    /// 该入口只读取工作表校验节点，不加载单元格值；数据校验仅保留有界区域、字面量列表和数值边界规则。
    /// 条件格式另保留无差异样式引用且可在源坐标不变时直接复制的节点；名称、外部引用和其他复杂规则由失败工作簿写出阶段拒绝。
    /// </remarks>
    internal static ExcelXlsxFailureWorkbookMetadata LoadFailureMetadata(ZipArchive archive,
        CancellationToken cancellationToken)
    {
        if (archive == null)
            throw new ArgumentNullException(nameof(archive));

        var workbookEntry = archive.GetEntry("xl/workbook.xml");
        if (workbookEntry == null)
            return ExcelXlsxFailureWorkbookMetadata.Empty;
        var relationshipEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
        var workbook = LoadXml(workbookEntry, cancellationToken);
        var workbookSheets = Elements(workbook, "sheet")
            .Select(element => Attribute(element, "name"))
            .ToArray();
        var hasDifferentialStyles = HasDifferentialStyles(archive, cancellationToken);
        var definedNames = ParseFailureDefinedNames(workbook, workbookSheets,
            cancellationToken);
        var relationships = relationshipEntry == null
            ? new Dictionary<string, string>(StringComparer.Ordinal)
            : ParseRelationships(LoadXml(relationshipEntry, cancellationToken));
        var sheets = new List<ExcelXlsxFailureSheetMetadata>();
        var hasConditionalFormatting = false;
        foreach (var sheetElement in Elements(workbook, "sheet"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var name = Attribute(sheetElement, "name");
            var relationshipId = Attribute(sheetElement, "id");
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(relationshipId)
                || !relationships.TryGetValue(relationshipId, out var target))
                continue;
            var entry = archive.GetEntry(target);
            if (entry == null)
                continue;
            var worksheet = LoadXml(entry, cancellationToken);
            var conditionalFormattings = ParseFailureConditionalFormattings(name, worksheet,
                hasDifferentialStyles, cancellationToken, out var conditionalFormattingCount);
            hasConditionalFormatting |= conditionalFormattingCount > 0;
            var validations = ParseFailureValidations(name, worksheet, cancellationToken);
            var columnRanges = ParseFailureColumnRanges(worksheet, cancellationToken,
                out var unsupportedColumnMetadataCount);
            if (validations.Count > 0 || conditionalFormattingCount > 0
                || columnRanges.Count > 0 || unsupportedColumnMetadataCount > 0)
                sheets.Add(new ExcelXlsxFailureSheetMetadata(name, validations,
                    conditionalFormattings, conditionalFormattingCount, columnRanges,
                    unsupportedColumnMetadataCount));
        }
        return new ExcelXlsxFailureWorkbookMetadata(sheets, definedNames,
            hasConditionalFormatting);
    }

    /// <summary>
    /// 从可定位的 XLSX 流创建校验快照。
    /// </summary>
    /// <param name="source">待读取的 XLSX 流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>工作簿原生校验快照。</returns>
    internal static ExcelXlsxWorkbookValidation Load(Stream source, CancellationToken cancellationToken)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanSeek)
            throw new InvalidOperationException("XLSX 原生校验解析需要可定位的输入流。");

        var originalPosition = source.Position;
        source.Position = 0;
        try
        {
            using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
            cancellationToken.ThrowIfCancellationRequested();
            var workbookEntry = archive.GetEntry("xl/workbook.xml")
                ?? throw new InvalidDataException("XLSX ZIP 缺少 xl/workbook.xml。");
            var relationshipEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
            var workbook = LoadXml(workbookEntry, cancellationToken);
            var relationships = relationshipEntry == null
                ? new Dictionary<string, string>(StringComparer.Ordinal)
                : ParseRelationships(LoadXml(relationshipEntry, cancellationToken));
            var sharedStrings = LoadSharedStrings(archive.GetEntry("xl/sharedStrings.xml"), cancellationToken);
            var definedNames = ParseDefinedNames(workbook);
            var sheets = new Dictionary<string, XlsxSheet>(StringComparer.OrdinalIgnoreCase);
            var index = 0;
            foreach (var sheetElement in Elements(workbook, "sheet"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var name = Attribute(sheetElement, "name");
                var relationshipId = Attribute(sheetElement, "id");
                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(relationshipId)
                    || !relationships.TryGetValue(relationshipId, out var target))
                {
                    index++;
                    continue;
                }
                var worksheetEntry = archive.GetEntry(target);
                if (worksheetEntry == null)
                {
                    index++;
                    continue;
                }
                var sheet = ParseSheet(name, index, LoadXml(worksheetEntry, cancellationToken), sharedStrings,
                    cancellationToken);
                sheets[name] = sheet;
                index++;
            }
            return new ExcelXlsxWorkbookValidation(sheets, definedNames);
        }
        finally
        {
            source.Position = originalPosition;
        }
    }

    /// <summary>
    /// 校验指定工作表单元格的全部原生规则。
    /// </summary>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">零基行号。</param>
    /// <param name="columnIndex">零基列号。</param>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="text">按导入空白策略规范化后的单元格文本。</param>
    /// <param name="culture">导入使用的区域性。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>命中的校验结果；没有规则时返回空集合。</returns>
    internal IReadOnlyList<ExcelXlsxValidationResult> ValidateCell(string sheetName, int rowIndex,
        int columnIndex, object raw, string text, CultureInfo culture,
        CancellationToken cancellationToken)
    {
        if (!_sheets.TryGetValue(sheetName ?? string.Empty, out var sheet))
            return Array.Empty<ExcelXlsxValidationResult>();
        var results = new List<ExcelXlsxValidationResult>();
        foreach (var validation in sheet.Validations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!validation.Ranges.Any(range => range.Contains(rowIndex, columnIndex)))
                continue;
            results.Add(Evaluate(sheet, validation, raw, text ?? string.Empty, culture,
                cancellationToken));
        }
        return results.Count == 0 ? Array.Empty<ExcelXlsxValidationResult>() : results;
    }

    /// <summary>
    /// 评估单个 XML 校验规则。
    /// </summary>
    private ExcelXlsxValidationResult Evaluate(XlsxSheet sheet, XlsxValidation validation,
        object raw, string text, CultureInfo culture, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(validation.Type)
            || string.Equals(validation.Type, "any", StringComparison.OrdinalIgnoreCase))
            return ExcelXlsxValidationResult.Valid();
        if (text.Length == 0)
        {
            return validation.AllowBlank
                ? ExcelXlsxValidationResult.Valid()
                : ExcelXlsxValidationResult.Invalid("不允许 Workbook 校验目标为空。");
        }

        switch (validation.Type.Trim().ToLowerInvariant())
        {
            case "list":
                return EvaluateList(sheet, validation, text, cancellationToken);
            case "whole":
                return EvaluateNumber(validation, text, culture, integerOnly: true);
            case "decimal":
                return EvaluateNumber(validation, text, culture, integerOnly: false);
            case "textlength":
                return EvaluateNumber(validation, text.Length.ToString(CultureInfo.InvariantCulture),
                    CultureInfo.InvariantCulture, integerOnly: true, textLength: true);
            case "date":
            case "time":
                return EvaluateDate(validation, raw, text, culture);
            default:
                return ExcelXlsxValidationResult.Unsupported(UnsupportedMessage);
        }
    }

    /// <summary>
    /// 评估列表规则。
    /// </summary>
    private ExcelXlsxValidationResult EvaluateList(XlsxSheet sheet, XlsxValidation validation,
        string text, CancellationToken cancellationToken)
    {
        var formula = validation.Formula1?.Trim();
        if (string.IsNullOrWhiteSpace(formula))
            return ExcelXlsxValidationResult.Unsupported(UnsupportedMessage);
        formula = formula.Replace("&#34;", "\"", StringComparison.Ordinal)
            .Replace("&quot;", "\"", StringComparison.Ordinal)
            .Replace("\\\"", "\"", StringComparison.Ordinal);
        if (formula.Length >= 2 && formula[0] == '"' && formula[^1] == '"')
        {
            var values = new HashSet<string>(formula.Substring(1, formula.Length - 2).Split(',')
                .Select(value => value.Trim()).Where(value => value.Length > 0), StringComparer.Ordinal);
            return values.Contains(text)
                ? ExcelXlsxValidationResult.Valid()
                : ExcelXlsxValidationResult.Invalid("不符合 Workbook 显式列表校验。");
        }

        if (!TryResolveListReference(sheet, formula, out var targetSheet, out var range))
            return ExcelXlsxValidationResult.Unsupported(UnsupportedMessage);
        var allowed = new HashSet<string>(StringComparer.Ordinal);
        foreach (var cell in targetSheet.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (range.Contains(cell.Key.Row, cell.Key.Column))
                allowed.Add(cell.Value ?? string.Empty);
        }
        return allowed.Contains(text)
            ? ExcelXlsxValidationResult.Valid()
            : ExcelXlsxValidationResult.Invalid("不符合 Workbook 列表校验。");
    }

    /// <summary>
    /// 评估整数、小数或文本长度规则。
    /// </summary>
    private static ExcelXlsxValidationResult EvaluateNumber(XlsxValidation validation, string text,
        CultureInfo culture, bool integerOnly, bool textLength = false)
    {
        var second = 0m;
        if (!TryParseDecimal(text, culture, out var value)
            || !TryParseDecimal(validation.Formula1, CultureInfo.InvariantCulture, out var first)
            || (RequiresSecond(validation.Operator)
                && !TryParseDecimal(validation.Formula2, CultureInfo.InvariantCulture, out second)))
            return ExcelXlsxValidationResult.Invalid(textLength
                ? "不符合 Workbook 文本长度校验。" : "不符合 Workbook 数值校验。");
        if (integerOnly && value != decimal.Truncate(value))
            return ExcelXlsxValidationResult.Invalid(textLength
                ? "不符合 Workbook 文本长度校验。" : "不符合 Workbook 数值校验。");
        var valid = Compare(value, first, second, validation.Operator);
        return valid
            ? ExcelXlsxValidationResult.Valid()
            : ExcelXlsxValidationResult.Invalid(textLength
                ? "不符合 Workbook 文本长度校验。" : "不符合 Workbook 数值校验。");
    }

    /// <summary>
    /// 评估日期或时间规则。
    /// </summary>
    private static ExcelXlsxValidationResult EvaluateDate(XlsxValidation validation, object raw,
        string text, CultureInfo culture)
    {
        double value;
        if (raw is DateTime dateTime)
            value = dateTime.ToOADate();
        else if (raw is DateTimeOffset dateTimeOffset)
            value = dateTimeOffset.DateTime.ToOADate();
        else if (raw is TimeSpan timeSpan)
            value = timeSpan.TotalDays;
        else if (TryParseDecimal(text, culture, out var parsedNumber))
            value = (double)parsedNumber;
        else
            value = double.NaN;
        var second = 0m;
        if (double.IsNaN(value)
            || !TryParseDateBoundary(validation.Formula1, culture, out var first)
            || (RequiresSecond(validation.Operator)
                && !TryParseDateBoundary(validation.Formula2, culture, out second)))
            return ExcelXlsxValidationResult.Invalid("不符合 Workbook 日期/时间校验。");
        return Compare((decimal)value, first, second, validation.Operator)
            ? ExcelXlsxValidationResult.Valid()
            : ExcelXlsxValidationResult.Invalid("不符合 Workbook 日期/时间校验。");
    }

    /// <summary>
    /// 解析列表公式对应的一维单元格区域。
    /// </summary>
    private bool TryResolveListReference(XlsxSheet currentSheet, string expression,
        out XlsxSheet targetSheet, out XlsxRange range)
    {
        targetSheet = null;
        range = default;
        expression = expression?.Trim();
        if (expression?.StartsWith("=", StringComparison.Ordinal) == true)
            expression = expression.Substring(1).Trim();
        if (TryParseReference(expression, currentSheet.Name, out var sheetName, out range))
            return _sheets.TryGetValue(sheetName, out targetSheet) && range.IsOneDimensional;

        var localNames = _definedNames.Where(name => string.Equals(name.Name, expression,
            StringComparison.OrdinalIgnoreCase) && name.LocalSheetIndex == currentSheet.Index).ToArray();
        var candidates = localNames.Length > 0 ? localNames
            : _definedNames.Where(name => string.Equals(name.Name, expression,
                StringComparison.OrdinalIgnoreCase) && !name.LocalSheetIndex.HasValue).ToArray();
        if (candidates.Length != 1)
            return false;
        var definition = candidates[0].RefersTo?.Trim();
        if (definition?.StartsWith("=", StringComparison.Ordinal) == true)
            definition = definition.Substring(1).Trim();
        if (!TryParseReference(definition, currentSheet.Name, out sheetName, out range)
            || !range.IsOneDimensional || !_sheets.TryGetValue(sheetName, out targetSheet))
            return false;
        return true;
    }

    /// <summary>
    /// 解析带工作表限定的 A1 区域。
    /// </summary>
    private static bool TryParseReference(string expression, string defaultSheet,
        out string sheetName, out XlsxRange range)
    {
        sheetName = defaultSheet;
        range = default;
        if (string.IsNullOrWhiteSpace(expression)
            || expression.IndexOfAny(new[] { '[', ']', ',', ';' }) >= 0)
            return false;
        var separator = expression.LastIndexOf('!');
        if (separator >= 0)
        {
            var token = expression.Substring(0, separator).Trim();
            if (token.Length >= 2 && token[0] == '\'' && token[^1] == '\'')
                token = token.Substring(1, token.Length - 2).Replace("''", "'",
                    StringComparison.Ordinal);
            if (token.Length == 0 || token.IndexOf('!') >= 0)
                return false;
            sheetName = token;
            expression = expression.Substring(separator + 1).Trim();
        }
        var parts = expression.Split(':');
        if (parts.Length is < 1 or > 2 || !TryParseCell(parts[0], out var firstRow, out var firstColumn))
            return false;
        var lastRow = firstRow;
        var lastColumn = firstColumn;
        if (parts.Length == 2 && !TryParseCell(parts[1], out lastRow, out lastColumn))
            return false;
        if (firstRow > lastRow || firstColumn > lastColumn)
            return false;
        range = new XlsxRange(firstRow, lastRow, firstColumn, lastColumn);
        return true;
    }

    /// <summary>
    /// 解析单个 A1 单元格地址。
    /// </summary>
    private static bool TryParseCell(string value, out int row, out int column)
    {
        row = 0;
        column = 0;
        value = value?.Trim().Replace("$", string.Empty, StringComparison.Ordinal);
        if (string.IsNullOrWhiteSpace(value))
            return false;
        var index = 0;
        while (index < value.Length && value[index] is >= 'A' and <= 'Z'
            or >= 'a' and <= 'z')
        {
            column = checked(column * 26 + char.ToUpperInvariant(value[index]) - 'A' + 1);
            index++;
        }
        if (column == 0 || index == value.Length || !int.TryParse(value.Substring(index),
            NumberStyles.None, CultureInfo.InvariantCulture, out var oneBasedRow) || oneBasedRow <= 0)
            return false;
        row = oneBasedRow - 1;
        column--;
        return true;
    }

    /// <summary>
    /// 解析工作表 XML 中的单元格及校验区域。
    /// </summary>
    private static XlsxSheet ParseSheet(string name, int index, XDocument document,
        IReadOnlyList<string> sharedStrings, CancellationToken cancellationToken)
    {
        var cells = new Dictionary<XlsxCellKey, string>();
        foreach (var cell in Elements(document, "c"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryParseCell(Attribute(cell, "r"), out var row, out var column))
                continue;
            var value = ElementValue(cell, "v");
            var type = Attribute(cell, "t");
            if (string.Equals(type, "s", StringComparison.Ordinal)
                && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var sharedIndex)
                && sharedIndex >= 0 && sharedIndex < sharedStrings.Count)
                value = sharedStrings[sharedIndex];
            else if (string.Equals(type, "inlineStr", StringComparison.Ordinal))
                value = string.Concat(Elements(cell, "t").Select(element => element.Value));
            else if (string.Equals(type, "b", StringComparison.Ordinal))
                value = value == "1" ? "TRUE" : "FALSE";
            cells[new XlsxCellKey(row, column)] = value ?? string.Empty;
        }

        var validations = new List<XlsxValidation>();
        foreach (var element in Elements(document, "dataValidation"))
        {
            var ranges = new List<XlsxRange>();
            foreach (var token in (Attribute(element, "sqref") ?? string.Empty)
                .Split((char[])null, StringSplitOptions.RemoveEmptyEntries))
                if (TryParseReference(token, name, out _, out var range))
                    ranges.Add(range);
            if (ranges.Count == 0)
                continue;
            validations.Add(new XlsxValidation(Attribute(element, "type"),
                Attribute(element, "operator"), ParseBoolean(Attribute(element, "allowBlank")),
                ElementValue(element, "formula1"), ElementValue(element, "formula2"), ranges));
        }
        return new XlsxSheet(name, index, cells, validations);
    }

    /// <summary>
    /// 读取可安全复制的工作表原生校验规则。
    /// </summary>
    private static IReadOnlyList<ExcelXlsxFailureValidation> ParseFailureValidations(string sheetName,
        XDocument document, CancellationToken cancellationToken)
    {
        var validations = new List<ExcelXlsxFailureValidation>();
        foreach (var element in Elements(document, "dataValidation"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var type = Attribute(element, "type")?.Trim();
            var formula1 = NormalizeFailureFormula(ElementValue(element, "formula1"));
            var formula2 = NormalizeFailureFormula(ElementValue(element, "formula2"));
            if (!IsFailureValidationSafe(type, formula1, formula2))
                continue;

            var ranges = new List<ExcelXlsxFailureRange>();
            foreach (var token in (Attribute(element, "sqref") ?? string.Empty)
                .Split((char[])null, StringSplitOptions.RemoveEmptyEntries))
            {
                if (!TryParseReference(token, sheetName, out var referencedSheet, out var range)
                    || !string.Equals(referencedSheet, sheetName, StringComparison.OrdinalIgnoreCase)
                    || range.FirstRow < 0 || range.LastRow >= 1_048_576
                    || range.FirstColumn < 0 || range.LastColumn >= 16_384)
                    continue;
                ranges.Add(new ExcelXlsxFailureRange(range.FirstRow, range.LastRow,
                    range.FirstColumn, range.LastColumn));
            }
            if (ranges.Count == 0)
                continue;
            validations.Add(new ExcelXlsxFailureValidation(type, Attribute(element, "operator"),
                ParseBoolean(Attribute(element, "allowBlank")), formula1, formula2,
                Attribute(element, "errorStyle"), Attribute(element, "showDropDown"),
                ParseBoolean(Attribute(element, "showInputMessage")),
                ParseBoolean(Attribute(element, "showErrorMessage")), Attribute(element, "errorTitle"),
                Attribute(element, "error"), Attribute(element, "promptTitle"),
                Attribute(element, "prompt"), ranges));
        }
        return validations;
    }

    /// <summary>
    /// 读取可以在源坐标不变时保留的条件格式节点。
    /// </summary>
    /// <param name="sheetName">源工作表名称。</param>
    /// <param name="document">源工作表 XML。</param>
    /// <param name="hasDifferentialStyles">样式表是否包含差异样式。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="conditionalFormattingCount">源工作表的条件格式节点总数。</param>
    /// <returns>能够安全保存原始 XML 的条件格式快照。</returns>
    private static IReadOnlyList<ExcelXlsxFailureConditionalFormatting>
        ParseFailureConditionalFormattings(string sheetName, XDocument document,
            bool hasDifferentialStyles, CancellationToken cancellationToken,
            out int conditionalFormattingCount)
    {
        var output = new List<ExcelXlsxFailureConditionalFormatting>();
        conditionalFormattingCount = 0;
        foreach (var element in Elements(document, "conditionalFormatting"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            conditionalFormattingCount++;
            var ranges = new List<ExcelXlsxFailureRange>();
            var valid = true;
            foreach (var token in (Attribute(element, "sqref") ?? string.Empty)
                .Split((char[])null, StringSplitOptions.RemoveEmptyEntries))
            {
                if (!TryParseReference(token, sheetName, out var referencedSheet, out var range)
                    || !string.Equals(referencedSheet, sheetName, StringComparison.OrdinalIgnoreCase)
                    || range.FirstRow < 0 || range.LastRow >= 1_048_576
                    || range.FirstColumn < 0 || range.LastColumn >= 16_384)
                {
                    valid = false;
                    break;
                }
                ranges.Add(new ExcelXlsxFailureRange(range.FirstRow, range.LastRow,
                    range.FirstColumn, range.LastColumn));
            }
            var copy = new XElement(element);
            if (valid)
            {
                foreach (var rule in Elements(copy, "cfRule"))
                {
                    var differentialStyleId = Attribute(rule, "dxfId");
                    if (string.IsNullOrWhiteSpace(differentialStyleId))
                        continue;
                    if (hasDifferentialStyles
                        || !string.Equals(differentialStyleId.Trim(), "0", StringComparison.Ordinal))
                    {
                        valid = false;
                        break;
                    }
                    rule.Attributes().First(attribute => string.Equals(attribute.Name.LocalName,
                        "dxfId", StringComparison.Ordinal)).Remove();
                }
            }
            if (valid && (Elements(copy, "extLst").Any()
                || Elements(copy, "cfRule").Any(HasUnsupportedConditionalFormattingReference)))
                valid = false;
            if (!valid || ranges.Count == 0)
                continue;
            output.Add(new ExcelXlsxFailureConditionalFormatting(
                copy.ToString(SaveOptions.DisableFormatting), ranges));
        }
        return output;
    }

    /// <summary>
    /// 读取失败工作簿可保留的显式列宽和隐藏状态。
    /// </summary>
    /// <param name="document">源工作表 XML。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="unsupportedCount">包含不支持属性或无效范围的列定义数量。</param>
    /// <returns>按列范围排列的显式列元数据。</returns>
    private static IReadOnlyList<(int FirstColumn, int LastColumn, double Width, bool HasWidth,
        bool Hidden)> ParseFailureColumnRanges(XDocument document,
        CancellationToken cancellationToken, out int unsupportedCount)
    {
        var ranges = new List<(int FirstColumn, int LastColumn, double Width, bool HasWidth,
            bool Hidden)>();
        unsupportedCount = 0;
        foreach (var element in Elements(document, "cols").SelectMany(columns =>
            columns.Elements().Where(column => string.Equals(column.Name.LocalName, "col",
                StringComparison.Ordinal))))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var unsupportedAttributes = element.Attributes().Any(attribute =>
                !attribute.IsNamespaceDeclaration
                && !IsSupportedFailureColumnAttribute(attribute.Name.LocalName));
            var minText = Attribute(element, "min");
            var maxText = Attribute(element, "max");
            var widthText = Attribute(element, "width");
            var hiddenText = Attribute(element, "hidden");
            var customWidthText = Attribute(element, "customWidth");
            if (unsupportedAttributes
                || !int.TryParse(minText?.Trim(), NumberStyles.None, CultureInfo.InvariantCulture,
                    out var firstColumn)
                || !int.TryParse(maxText?.Trim(), NumberStyles.None, CultureInfo.InvariantCulture,
                    out var lastColumn)
                || firstColumn < 1 || lastColumn < firstColumn || lastColumn > 16_384
                || !TryParseFailureColumnBoolean(hiddenText, out var hidden)
                || !TryParseFailureColumnBoolean(customWidthText, out _))
            {
                unsupportedCount++;
                continue;
            }

            var hasWidth = widthText != null;
            var width = 0d;
            if (hasWidth && (!double.TryParse(widthText.Trim(), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out width)
                || double.IsNaN(width) || double.IsInfinity(width) || width < 0 || width > 255))
            {
                unsupportedCount++;
                continue;
            }

            if (hasWidth || hidden)
                ranges.Add((firstColumn, lastColumn, width, hasWidth, hidden));
        }

        ranges.Sort((left, right) => left.FirstColumn.CompareTo(right.FirstColumn));
        for (var index = 1; index < ranges.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (ranges[index].FirstColumn <= ranges[index - 1].LastColumn)
                unsupportedCount++;
        }
        return Array.AsReadOnly(ranges.ToArray());
    }

    /// <summary>
    /// 判断列元数据属性是否属于有限复制范围。
    /// </summary>
    private static bool IsSupportedFailureColumnAttribute(string name)
        => string.Equals(name, "min", StringComparison.Ordinal)
        || string.Equals(name, "max", StringComparison.Ordinal)
        || string.Equals(name, "width", StringComparison.Ordinal)
        || string.Equals(name, "hidden", StringComparison.Ordinal)
        || string.Equals(name, "customWidth", StringComparison.Ordinal);

    /// <summary>
    /// 解析 SpreadsheetML 布尔值。
    /// </summary>
    private static bool TryParseFailureColumnBoolean(string value, out bool result)
    {
        if (string.IsNullOrWhiteSpace(value) || value == "0"
            || string.Equals(value, "false", StringComparison.OrdinalIgnoreCase))
        {
            result = false;
            return true;
        }
        if (value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase))
        {
            result = true;
            return true;
        }
        result = false;
        return false;
    }

    /// <summary>
    /// 判断样式表是否包含差异样式定义。
    /// </summary>
    /// <param name="archive">已打开的 XLSX ZIP。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>存在至少一个差异样式时返回 <see langword="true" />。</returns>
    private static bool HasDifferentialStyles(ZipArchive archive,
        CancellationToken cancellationToken)
    {
        var styles = archive.GetEntry("xl/styles.xml");
        return styles != null && Elements(LoadXml(styles, cancellationToken), "dxf").Any();
    }

    /// <summary>
    /// 判断条件格式公式是否包含需要重写的外部或无效引用。
    /// </summary>
    /// <param name="rule">条件格式规则。</param>
    /// <returns>包含不支持引用时返回 <see langword="true" />。</returns>
    private static bool HasUnsupportedConditionalFormattingReference(XElement rule)
        => Elements(rule, "formula").Any(formula =>
            string.IsNullOrWhiteSpace(formula.Value)
            || formula.Value.IndexOfAny(new[] { '!', '[', ']' }) >= 0
            || formula.Value.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0);

    /// <summary>
    /// 判断原生校验公式是否只包含可复制的字面量边界。
    /// </summary>
    private static bool IsFailureValidationSafe(string type, string formula1, string formula2)
    {
        formula1 = NormalizeFailureFormula(formula1);
        formula2 = NormalizeFailureFormula(formula2);
        if (string.IsNullOrWhiteSpace(type)
            || string.Equals(type, "any", StringComparison.OrdinalIgnoreCase))
            return false;
        type = type.Trim().ToLowerInvariant();
        if (string.Equals(type, "list", StringComparison.Ordinal))
        {
            var value = (formula1 ?? string.Empty).Trim();
            return value.Length >= 2 && value[0] == '"' && value[^1] == '"'
                && value.IndexOfAny(new[] { '[', ']', '!' }) < 0;
        }
        if (type is not ("whole" or "decimal" or "date" or "time" or "textlength"))
            return false;
        return TryParseDecimal(formula1, CultureInfo.InvariantCulture, out _)
            && (string.IsNullOrWhiteSpace(formula2)
                || TryParseDecimal(formula2, CultureInfo.InvariantCulture, out _));
    }

    /// <summary>
    /// 规范化常见 Provider 写出的列表公式转义文本。
    /// </summary>
    /// <param name="formula">原始公式文本。</param>
    /// <returns>可用于安全判断和再次写出的公式文本。</returns>
    private static string NormalizeFailureFormula(string formula)
        => formula?.Replace("&#34;", "\"", StringComparison.Ordinal)
            .Replace("&quot;", "\"", StringComparison.Ordinal)
            .Replace("\\\"", "\"", StringComparison.Ordinal);

    /// <summary>
    /// 读取可安全重映射的工作簿定义名称。
    /// </summary>
    /// <param name="document">工作簿 XML 文档。</param>
    /// <param name="sheetNames">按工作簿顺序排列的源工作表名称。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>引用单个一维绝对区域的定义名称。</returns>
    private static IReadOnlyList<ExcelXlsxFailureDefinedName> ParseFailureDefinedNames(
        XDocument document, IReadOnlyList<string> sheetNames,
        CancellationToken cancellationToken)
    {
        var names = new List<ExcelXlsxFailureDefinedName>();
        foreach (var element in Elements(document, "definedName"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var name = Attribute(element, "name")?.Trim();
            if (string.IsNullOrWhiteSpace(name))
                continue;
            if (ParseBoolean(Attribute(element, "hidden")))
                continue;

            string localSheetName = null;
            var localSheetText = Attribute(element, "localSheetId");
            if (!string.IsNullOrWhiteSpace(localSheetText))
            {
                if (!int.TryParse(localSheetText, NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out var localSheetIndex)
                    || localSheetIndex < 0 || localSheetIndex >= sheetNames.Count
                    || string.IsNullOrWhiteSpace(sheetNames[localSheetIndex]))
                    continue;
                localSheetName = sheetNames[localSheetIndex];
            }

            var expression = element.Value?.Trim();
            if (expression?.StartsWith("=", StringComparison.Ordinal) == true)
                expression = expression.Substring(1).Trim();
            if (!TryParseReference(expression, localSheetName ?? string.Empty,
                    out var targetSheetName, out var range)
                || string.IsNullOrWhiteSpace(targetSheetName)
                || !range.IsOneDimensional || range.FirstRow < 0
                || range.LastRow >= 1_048_576 || range.FirstColumn < 0
                || range.LastColumn >= 16_384)
                continue;

            names.Add(new ExcelXlsxFailureDefinedName(name, targetSheetName,
                localSheetName, new ExcelXlsxFailureRange(range.FirstRow, range.LastRow,
                    range.FirstColumn, range.LastColumn)));
        }
        return names;
    }

    /// <summary>
    /// 解析工作簿命名区域。
    /// </summary>
    private static IReadOnlyList<XlsxDefinedName> ParseDefinedNames(XDocument document)
        => Elements(document, "definedName").Select(element => new XlsxDefinedName(
            Attribute(element, "name"),
            int.TryParse(Attribute(element, "localSheetId"), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var index) ? index : (int?)null,
            element.Value)).Where(name => !string.IsNullOrWhiteSpace(name.Name)).ToArray();

    /// <summary>
    /// 解析工作簿关系映射。
    /// </summary>
    private static Dictionary<string, string> ParseRelationships(XDocument document)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var relationship in Elements(document, "Relationship"))
        {
            var id = Attribute(relationship, "Id");
            var target = Attribute(relationship, "Target")?.Replace('\\', '/').TrimStart('/');
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(target))
                continue;
            if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                target = "xl/" + target;
            result[id] = target;
        }
        return result;
    }

    /// <summary>
    /// 读取共享字符串表。
    /// </summary>
    private static IReadOnlyList<string> LoadSharedStrings(ZipArchiveEntry entry,
        CancellationToken cancellationToken)
    {
        if (entry == null)
            return Array.Empty<string>();
        var document = LoadXml(entry, cancellationToken);
        return Elements(document, "si").Select(item => string.Concat(Elements(item, "t")
            .Select(element => element.Value))).ToArray();
    }

    /// <summary>
    /// 安全读取 XML 文档。
    /// </summary>
    private static XDocument LoadXml(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true
        });
        cancellationToken.ThrowIfCancellationRequested();
        return XDocument.Load(reader, LoadOptions.None);
    }

    /// <summary>
    /// 获取不依赖 XML 命名空间的元素集合。
    /// </summary>
    private static IEnumerable<XElement> Elements(XContainer container, string localName)
        => container.Descendants().Where(element => string.Equals(element.Name.LocalName,
            localName, StringComparison.Ordinal));

    /// <summary>
    /// 获取不依赖 XML 命名空间的属性值。
    /// </summary>
    private static string Attribute(XElement element, string localName)
        => element?.Attributes().FirstOrDefault(attribute => string.Equals(attribute.Name.LocalName,
            localName, StringComparison.Ordinal))?.Value;

    /// <summary>
    /// 获取不依赖 XML 命名空间的子元素文本。
    /// </summary>
    private static string ElementValue(XElement element, string localName)
        => element?.Elements().FirstOrDefault(child => string.Equals(child.Name.LocalName,
            localName, StringComparison.Ordinal))?.Value;

    /// <summary>
    /// 读取 OOXML 布尔属性。
    /// </summary>
    private static bool ParseBoolean(string value)
        => string.Equals(value, "1", StringComparison.Ordinal)
        || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// 解析十进制边界值。
    /// </summary>
    private static bool TryParseDecimal(string value, CultureInfo culture, out decimal number)
        => decimal.TryParse(value?.Trim().TrimStart('='), NumberStyles.Number | NumberStyles.Float,
            culture ?? CultureInfo.InvariantCulture, out number)
        || decimal.TryParse(value?.Trim().TrimStart('='), NumberStyles.Number | NumberStyles.Float,
            CultureInfo.InvariantCulture, out number);

    /// <summary>
    /// 解析日期校验边界的 Excel 序列值或日期文本。
    /// </summary>
    private static bool TryParseDateBoundary(string value, CultureInfo culture, out decimal number)
    {
        if (TryParseDecimal(value, culture, out number))
            return true;
        if (DateTime.TryParse(value?.Trim().TrimStart('='), culture ?? CultureInfo.InvariantCulture,
            DateTimeStyles.AllowWhiteSpaces, out var date)
            || DateTime.TryParse(value?.Trim().TrimStart('='), CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces, out date))
        {
            number = (decimal)date.ToOADate();
            return true;
        }
        number = 0;
        return false;
    }

    /// <summary>
    /// 判断操作符是否需要第二个边界。
    /// </summary>
    private static bool RequiresSecond(string operation)
    {
        operation = operation?.Trim();
        return string.Equals(operation, "between", StringComparison.OrdinalIgnoreCase)
            || string.Equals(operation, "notBetween", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 执行数值比较。
    /// </summary>
    private static bool Compare(decimal value, decimal first, decimal second, string operation)
    {
        operation = string.IsNullOrWhiteSpace(operation) ? "between" : operation.Trim();
        return operation.ToLowerInvariant() switch
        {
            "between" => value >= first && value <= second,
            "notbetween" => value < first || value > second,
            "equal" => value == first,
            "notequal" => value != first,
            "greaterthan" => value > first,
            "lessthan" => value < first,
            "greaterthanorequal" => value >= first,
            "lessthanorequal" => value <= first,
            _ => false
        };
    }

    /// <summary>
    /// 表示一个校验规则执行结果。
    /// </summary>
    internal readonly struct ExcelXlsxValidationResult
    {
        /// <summary>
        /// 初始化一个 <see cref="ExcelXlsxValidationResult" /> 类型的实例。
        /// </summary>
        /// <param name="isValid">是否通过校验。</param>
        /// <param name="isUnsupported">是否因规则超出支持范围而无法校验。</param>
        /// <param name="message">校验结果消息。</param>
        private ExcelXlsxValidationResult(bool isValid, bool isUnsupported, string message)
        {
            IsValid = isValid;
            IsUnsupported = isUnsupported;
            Message = message;
        }

        /// <summary>
        /// 获取规则是否通过校验。
        /// </summary>
        internal bool IsValid { get; }
        /// <summary>
        /// 获取规则是否因超出支持范围而无法校验。
        /// </summary>
        internal bool IsUnsupported { get; }
        /// <summary>
        /// 获取校验结果消息。
        /// </summary>
        internal string Message { get; }

        /// <summary>
        /// 创建通过校验的结果。
        /// </summary>
        /// <returns>通过校验的结果。</returns>
        internal static ExcelXlsxValidationResult Valid() => new(true, false, null);
        /// <summary>
        /// 创建校验失败的结果。
        /// </summary>
        /// <param name="message">失败消息。</param>
        /// <returns>校验失败的结果。</returns>
        internal static ExcelXlsxValidationResult Invalid(string message) => new(false, false, message);
        /// <summary>
        /// 创建不支持规则的结果。
        /// </summary>
        /// <param name="message">不支持原因。</param>
        /// <returns>不支持规则的结果。</returns>
        internal static ExcelXlsxValidationResult Unsupported(string message)
            => new(false, true, message);
    }

    /// <summary>
    /// 保存一个工作表的单元格快照和校验规则。
    /// </summary>
    private sealed class XlsxSheet
    {
        /// <summary>
        /// 初始化一个 <see cref="XlsxSheet" /> 类型的实例。
        /// </summary>
        /// <param name="name">工作表名称。</param>
        /// <param name="index">工作表物理索引。</param>
        /// <param name="cells">单元格文本快照。</param>
        /// <param name="validations">工作表校验规则。</param>
        internal XlsxSheet(string name, int index, IReadOnlyDictionary<XlsxCellKey, string> cells,
            IReadOnlyList<XlsxValidation> validations)
        {
            Name = name;
            Index = index;
            Cells = cells;
            Validations = validations;
        }

        /// <summary>
        /// 获取工作表名称。
        /// </summary>
        internal string Name { get; }
        /// <summary>
        /// 获取工作表物理索引。
        /// </summary>
        internal int Index { get; }
        /// <summary>
        /// 获取单元格文本快照。
        /// </summary>
        internal IReadOnlyDictionary<XlsxCellKey, string> Cells { get; }
        /// <summary>
        /// 获取工作表校验规则。
        /// </summary>
        internal IReadOnlyList<XlsxValidation> Validations { get; }
    }

    /// <summary>
    /// 保存一个工作表原生校验规则的解析结果。
    /// </summary>
    private sealed class XlsxValidation
    {
        /// <summary>
        /// 初始化一个 <see cref="XlsxValidation" /> 类型的实例。
        /// </summary>
        /// <param name="type">校验类型。</param>
        /// <param name="operation">比较运算符。</param>
        /// <param name="allowBlank">是否允许空值。</param>
        /// <param name="formula1">第一个校验公式。</param>
        /// <param name="formula2">第二个校验公式。</param>
        /// <param name="ranges">适用区域。</param>
        internal XlsxValidation(string type, string operation, bool allowBlank, string formula1,
            string formula2, IReadOnlyList<XlsxRange> ranges)
        {
            Type = type;
            Operator = operation;
            AllowBlank = allowBlank;
            Formula1 = formula1;
            Formula2 = formula2;
            Ranges = ranges;
        }

        /// <summary>
        /// 获取校验类型。
        /// </summary>
        internal string Type { get; }
        /// <summary>
        /// 获取比较运算符。
        /// </summary>
        internal string Operator { get; }
        /// <summary>
        /// 获取是否允许空值。
        /// </summary>
        internal bool AllowBlank { get; }
        /// <summary>
        /// 获取第一个校验公式。
        /// </summary>
        internal string Formula1 { get; }
        /// <summary>
        /// 获取第二个校验公式。
        /// </summary>
        internal string Formula2 { get; }
        /// <summary>
        /// 获取适用区域。
        /// </summary>
        internal IReadOnlyList<XlsxRange> Ranges { get; }
    }

    /// <summary>
    /// 保存一个工作簿名称定义。
    /// </summary>
    private sealed class XlsxDefinedName
    {
        /// <summary>
        /// 初始化一个 <see cref="XlsxDefinedName" /> 类型的实例。
        /// </summary>
        /// <param name="name">名称。</param>
        /// <param name="localSheetIndex">局部名称所属的工作表索引。</param>
        /// <param name="refersTo">名称引用公式。</param>
        internal XlsxDefinedName(string name, int? localSheetIndex, string refersTo)
        {
            Name = name;
            LocalSheetIndex = localSheetIndex;
            RefersTo = refersTo;
        }

        /// <summary>
        /// 获取名称。
        /// </summary>
        internal string Name { get; }
        /// <summary>
        /// 获取局部名称所属的工作表索引。
        /// </summary>
        internal int? LocalSheetIndex { get; }
        /// <summary>
        /// 获取名称引用公式。
        /// </summary>
        internal string RefersTo { get; }
    }

    /// <summary>
    /// 表示 Xlsx 单元格的零基坐标。
    /// </summary>
    private readonly struct XlsxCellKey : IEquatable<XlsxCellKey>
    {
        /// <summary>
        /// 初始化一个 <see cref="XlsxCellKey" /> 类型的实例。
        /// </summary>
        /// <param name="row">零基行号。</param>
        /// <param name="column">零基列号。</param>
        internal XlsxCellKey(int row, int column)
        {
            Row = row;
            Column = column;
        }

        /// <summary>
        /// 获取零基行号。
        /// </summary>
        internal int Row { get; }
        /// <summary>
        /// 获取零基列号。
        /// </summary>
        internal int Column { get; }

        /// <summary>
        /// 判断坐标是否相同。
        /// </summary>
        /// <param name="other">待比较的坐标。</param>
        /// <returns>坐标相同时为 <see langword="true" />；否则为 <see langword="false" />。</returns>
        public bool Equals(XlsxCellKey other) => Row == other.Row && Column == other.Column;
        /// <inheritdoc />
        public override bool Equals(object obj) => obj is XlsxCellKey other && Equals(other);
        /// <inheritdoc />
        public override int GetHashCode() => (Row * 397) ^ Column;
    }

    /// <summary>
    /// 表示 Xlsx 校验规则覆盖的零基区域。
    /// </summary>
    private readonly struct XlsxRange
    {
        /// <summary>
        /// 初始化一个 <see cref="XlsxRange" /> 类型的实例。
        /// </summary>
        /// <param name="firstRow">起始零基行号。</param>
        /// <param name="lastRow">结束零基行号。</param>
        /// <param name="firstColumn">起始零基列号。</param>
        /// <param name="lastColumn">结束零基列号。</param>
        internal XlsxRange(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
        }

        /// <summary>
        /// 获取起始零基行号。
        /// </summary>
        internal int FirstRow { get; }
        /// <summary>
        /// 获取结束零基行号。
        /// </summary>
        internal int LastRow { get; }
        /// <summary>
        /// 获取起始零基列号。
        /// </summary>
        internal int FirstColumn { get; }
        /// <summary>
        /// 获取结束零基列号。
        /// </summary>
        internal int LastColumn { get; }
        /// <summary>
        /// 获取区域是否为单行或单列区域。
        /// </summary>
        internal bool IsOneDimensional => FirstRow == LastRow || FirstColumn == LastColumn;
        /// <summary>
        /// 判断坐标是否位于区域内。
        /// </summary>
        /// <param name="row">零基行号。</param>
        /// <param name="column">零基列号。</param>
        /// <returns>坐标位于区域内时为 <see langword="true" />；否则为 <see langword="false" />。</returns>
        internal bool Contains(int row, int column) => row >= FirstRow && row <= LastRow
            && column >= FirstColumn && column <= LastColumn;
    }
}

/// <summary>
/// 保存失败工作簿可重映射的 XLSX 原生校验元数据。
/// </summary>
internal sealed class ExcelXlsxFailureWorkbookMetadata
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelXlsxFailureWorkbookMetadata" /> 类型的实例。
    /// </summary>
    /// <param name="sheets">工作表校验元数据。</param>
    /// <param name="definedNames">可重映射的定义名称。</param>
    /// <param name="hasConditionalFormatting">是否存在需要保留的条件格式规则。</param>
    internal ExcelXlsxFailureWorkbookMetadata(IReadOnlyList<ExcelXlsxFailureSheetMetadata> sheets,
        IReadOnlyList<ExcelXlsxFailureDefinedName> definedNames = null,
        bool hasConditionalFormatting = false)
    {
        Sheets = Array.AsReadOnly((sheets ?? Array.Empty<ExcelXlsxFailureSheetMetadata>()).ToArray());
        DefinedNames = Array.AsReadOnly((definedNames
            ?? Array.Empty<ExcelXlsxFailureDefinedName>()).ToArray());
        HasConditionalFormatting = hasConditionalFormatting;
    }

    /// <summary>
    /// 获取空元数据快照。
    /// </summary>
    internal static ExcelXlsxFailureWorkbookMetadata Empty { get; } =
        new ExcelXlsxFailureWorkbookMetadata(Array.Empty<ExcelXlsxFailureSheetMetadata>());

    /// <summary>
    /// 获取按源工作表排列的校验元数据。
    /// </summary>
    internal IReadOnlyList<ExcelXlsxFailureSheetMetadata> Sheets { get; }

    /// <summary>
    /// 获取可安全重映射的定义名称。
    /// </summary>
    internal IReadOnlyList<ExcelXlsxFailureDefinedName> DefinedNames { get; }

    /// <summary>
    /// 获取工作簿是否包含条件格式规则。
    /// </summary>
    internal bool HasConditionalFormatting { get; }

    /// <summary>
    /// 查找指定源工作表的校验元数据。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <returns>匹配的元数据；不存在时返回 <see langword="null" />。</returns>
    internal ExcelXlsxFailureSheetMetadata FindSheet(string name) => Sheets.FirstOrDefault(sheet =>
        string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase));
}

/// <summary>
/// 保存一个可安全重映射的 XLSX 定义名称。
/// </summary>
internal sealed class ExcelXlsxFailureDefinedName
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelXlsxFailureDefinedName" /> 类型的实例。
    /// </summary>
    /// <param name="name">定义名称。</param>
    /// <param name="targetSheetName">定义名称引用的源工作表名称。</param>
    /// <param name="localSheetName">局部名称所属的源工作表名称；全局名称为 null。</param>
    /// <param name="range">定义名称引用的源区域。</param>
    internal ExcelXlsxFailureDefinedName(string name, string targetSheetName,
        string localSheetName, ExcelXlsxFailureRange range)
    {
        Name = name;
        TargetSheetName = targetSheetName;
        LocalSheetName = localSheetName;
        Range = range;
    }

    /// <summary>
    /// 获取定义名称。
    /// </summary>
    internal string Name { get; }

    /// <summary>
    /// 获取定义名称引用的源工作表名称。
    /// </summary>
    internal string TargetSheetName { get; }

    /// <summary>
    /// 获取局部名称所属的源工作表名称。
    /// </summary>
    internal string LocalSheetName { get; }

    /// <summary>
    /// 获取定义名称引用的源区域。
    /// </summary>
    internal ExcelXlsxFailureRange Range { get; }
}

/// <summary>
/// 保存单个源工作表的失败校验元数据。
/// </summary>
internal sealed class ExcelXlsxFailureSheetMetadata
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelXlsxFailureSheetMetadata" /> 类型的实例。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <param name="validations">可重映射的原生校验规则。</param>
    /// <param name="conditionalFormattings">可在源坐标不变时保留的条件格式。</param>
    /// <param name="conditionalFormattingCount">源工作表条件格式节点总数。</param>
    /// <param name="columnRanges">可保留的显式列宽和隐藏状态。</param>
    /// <param name="unsupportedColumnMetadataCount">不支持或无效的列定义数量。</param>
    internal ExcelXlsxFailureSheetMetadata(string name,
        IReadOnlyList<ExcelXlsxFailureValidation> validations,
        IReadOnlyList<ExcelXlsxFailureConditionalFormatting> conditionalFormattings = null,
        int conditionalFormattingCount = 0,
        IReadOnlyList<(int FirstColumn, int LastColumn, double Width, bool HasWidth, bool Hidden)>
            columnRanges = null,
        int unsupportedColumnMetadataCount = 0)
    {
        Name = name;
        Validations = Array.AsReadOnly((validations ?? Array.Empty<ExcelXlsxFailureValidation>()).ToArray());
        ConditionalFormattings = Array.AsReadOnly((conditionalFormattings
            ?? Array.Empty<ExcelXlsxFailureConditionalFormatting>()).ToArray());
        ConditionalFormattingCount = Math.Max(0, conditionalFormattingCount);
        ColumnRanges = Array.AsReadOnly((columnRanges
            ?? Array.Empty<(int FirstColumn, int LastColumn, double Width, bool HasWidth, bool Hidden)>()
            ).ToArray());
        UnsupportedColumnMetadataCount = Math.Max(0, unsupportedColumnMetadataCount);
    }

    /// <summary>
    /// 获取源工作表名称。
    /// </summary>
    internal string Name { get; }

    /// <summary>
    /// 获取可重映射的原生校验规则。
    /// </summary>
    internal IReadOnlyList<ExcelXlsxFailureValidation> Validations { get; }

    /// <summary>
    /// 获取可以在源坐标不变时保留的条件格式。
    /// </summary>
    internal IReadOnlyList<ExcelXlsxFailureConditionalFormatting> ConditionalFormattings { get; }

    /// <summary>
    /// 获取源工作表条件格式节点总数。
    /// </summary>
    internal int ConditionalFormattingCount { get; }

    /// <summary>
    /// 获取可在失败工作簿中保留的列宽和隐藏状态。
    /// </summary>
    internal IReadOnlyList<(int FirstColumn, int LastColumn, double Width, bool HasWidth, bool Hidden)>
        ColumnRanges { get; }

    /// <summary>
    /// 获取包含不支持属性或无效范围的列定义数量。
    /// </summary>
    internal int UnsupportedColumnMetadataCount { get; }
}

/// <summary>
/// 保存一个可在源行坐标不变时直接复制的条件格式节点。
/// </summary>
internal sealed class ExcelXlsxFailureConditionalFormatting
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelXlsxFailureConditionalFormatting" /> 类型的实例。
    /// </summary>
    /// <param name="xml">条件格式节点 XML。</param>
    /// <param name="ranges">节点覆盖的源区域。</param>
    internal ExcelXlsxFailureConditionalFormatting(string xml,
        IReadOnlyList<ExcelXlsxFailureRange> ranges)
    {
        Xml = string.IsNullOrWhiteSpace(xml)
            ? throw new ArgumentException("条件格式 XML 不能为空。", nameof(xml)) : xml;
        Ranges = Array.AsReadOnly((ranges ?? Array.Empty<ExcelXlsxFailureRange>()).ToArray());
    }

    /// <summary>
    /// 获取条件格式节点 XML。
    /// </summary>
    internal string Xml { get; }

    /// <summary>
    /// 获取条件格式覆盖的源区域。
    /// </summary>
    internal IReadOnlyList<ExcelXlsxFailureRange> Ranges { get; }
}

/// <summary>
/// 保存单个可重映射的 XLSX 原生校验规则。
/// </summary>
internal sealed class ExcelXlsxFailureValidation
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelXlsxFailureValidation" /> 类型的实例。
    /// </summary>
    internal ExcelXlsxFailureValidation(string type, string @operator, bool allowBlank,
        string formula1, string formula2, string errorStyle, string showDropDown,
        bool showInputMessage, bool showErrorMessage, string errorTitle, string error,
        string promptTitle, string prompt, IReadOnlyList<ExcelXlsxFailureRange> ranges)
    {
        Type = type;
        Operator = @operator;
        AllowBlank = allowBlank;
        Formula1 = formula1;
        Formula2 = formula2;
        ErrorStyle = errorStyle;
        ShowDropDown = showDropDown;
        ShowInputMessage = showInputMessage;
        ShowErrorMessage = showErrorMessage;
        ErrorTitle = errorTitle;
        Error = error;
        PromptTitle = promptTitle;
        Prompt = prompt;
        Ranges = Array.AsReadOnly((ranges ?? Array.Empty<ExcelXlsxFailureRange>()).ToArray());
    }

    /// <summary>获取校验类型。</summary>
    internal string Type { get; }
    /// <summary>获取比较运算符。</summary>
    internal string Operator { get; }
    /// <summary>获取是否允许空值。</summary>
    internal bool AllowBlank { get; }
    /// <summary>获取第一个边界公式。</summary>
    internal string Formula1 { get; }
    /// <summary>获取第二个边界公式。</summary>
    internal string Formula2 { get; }
    /// <summary>获取错误样式。</summary>
    internal string ErrorStyle { get; }
    /// <summary>获取下拉箭头显示属性。</summary>
    internal string ShowDropDown { get; }
    /// <summary>获取是否显示输入提示。</summary>
    internal bool ShowInputMessage { get; }
    /// <summary>获取是否显示错误提示。</summary>
    internal bool ShowErrorMessage { get; }
    /// <summary>获取错误标题。</summary>
    internal string ErrorTitle { get; }
    /// <summary>获取错误说明。</summary>
    internal string Error { get; }
    /// <summary>获取输入提示标题。</summary>
    internal string PromptTitle { get; }
    /// <summary>获取输入提示说明。</summary>
    internal string Prompt { get; }
    /// <summary>获取源工作表坐标区域。</summary>
    internal IReadOnlyList<ExcelXlsxFailureRange> Ranges { get; }
}

/// <summary>
/// 表示 XLSX 原生校验覆盖的零基矩形区域。
/// </summary>
internal readonly struct ExcelXlsxFailureRange
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelXlsxFailureRange" /> 类型的实例。
    /// </summary>
    internal ExcelXlsxFailureRange(int firstRow, int lastRow, int firstColumn, int lastColumn)
    {
        FirstRow = firstRow;
        LastRow = lastRow;
        FirstColumn = firstColumn;
        LastColumn = lastColumn;
    }

    /// <summary>获取起始零基行号。</summary>
    internal int FirstRow { get; }
    /// <summary>获取结束零基行号。</summary>
    internal int LastRow { get; }
    /// <summary>获取起始零基列号。</summary>
    internal int FirstColumn { get; }
    /// <summary>获取结束零基列号。</summary>
    internal int LastColumn { get; }
}
