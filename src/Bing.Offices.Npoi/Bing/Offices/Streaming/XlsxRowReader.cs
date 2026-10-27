using System.Globalization;
using System.Text;
using System.Xml;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using NPOI.SS.UserModel;

namespace Bing.Offices.Streaming;

/// <summary>仅保留当前行中的受限列，所有坐标在解析边界转换为零基。</summary>
internal sealed class XlsxRow
{
    internal int Index { get; init; }
    internal int LastColumn { get; set; } = -1;
    internal Dictionary<int, ExcelCellValue> Cells { get; } = new();
    internal Dictionary<int, (string Message, string RawValue)> Errors { get; } = new();
    internal ExcelCellValue Get(int column) => Cells.TryGetValue(column, out var value)
        ? value : new ExcelCellValue(null, "", ExcelCellKind.Empty);
}

/// <summary>顺序读取 Sheet XML，不依赖 dimension，也不分配完整行列表。</summary>
internal sealed class XlsxRowReader : IDisposable
{
    private readonly XlsxPackageSession _package;
    private readonly ExcelStreamingImportOptions _options;
    private readonly XmlReader _reader;
    private readonly int _columns;
    private readonly bool _asynchronous;
    private readonly CancellationToken _token;
    private int _lastRow = -1;
    private int _sheetDataDepth = -1;

    /// <summary>正在解析的物理行；供行内资源错误定位，不能用下一期望行替代。</summary>
    internal int CurrentRowIndex => _lastRow;

    internal XlsxRowReader(XlsxPackageSession package, string path, int columns,
        ExcelStreamingImportOptions options, bool asynchronous, CancellationToken token)
    {
        _package = package; _options = options; _columns = columns; _asynchronous = asynchronous; _token = token;
        _reader = package.OpenXml(path, asynchronous, options.ResourceLimits.MaxWorksheetBytes);
    }
    private ValueTask<bool> Read() => StreamingXml.Read(_reader, _asynchronous, _options.ResourceLimits.MaxXmlDepth.Value, _token);

    internal async ValueTask<XlsxRow> Next()
    {
        while (await Read().ConfigureAwait(false))
        {
            if (_reader.NodeType == XmlNodeType.Element && _reader.Depth == 0
                && (_reader.LocalName != "worksheet" || _reader.NamespaceURI != XlsxPackageSession.MainNamespace))
                throw new NotSupportedException("不支持此工作表 XML 格式。");
            if (_reader.LocalName == "sheetData" && _reader.NamespaceURI == XlsxPackageSession.MainNamespace)
            {
                if (_reader.NodeType == XmlNodeType.Element && !_reader.IsEmptyElement) _sheetDataDepth = _reader.Depth;
                else _sheetDataDepth = -1;
            }
            if (_sheetDataDepth < 0 || _reader.NodeType != XmlNodeType.Element || _reader.LocalName != "row"
                || _reader.NamespaceURI != XlsxPackageSession.MainNamespace || _reader.Depth != _sheetDataDepth + 1) continue;
            var rowNumber = _reader.GetAttribute("r");
            var index = rowNumber == null ? _lastRow + 1 : int.Parse(rowNumber, CultureInfo.InvariantCulture) - 1;
            if (index <= _lastRow || index < 0 || index >= 1048576) throw new InvalidDataException("工作表行坐标无效或不递增。");
            _lastRow = index;
            var row = new XlsxRow { Index = index };
            if (_reader.IsEmptyElement) return row;
            var depth = _reader.Depth;
            long rowCharacters = 0;
            while (await Read().ConfigureAwait(false))
            {
                if (_reader.NodeType == XmlNodeType.EndElement && _reader.Depth == depth) return row;
                if (_reader.NodeType != XmlNodeType.Element || _reader.LocalName != "c" || _reader.Depth != depth + 1) continue;
                var reference = _reader.GetAttribute("r");
                var column = reference == null ? row.LastColumn + 1 : Column(reference, index);
                if (column <= row.LastColumn || column < 0 || column >= 16384) throw new InvalidDataException("单元格列坐标无效或重复。");
                row.LastColumn = column;
                var type = _reader.GetAttribute("t") ?? "n";
                var styleText = _reader.GetAttribute("s") ?? "0";
                var value = new StringBuilder();
                var formula = new StringBuilder();
                var hasValue = false;
                var hasFormula = false;
                var valueDepth = -1;
                var formulaDepth = -1;
                var phoneticDepth = -1;
                var cellDepth = _reader.Depth;
                if (!_reader.IsEmptyElement)
                {
                    while (await Read().ConfigureAwait(false))
                    {
                        if (_reader.NodeType == XmlNodeType.EndElement && _reader.Depth == cellDepth) break;
                        if (_reader.NodeType == XmlNodeType.Element)
                        {
                            if (_reader.LocalName == "rPh" && !_reader.IsEmptyElement) phoneticDepth = _reader.Depth;
                            if (_reader.LocalName == "v" || (_reader.LocalName == "t" && phoneticDepth < 0))
                            { hasValue = true; if (!_reader.IsEmptyElement) valueDepth = _reader.Depth; }
                            if (_reader.LocalName == "f")
                            { hasFormula = true; if (!_reader.IsEmptyElement) formulaDepth = _reader.Depth; }
                        }
                        else if (_reader.NodeType == XmlNodeType.EndElement)
                        {
                            if (_reader.Depth == valueDepth) valueDepth = -1;
                            if (_reader.Depth == formulaDepth) formulaDepth = -1;
                            if (_reader.Depth == phoneticDepth) phoneticDepth = -1;
                        }
                        else if (_reader.NodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.SignificantWhitespace or XmlNodeType.Whitespace)
                        {
                            var destination = formulaDepth >= 0 ? formula : valueDepth >= 0 && phoneticDepth < 0 ? value : null;
                            if (destination != null)
                                destination.Append(await StreamingXml.Text(_reader, _asynchronous,
                                    _options.MaxCellCharacters - destination.Length, _token).ConfigureAwait(false));
                        }
                    }
                }
                rowCharacters += value.Length + formula.Length;
                if (column < _columns)
                {
                    var raw = value.ToString();
                    try
                    {
                        var parsed = await Parse(type, styleText, raw, hasValue, hasFormula, formula.ToString()).ConfigureAwait(false);
                        // 共享字符串索引的短文本不能绕过当前行的真实字符串额度。
                        if (type == "s") rowCharacters += parsed.Text.Length;
                        row.Cells.Add(column, parsed);
                    }
                    catch (Exception e) when (e is FormatException or OverflowException or ArgumentException or InvalidDataException)
                    { row.Errors.Add(column, (e.Message, raw)); }
                }
                if (rowCharacters > _options.MaxRowCharacters) throw new StreamingResourceLimitException("单行文本超过限制。");
            }
            throw new InvalidDataException("工作表行未正确结束。");
        }
        return null;
    }

    private async ValueTask<ExcelCellValue> Parse(string type, string styleText, string raw,
        bool hasValue, bool hasFormula, string formula)
    {
        var style = int.Parse(styleText, CultureInfo.InvariantCulture);
        if (style < 0 || style >= _package.StyleFormats.Count) throw new InvalidDataException("单元格样式索引越界。");
        var format = _package.StyleFormats[style];
        if (hasFormula && !hasValue) throw new InvalidDataException("公式缺少缓存结果，流式导入不执行公式计算。");
        object value = null;
        var text = raw;
        var kind = ExcelCellKind.Empty;
        int? errorCode = null;
        if (hasValue)
        {
            switch (type)
            {
                case "s":
                    var index = int.Parse(raw, CultureInfo.InvariantCulture);
                    text = _asynchronous ? await _package.Strings.GetAsync(index, _token).ConfigureAwait(false) : _package.Strings.Get(index);
                    value = text; kind = ExcelCellKind.Text; break;
                case "inlineStr": case "str":
                    value = text = XlsxPackageSession.Decode(raw); kind = ExcelCellKind.Text; break;
                case "b":
                    value = raw switch { "1" or "true" => true, "0" or "false" => false, _ => throw new InvalidDataException("布尔单元格值无效。") };
                    text = (bool)value ? "True" : "False"; kind = ExcelCellKind.Boolean; break;
                case "e":
                    errorCode = FormulaError.ForString(raw).Code;
                    text = errorCode.Value.ToString(CultureInfo.InvariantCulture);
                    kind = ExcelCellKind.Error; break;
                case "d":
                    value = DateTime.Parse(raw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
                    kind = ExcelCellKind.DateTime; break;
                case "n":
                    var number = double.Parse(raw, NumberStyles.Float, CultureInfo.InvariantCulture);
                    if (double.IsNaN(number) || double.IsInfinity(number)) throw new InvalidDataException("数值单元格必须为有限数值。");
                    var code = _package.NumberFormats.TryGetValue(format, out var custom) ? custom : BuiltinFormats.GetBuiltinFormat(format);
                    if (DateUtil.IsADateFormat(format, code))
                    {
                        if (!DateUtil.IsValidExcelDate(number)) throw new InvalidDataException("Excel 日期序号无效。");
                        value = DateUtil.GetJavaDate(number, _package.Date1904); kind = ExcelCellKind.DateTime;
                    }
                    else { value = number; kind = ExcelCellKind.Number; }
                    break;
                default: throw new InvalidDataException("不支持的单元格类型: " + type);
            }
        }
        if (kind == ExcelCellKind.DateTime) text = NpoiScalarText.Date((DateTime)value);
        return new ExcelCellValue(value, text, hasFormula ? ExcelCellKind.Formula : kind,
            hasFormula ? kind : null, hasFormula ? formula : null, errorCode, format, _package.Date1904);
    }

    private static int Column(string reference, int rowIndex)
    {
        var i = 0; var column = 0;
        while (i < reference.Length && reference[i] is >= 'A' and <= 'Z')
        {
            if (i >= 3) throw new InvalidDataException("单元格引用列超出范围。");
            column = column * 26 + reference[i++] - 'A' + 1;
        }
        if (i == 0 || !int.TryParse(reference.Substring(i), NumberStyles.None, CultureInfo.InvariantCulture, out var row)
            || row != rowIndex + 1) throw new InvalidDataException("单元格引用与行坐标不一致。");
        return column - 1;
    }
    public void Dispose() => _reader.Dispose();
}
