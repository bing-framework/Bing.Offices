using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;

namespace Bing.Offices.ExcelDataReader;

/// <summary>
/// 将前向导入捕获的失败信息写入失败 XLSX 工作簿。
/// </summary>
/// <remarks>
/// ErrorRowsOnly 生成最小工作簿；AnnotatedOriginal 保留暂存 XLSX 的原始部件，并追加错误摘要和单元格批注。
/// </remarks>
internal static class ExcelDataReaderFailureWorkbookWriter
{
    /// <summary>
    /// SpreadsheetML 主命名空间。
    /// </summary>
    private const string SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// Office 文档关系命名空间。
    /// </summary>
    private const string OfficeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// Package 关系命名空间。
    /// </summary>
    private const string PackageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>
    /// Content Types 命名空间。
    /// </summary>
    private const string ContentTypeNs = "http://schemas.openxmlformats.org/package/2006/content-types";

    /// <summary>
    /// 失败摘要工作表名称。
    /// </summary>
    private const string SummarySheetName = "__BingErrors";

    /// <summary>
    /// 摘要工作表的固定列标题。
    /// </summary>
    private static readonly string[] SummaryHeaders =
    {
        "Sheet", "Row", "Column", "Code", "Message", "Property", "ColumnKey", "Header", "RawValue",
        "FirstRow"
    };

    /// <summary>
    /// 失败工作簿中重映射后的原生校验规则。
    /// </summary>
    private sealed class OutputValidation
    {
        /// <summary>
        /// 初始化一个 <see cref="OutputValidation" /> 类型的实例。
        /// </summary>
        /// <param name="source">源校验规则。</param>
        /// <param name="sqref">输出工作表中的校验区域。</param>
        internal OutputValidation(ExcelXlsxFailureValidation source, IReadOnlyList<string> sqref)
        {
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Sqref = sqref ?? throw new ArgumentNullException(nameof(sqref));
        }

        /// <summary>
        /// 获取源校验规则及其提示属性。
        /// </summary>
        internal ExcelXlsxFailureValidation Source { get; }

        /// <summary>
        /// 获取输出工作表中的校验区域文本。
        /// </summary>
        internal IReadOnlyList<string> Sqref { get; }
    }

    /// <summary>
    /// 失败工作簿中重映射后的定义名称。
    /// </summary>
    private sealed class OutputDefinedName
    {
        /// <summary>
        /// 初始化一个 <see cref="OutputDefinedName" /> 类型的实例。
        /// </summary>
        /// <param name="name">定义名称。</param>
        /// <param name="refersTo">输出工作表区域引用。</param>
        /// <param name="localSheetId">输出工作表的局部作用域索引。</param>
        internal OutputDefinedName(string name, string refersTo, int? localSheetId)
        {
            Name = name;
            RefersTo = refersTo;
            LocalSheetId = localSheetId;
        }

        /// <summary>
        /// 获取定义名称。
        /// </summary>
        internal string Name { get; }

        /// <summary>
        /// 获取输出区域引用。
        /// </summary>
        internal string RefersTo { get; }

        /// <summary>
        /// 获取局部作用域工作表索引。
        /// </summary>
        internal int? LocalSheetId { get; }
    }

    /// <summary>
    /// 保存失败工作簿输出行及其单元格展示状态。
    /// </summary>
    private sealed class OutputRow
    {
        /// <summary>
        /// 初始化一个 <see cref="OutputRow" /> 类型的实例。
        /// </summary>
        internal OutputRow(IReadOnlyList<object> values,
            IReadOnlyDictionary<int, string> numberFormats,
            IReadOnlyDictionary<int, bool> boldColumns,
            IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> fontStyles)
        {
            Values = values ?? Array.Empty<object>();
            NumberFormats = numberFormats ?? new Dictionary<int, string>();
            BoldColumns = boldColumns ?? new Dictionary<int, bool>();
            FontStyles = fontStyles
                ?? new Dictionary<int, ExcelDataReaderFailureFontStyle>();
        }

        /// <summary>
        /// 获取输出单元格值。
        /// </summary>
        internal IReadOnlyList<object> Values { get; }

        /// <summary>
        /// 获取输出单元格数字格式。
        /// </summary>
        internal IReadOnlyDictionary<int, string> NumberFormats { get; }

        /// <summary>
        /// 获取输出单元格加粗状态。
        /// </summary>
        internal IReadOnlyDictionary<int, bool> BoldColumns { get; }

        /// <summary>
        /// 获取输出单元格字体及展示状态。
        /// </summary>
        internal IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> FontStyles { get; }
    }

    /// <summary>
    /// 为低保留 XLSX 输出建立稳定的单元格样式索引。
    /// </summary>
    private sealed class FailureStyleRegistry
    {
        private sealed class StyleEntry
        {
            /// <summary>
            /// 初始化一个样式条目。
            /// </summary>
            /// <param name="format">单元格数字格式。</param>
            /// <param name="isBold">是否使用加粗字体。</param>
            /// <param name="fontStyle">字体及展示状态；无额外状态时为 null。</param>
            internal StyleEntry(string format, bool isBold,
                ExcelDataReaderFailureFontStyle fontStyle)
            {
                Format = format;
                IsBold = isBold;
                FontStyle = fontStyle;
            }

            /// <summary>
            /// 获取单元格数字格式。
            /// </summary>
            internal string Format { get; }

            /// <summary>
            /// 获取是否使用加粗字体。
            /// </summary>
            internal bool IsBold { get; }

            /// <summary>
            /// 获取字体及展示状态。
            /// </summary>
            internal ExcelDataReaderFailureFontStyle FontStyle { get; }

            /// <summary>
            /// 获取或设置样式使用的数字格式索引。
            /// </summary>
            internal int NumberFormatId { get; set; }

            /// <summary>
            /// 获取或设置样式使用的字体索引。
            /// </summary>
            internal int FontId { get; set; }

            /// <summary>
            /// 获取或设置样式使用的填充索引。
            /// </summary>
            internal int FillId { get; set; }

            /// <summary>
            /// 获取或设置样式使用的边框索引。
            /// </summary>
            internal int BorderId { get; set; }
        }

        private sealed class FontEntry
        {
            /// <summary>
            /// 初始化一个字体条目。
            /// </summary>
            /// <param name="isBold">是否使用加粗字体。</param>
            /// <param name="fontStyle">字体及展示状态；无额外状态时为 null。</param>
            internal FontEntry(bool isBold, ExcelDataReaderFailureFontStyle fontStyle)
            {
                IsBold = isBold;
                FontStyle = fontStyle;
            }

            /// <summary>
            /// 获取是否使用加粗字体。
            /// </summary>
            internal bool IsBold { get; }

            /// <summary>
            /// 获取字体及展示状态。
            /// </summary>
            internal ExcelDataReaderFailureFontStyle FontStyle { get; }
        }

        /// <summary>
        /// 保存去重后的单元格样式条目。
        /// </summary>
        private readonly List<StyleEntry> _styles = new List<StyleEntry>();

        /// <summary>
        /// 保存样式复合键到 SpreadsheetML 样式索引的映射。
        /// </summary>
        private readonly Dictionary<string, int> _styleIds =
            new Dictionary<string, int>(StringComparer.Ordinal);

        /// <summary>
        /// 保存去重后的自定义数字格式。
        /// </summary>
        private readonly List<string> _customFormats = new List<string>();

        /// <summary>
        /// 保存去重后的字体定义。
        /// </summary>
        private readonly List<FontEntry> _fonts = new List<FontEntry>();

        /// <summary>
        /// 保存字体复合键到字体索引的映射。
        /// </summary>
        private readonly Dictionary<string, int> _fontIds =
            new Dictionary<string, int>(StringComparer.Ordinal);

        /// <summary>
        /// 保存去重后的纯色填充定义。
        /// </summary>
        private readonly List<ExcelDataReaderFailureFontStyle> _fills =
            new List<ExcelDataReaderFailureFontStyle>();

        /// <summary>
        /// 保存填充复合键到填充索引的映射。
        /// </summary>
        private readonly Dictionary<string, int> _fillIds =
            new Dictionary<string, int>(StringComparer.Ordinal);

        /// <summary>
        /// 保存去重后的四边框定义。
        /// </summary>
        private readonly List<ExcelDataReaderFailureFontStyle> _borders =
            new List<ExcelDataReaderFailureFontStyle>();

        /// <summary>
        /// 保存边框复合键到边框索引的映射。
        /// </summary>
        private readonly Dictionary<string, int> _borderIds =
            new Dictionary<string, int>(StringComparer.Ordinal);

        /// <summary>
        /// 初始化一个 <see cref="FailureStyleRegistry" /> 类型的实例。
        /// </summary>
        private FailureStyleRegistry()
        {
            _fonts.Add(new FontEntry(false, null));
            _fontIds[string.Empty] = 0;
        }

        /// <summary>
        /// 从失败行快照建立样式注册表。
        /// </summary>
        /// <param name="capture">失败行和表头展示状态快照。</param>
        /// <param name="cancellationToken">取消操作的令牌。</param>
        /// <returns>已注册失败行样式的样式注册表。</returns>
        internal static FailureStyleRegistry Build(
            ExcelDataReaderFailureWorkbookCapture capture,
            CancellationToken cancellationToken)
        {
            var registry = new FailureStyleRegistry();
            foreach (var sheet in capture?.Sheets ?? Array.Empty<ExcelDataReaderFailureSheet>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                registry.Register(sheet.HeaderFormats, sheet.HeaderBoldColumns,
                    sheet.HeaderFontStyles, cancellationToken);
                foreach (var row in sheet.Rows)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    registry.Register(row.NumberFormats, row.BoldColumns, row.FontStyles,
                        cancellationToken);
                }
            }
            return registry;
        }

        /// <summary>
        /// 获取单元格样式索引；无额外展示状态时返回默认样式。
        /// </summary>
        /// <param name="column">零基列号。</param>
        /// <param name="numberFormats">单元格数字格式。</param>
        /// <param name="boldColumns">单元格加粗状态。</param>
        /// <param name="fontStyles">单元格字体及展示状态。</param>
        /// <returns>样式表中的单元格格式索引。</returns>
        internal int GetStyleIndex(int column,
            IReadOnlyDictionary<int, string> numberFormats,
            IReadOnlyDictionary<int, bool> boldColumns,
            IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> fontStyles)
        {
            var format = numberFormats != null && numberFormats.TryGetValue(column, out var value)
                ? NormalizeFormat(value) : null;
            var bold = boldColumns != null && boldColumns.TryGetValue(column, out var isBold)
                && isBold;
            var style = fontStyles != null && fontStyles.TryGetValue(column, out var font)
                ? font : null;
            return Register(format, bold, style);
        }

        /// <summary>
        /// 写入完整的最小样式表。
        /// </summary>
        /// <param name="writer">样式表 XML 写入器。</param>
        internal void Write(XmlWriter writer)
        {
            PrepareDefinitions();
            writer.WriteStartElement("styleSheet", SpreadsheetNs);
            WriteNumberFormats(writer);
            WriteFonts(writer);
            WriteFills(writer);
            WriteBorders(writer);
            writer.WriteStartElement("cellStyleXfs", SpreadsheetNs);
            writer.WriteAttributeString("count", "1");
            Empty(writer, "xf", SpreadsheetNs, ("numFmtId", "0"), ("fontId", "0"),
                ("fillId", "0"), ("borderId", "0"));
            writer.WriteEndElement();
            writer.WriteStartElement("cellXfs", SpreadsheetNs);
            writer.WriteAttributeString("count",
                (2 + _styles.Count).ToString(CultureInfo.InvariantCulture));
            WriteXf(writer, 0, 0, 0, 0, null, false);
            WriteXf(writer, 14, 0, 0, 0, null, true);
            foreach (var style in _styles)
                WriteXf(writer, style.NumberFormatId, style.FontId, style.FillId,
                    style.BorderId, style, false);
            writer.WriteEndElement();
            writer.WriteStartElement("cellStyles", SpreadsheetNs);
            writer.WriteAttributeString("count", "1");
            Empty(writer, "cellStyle", SpreadsheetNs, ("name", "Normal"), ("xfId", "0"),
                ("builtinId", "0"));
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        /// <summary>
        /// 注册一组单元格元数据中使用的样式。
        /// </summary>
        /// <param name="numberFormats">单元格数字格式。</param>
        /// <param name="boldColumns">单元格加粗状态。</param>
        /// <param name="fontStyles">单元格字体及展示状态。</param>
        /// <param name="cancellationToken">取消操作的令牌。</param>
        private void Register(IReadOnlyDictionary<int, string> numberFormats,
            IReadOnlyDictionary<int, bool> boldColumns,
            IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> fontStyles,
            CancellationToken cancellationToken)
        {
            var columns = new HashSet<int>();
            if (numberFormats != null)
                foreach (var pair in numberFormats)
                    if (pair.Key >= 0)
                        columns.Add(pair.Key);
            if (boldColumns != null)
                foreach (var pair in boldColumns)
                    if (pair.Key >= 0 && pair.Value)
                        columns.Add(pair.Key);
            if (fontStyles != null)
                foreach (var pair in fontStyles)
                    if (pair.Key >= 0 && pair.Value != null)
                        columns.Add(pair.Key);
            foreach (var column in columns.OrderBy(value => value))
            {
                cancellationToken.ThrowIfCancellationRequested();
                GetStyleIndex(column, numberFormats, boldColumns, fontStyles);
            }
        }

        /// <summary>
        /// 注册一个样式并返回其 SpreadsheetML 索引。
        /// </summary>
        /// <param name="format">单元格数字格式。</param>
        /// <param name="isBold">是否使用加粗字体。</param>
        /// <param name="fontStyle">字体及展示状态。</param>
        /// <returns>样式索引；无额外状态时返回默认索引。</returns>
        private int Register(string format, bool isBold,
            ExcelDataReaderFailureFontStyle fontStyle)
        {
            if (string.IsNullOrWhiteSpace(format) && !isBold && fontStyle == null)
                return 0;
            var key = CreateStyleKey(format, isBold, fontStyle);
            if (_styleIds.TryGetValue(key, out var existing))
                return existing;
            var entry = new StyleEntry(format, isBold, fontStyle);
            _styles.Add(entry);
            var styleIndex = _styles.Count + 1;
            _styleIds[key] = styleIndex;
            if (format != null && !_customFormats.Contains(format, StringComparer.Ordinal))
                _customFormats.Add(format);
            return styleIndex;
        }

        /// <summary>
        /// 为已注册样式分配字体、填充和边框索引。
        /// </summary>
        private void PrepareDefinitions()
        {
            foreach (var style in _styles)
            {
                style.NumberFormatId = style.Format == null
                    ? 0 : 164 + _customFormats.IndexOf(style.Format);
                style.FontId = GetFontId(style.IsBold, style.FontStyle);
                style.FillId = GetFillId(style.FontStyle);
                style.BorderId = GetBorderId(style.FontStyle);
            }
        }

        /// <summary>
        /// 获取或创建字体定义索引。
        /// </summary>
        /// <param name="isBold">是否使用加粗字体。</param>
        /// <param name="fontStyle">字体及展示状态。</param>
        /// <returns>字体定义索引。</returns>
        private int GetFontId(bool isBold, ExcelDataReaderFailureFontStyle fontStyle)
        {
            var key = CreateFontKey(isBold, fontStyle);
            if (_fontIds.TryGetValue(key, out var existing))
                return existing;
            var id = _fonts.Count;
            _fonts.Add(new FontEntry(isBold, fontStyle));
            _fontIds[key] = id;
            return id;
        }

        /// <summary>
        /// 获取或创建纯色填充定义索引。
        /// </summary>
        /// <param name="fontStyle">包含填充状态的字体及展示描述。</param>
        /// <returns>填充定义索引；未配置可写入的填充时返回零。</returns>
        private int GetFillId(ExcelDataReaderFailureFontStyle fontStyle)
        {
            if (!IsSolidFill(fontStyle))
                return 0;
            var key = (fontStyle.FillPattern ?? string.Empty) + "|"
                + fontStyle.FillForegroundColorIndex.Value.ToString(CultureInfo.InvariantCulture)
                + "|" + (fontStyle.FillBackgroundColorIndex?.ToString(
                    CultureInfo.InvariantCulture) ?? string.Empty);
            if (_fillIds.TryGetValue(key, out var existing))
                return existing + 2;
            var id = _fills.Count;
            _fills.Add(fontStyle);
            _fillIds[key] = id;
            return id + 2;
        }

        /// <summary>
        /// 获取或创建四边框定义索引。
        /// </summary>
        /// <param name="fontStyle">包含边框状态的字体及展示描述。</param>
        /// <returns>边框定义索引；未配置边框时返回零。</returns>
        private int GetBorderId(ExcelDataReaderFailureFontStyle fontStyle)
        {
            if (fontStyle == null || (fontStyle.BorderLeft == null && fontStyle.BorderRight == null
                && fontStyle.BorderTop == null && fontStyle.BorderBottom == null))
                return 0;
            var key = BorderKey(fontStyle.BorderLeft) + "|" + BorderKey(fontStyle.BorderRight)
                + "|" + BorderKey(fontStyle.BorderTop) + "|" + BorderKey(fontStyle.BorderBottom);
            if (_borderIds.TryGetValue(key, out var existing))
                return existing + 1;
            var id = _borders.Count;
            _borders.Add(fontStyle);
            _borderIds[key] = id;
            return id + 1;
        }

        /// <summary>
        /// 写入自定义数字格式定义。
        /// </summary>
        /// <param name="writer">样式表 XML 写入器。</param>
        private void WriteNumberFormats(XmlWriter writer)
        {
            if (_customFormats.Count == 0)
                return;
            writer.WriteStartElement("numFmts", SpreadsheetNs);
            writer.WriteAttributeString("count", _customFormats.Count.ToString(
                CultureInfo.InvariantCulture));
            for (var index = 0; index < _customFormats.Count; index++)
                Empty(writer, "numFmt", SpreadsheetNs,
                    ("numFmtId", (164 + index).ToString(CultureInfo.InvariantCulture)),
                    ("formatCode", _customFormats[index]));
            writer.WriteEndElement();
        }

        /// <summary>
        /// 写入字体定义。
        /// </summary>
        /// <param name="writer">样式表 XML 写入器。</param>
        private void WriteFonts(XmlWriter writer)
        {
            writer.WriteStartElement("fonts", SpreadsheetNs);
            writer.WriteAttributeString("count", _fonts.Count.ToString(CultureInfo.InvariantCulture));
            foreach (var font in _fonts)
            {
                var style = font.FontStyle;
                writer.WriteStartElement("font", SpreadsheetNs);
                if (font.IsBold || style?.IsBold == true)
                    Empty(writer, "b", SpreadsheetNs);
                if (style?.IsItalic == true)
                    Empty(writer, "i", SpreadsheetNs);
                if (style?.IsStrikeout == true)
                    Empty(writer, "strike", SpreadsheetNs);
                if (style?.IsUnderline == true)
                    Empty(writer, "u", SpreadsheetNs);
                Empty(writer, "sz", SpreadsheetNs, ("val", style?.FontHeightInPoints > 0
                    ? style.FontHeightInPoints.ToString("R", CultureInfo.InvariantCulture) : "11"));
                if (style?.FontColorIndex.HasValue == true)
                    Empty(writer, "color", SpreadsheetNs, ("indexed", style.FontColorIndex.Value
                        .ToString(CultureInfo.InvariantCulture)));
                else
                    Empty(writer, "color", SpreadsheetNs, ("theme", "1"));
                Empty(writer, "name", SpreadsheetNs, ("val", string.IsNullOrWhiteSpace(
                    style?.FontName) ? "Calibri" : style.FontName));
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        /// <summary>
        /// 写入纯色填充定义。
        /// </summary>
        /// <param name="writer">样式表 XML 写入器。</param>
        private void WriteFills(XmlWriter writer)
        {
            writer.WriteStartElement("fills", SpreadsheetNs);
            writer.WriteAttributeString("count", (2 + _fills.Count).ToString(
                CultureInfo.InvariantCulture));
            writer.WriteStartElement("fill", SpreadsheetNs);
            Empty(writer, "patternFill", SpreadsheetNs, ("patternType", "none"));
            writer.WriteEndElement();
            writer.WriteStartElement("fill", SpreadsheetNs);
            Empty(writer, "patternFill", SpreadsheetNs, ("patternType", "gray125"));
            writer.WriteEndElement();
            foreach (var fill in _fills)
            {
                writer.WriteStartElement("fill", SpreadsheetNs);
                writer.WriteStartElement("patternFill", SpreadsheetNs);
                writer.WriteAttributeString("patternType", fill.FillPattern);
                Empty(writer, "fgColor", SpreadsheetNs, ("indexed",
                    fill.FillForegroundColorIndex.Value.ToString(CultureInfo.InvariantCulture)));
                if (fill.FillBackgroundColorIndex.HasValue)
                    Empty(writer, "bgColor", SpreadsheetNs, ("indexed",
                        fill.FillBackgroundColorIndex.Value.ToString(CultureInfo.InvariantCulture)));
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        /// <summary>
        /// 写入四边框定义。
        /// </summary>
        /// <param name="writer">样式表 XML 写入器。</param>
        private void WriteBorders(XmlWriter writer)
        {
            writer.WriteStartElement("borders", SpreadsheetNs);
            writer.WriteAttributeString("count", (1 + _borders.Count).ToString(
                CultureInfo.InvariantCulture));
            WriteEmptyBorder(writer);
            foreach (var border in _borders)
            {
                writer.WriteStartElement("border", SpreadsheetNs);
                WriteBorderSide(writer, "left", border.BorderLeft);
                WriteBorderSide(writer, "right", border.BorderRight);
                WriteBorderSide(writer, "top", border.BorderTop);
                WriteBorderSide(writer, "bottom", border.BorderBottom);
                Empty(writer, "diagonal", SpreadsheetNs);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        /// <summary>
        /// 写入一个单元格格式定义。
        /// </summary>
        /// <param name="writer">样式表 XML 写入器。</param>
        /// <param name="numberFormatId">数字格式索引。</param>
        /// <param name="fontId">字体索引。</param>
        /// <param name="fillId">填充索引。</param>
        /// <param name="borderId">边框索引。</param>
        /// <param name="style">样式展示状态。</param>
        /// <param name="applyNumberFormat">是否显式应用内置数字格式。</param>
        private static void WriteXf(XmlWriter writer, int numberFormatId, int fontId,
            int fillId, int borderId, StyleEntry style, bool applyNumberFormat)
        {
            writer.WriteStartElement("xf", SpreadsheetNs);
            writer.WriteAttributeString("numFmtId", numberFormatId.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("fontId", fontId.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("fillId", fillId.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("borderId", borderId.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("xfId", "0");
            if (applyNumberFormat || numberFormatId >= 164)
                writer.WriteAttributeString("applyNumberFormat", "1");
            if (borderId > 0)
                writer.WriteAttributeString("applyBorder", "1");
            if (style?.FontStyle != null)
            {
                var alignment = style.FontStyle;
                if (!string.IsNullOrWhiteSpace(alignment.HorizontalAlignment)
                    || !string.IsNullOrWhiteSpace(alignment.VerticalAlignment)
                    || alignment.WrapText || alignment.ShrinkToFit
                    || alignment.TextRotation.HasValue || alignment.Indentation > 0)
                    writer.WriteAttributeString("applyAlignment", "1");
                if (alignment.IsLocked.HasValue || alignment.IsHidden.HasValue)
                    writer.WriteAttributeString("applyProtection", "1");
            }
            if (style?.FontStyle != null)
            {
                var alignment = style.FontStyle;
                if (!string.IsNullOrWhiteSpace(alignment.HorizontalAlignment)
                    || !string.IsNullOrWhiteSpace(alignment.VerticalAlignment)
                    || alignment.WrapText || alignment.ShrinkToFit
                    || alignment.TextRotation.HasValue || alignment.Indentation > 0)
                {
                    writer.WriteStartElement("alignment", SpreadsheetNs);
                    WriteOptionalAttribute(writer, "horizontal", alignment.HorizontalAlignment);
                    WriteOptionalAttribute(writer, "vertical", alignment.VerticalAlignment);
                    if (alignment.WrapText)
                        writer.WriteAttributeString("wrapText", "1");
                    if (alignment.ShrinkToFit)
                        writer.WriteAttributeString("shrinkToFit", "1");
                    if (alignment.TextRotation.HasValue)
                        writer.WriteAttributeString("textRotation", alignment.TextRotation.Value
                            .ToString(CultureInfo.InvariantCulture));
                    if (alignment.Indentation > 0)
                        writer.WriteAttributeString("indent", alignment.Indentation.ToString(
                            CultureInfo.InvariantCulture));
                    writer.WriteEndElement();
                }
                if (alignment.IsLocked.HasValue || alignment.IsHidden.HasValue)
                {
                    writer.WriteStartElement("protection", SpreadsheetNs);
                    if (alignment.IsLocked.HasValue)
                        writer.WriteAttributeString("locked", alignment.IsLocked.Value ? "1" : "0");
                    if (alignment.IsHidden.HasValue)
                        writer.WriteAttributeString("hidden", alignment.IsHidden.Value ? "1" : "0");
                    writer.WriteEndElement();
                }
            }
            writer.WriteEndElement();
        }

        /// <summary>
        /// 写入不包含边框线的基础边框定义。
        /// </summary>
        /// <param name="writer">样式表 XML 写入器。</param>
        private static void WriteEmptyBorder(XmlWriter writer)
        {
            writer.WriteStartElement("border", SpreadsheetNs);
            Empty(writer, "left", SpreadsheetNs);
            Empty(writer, "right", SpreadsheetNs);
            Empty(writer, "top", SpreadsheetNs);
            Empty(writer, "bottom", SpreadsheetNs);
            Empty(writer, "diagonal", SpreadsheetNs);
            writer.WriteEndElement();
        }

        /// <summary>
        /// 写入一侧边框及其索引色。
        /// </summary>
        /// <param name="writer">样式表 XML 写入器。</param>
        /// <param name="name">边框方向名称。</param>
        /// <param name="side">边框线型和颜色。</param>
        private static void WriteBorderSide(XmlWriter writer, string name,
            ExcelDataReaderFailureBorderSide side)
        {
            writer.WriteStartElement(name, SpreadsheetNs);
            if (side != null && !string.IsNullOrWhiteSpace(side.Style))
            {
                writer.WriteAttributeString("style", side.Style);
                if (side.ColorIndex.HasValue)
                    Empty(writer, "color", SpreadsheetNs, ("indexed", side.ColorIndex.Value
                        .ToString(CultureInfo.InvariantCulture)));
            }
            writer.WriteEndElement();
        }

        /// <summary>
        /// 过滤不能安全写入最小样式表的数字格式。
        /// </summary>
        /// <param name="value">原始数字格式。</param>
        /// <returns>可写入的格式；无效时返回 null。</returns>
        private static string NormalizeFormat(string value)
            => string.IsNullOrWhiteSpace(value) || string.Equals(value, "General",
                StringComparison.OrdinalIgnoreCase) || value.Length > 255 ? null : value;

        /// <summary>
        /// 判断展示状态是否包含可写入的纯色前景填充。
        /// </summary>
        /// <param name="style">字体及展示状态。</param>
        /// <returns>包含纯色前景填充时返回 true。</returns>
        private static bool IsSolidFill(ExcelDataReaderFailureFontStyle style)
            => style != null && string.Equals(style.FillPattern, "solid",
                StringComparison.OrdinalIgnoreCase) && style.FillForegroundColorIndex.HasValue;

        /// <summary>
        /// 创建单侧边框的稳定去重键。
        /// </summary>
        /// <param name="side">边框线型和颜色。</param>
        /// <returns>边框去重键。</returns>
        private static string BorderKey(ExcelDataReaderFailureBorderSide side)
            => side == null ? string.Empty : (side.Style ?? string.Empty) + ":"
                + (side.ColorIndex?.ToString(CultureInfo.InvariantCulture) ?? string.Empty);

        /// <summary>
        /// 创建字体定义的稳定去重键。
        /// </summary>
        /// <param name="isBold">是否使用加粗字体。</param>
        /// <param name="style">字体及展示状态。</param>
        /// <returns>字体去重键。</returns>
        private static string CreateFontKey(bool isBold,
            ExcelDataReaderFailureFontStyle style)
            => (isBold || style?.IsBold == true ? "1" : "0") + "|"
                + (style?.IsItalic == true ? "1" : "0") + "|"
                + (style?.IsStrikeout == true ? "1" : "0") + "|"
                + (style?.IsUnderline == true ? "1" : "0") + "|"
                + (style?.FontName ?? string.Empty) + "|"
                + (style?.FontHeightInPoints.ToString("R", CultureInfo.InvariantCulture) ?? "0") + "|"
                + (style?.FontColorIndex?.ToString(CultureInfo.InvariantCulture) ?? string.Empty);

        /// <summary>
        /// 创建单元格样式的稳定去重键。
        /// </summary>
        /// <param name="format">单元格数字格式。</param>
        /// <param name="isBold">是否使用加粗字体。</param>
        /// <param name="style">字体及展示状态。</param>
        /// <returns>样式去重键。</returns>
        private static string CreateStyleKey(string format, bool isBold,
            ExcelDataReaderFailureFontStyle style)
            => (format ?? string.Empty) + "\u001f" + CreateFontKey(isBold, style) + "\u001f"
                + (style?.HorizontalAlignment ?? string.Empty) + "|"
                + (style?.VerticalAlignment ?? string.Empty) + "|"
                + (style?.WrapText == true ? "1" : "0") + "|"
                + (style?.ShrinkToFit == true ? "1" : "0") + "|"
                + (style?.TextRotation?.ToString(CultureInfo.InvariantCulture) ?? string.Empty) + "|"
                + (style?.Indentation.ToString(CultureInfo.InvariantCulture) ?? string.Empty) + "|"
                + (style?.FillPattern ?? string.Empty) + "|"
                + (style?.FillForegroundColorIndex?.ToString(CultureInfo.InvariantCulture) ?? string.Empty) + "|"
                + (style?.FillBackgroundColorIndex?.ToString(CultureInfo.InvariantCulture) ?? string.Empty) + "|"
                + BorderKey(style?.BorderLeft) + "|" + BorderKey(style?.BorderRight) + "|"
                + BorderKey(style?.BorderTop) + "|" + BorderKey(style?.BorderBottom) + "|"
                + (style?.IsLocked?.ToString() ?? string.Empty) + "|"
                + (style?.IsHidden?.ToString() ?? string.Empty);
    }

    /// <summary>
    /// 将失败工作簿写入调用方流，或通过提交器原子写入目标路径。
    /// </summary>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">导入期间收集的结构化错误。</param>
    /// <param name="capture">前向读取阶段捕获的表头和失败行。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <param name="fileExportCommitter">路径输出的提交器；为 null 时使用默认提交器。</param>
    /// <param name="sourcePath">暂存的原始 XLSX 路径；AnnotatedOriginal 模式必需。</param>
    /// <param name="summarySheetName">可选的错误摘要工作表名称；为空时使用默认名称。</param>
    internal static void Write(ExcelImportFailureOptions options, IReadOnlyList<ExcelImportError> errors,
        ExcelDataReaderFailureWorkbookCapture capture, CancellationToken cancellationToken,
        IFileExportCommitter fileExportCommitter, string sourcePath = null,
        string summarySheetName = null)
    {
        if (options == null || options.Mode == ExcelImportFailureWorkbookMode.None
            || errors == null || errors.Count == 0)
            return;
        options.Validate();
        using var workbook = Serialize(options, errors, capture, cancellationToken, sourcePath,
            summarySheetName);
        cancellationToken.ThrowIfCancellationRequested();
        string temporaryPath = null;
        Exception primaryException = null;
        try
        {
            if (!string.IsNullOrWhiteSpace(options.TemporaryDirectory))
                temporaryPath = WriteTemporaryFile(workbook, options, cancellationToken,
                    capture?.ProviderName);
            if (temporaryPath == null)
            {
                WriteSerializedWorkbook(workbook, options, cancellationToken, fileExportCommitter);
                return;
            }

            using var temporary = File.OpenRead(temporaryPath);
            WriteSerializedWorkbook(temporary, options, cancellationToken, fileExportCommitter);
        }
        catch (Exception exception)
        {
            primaryException = exception;
            throw;
        }
        finally
        {
            DeleteTemporaryFile(temporaryPath, options, primaryException);
        }
    }

    /// <summary>
    /// 异步将失败工作簿写入调用方流，或通过提交器原子写入目标路径。
    /// </summary>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">导入期间收集的结构化错误。</param>
    /// <param name="capture">前向读取阶段捕获的表头和失败行。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <param name="fileExportCommitter">路径输出的提交器；为 null 时使用默认提交器。</param>
    /// <param name="sourcePath">暂存的原始 XLSX 路径；AnnotatedOriginal 模式必需。</param>
    /// <param name="summarySheetName">可选的错误摘要工作表名称；为空时使用默认名称。</param>
    /// <returns>代表异步写入或提交操作的任务。</returns>
    internal static async Task WriteAsync(ExcelImportFailureOptions options,
        IReadOnlyList<ExcelImportError> errors, ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken, IFileExportCommitter fileExportCommitter,
        string sourcePath = null, string summarySheetName = null)
    {
        if (options == null || options.Mode == ExcelImportFailureWorkbookMode.None
            || errors == null || errors.Count == 0)
            return;
        options.Validate();
        using var workbook = Serialize(options, errors, capture, cancellationToken, sourcePath,
            summarySheetName);
        cancellationToken.ThrowIfCancellationRequested();
        string temporaryPath = null;
        Exception primaryException = null;
        try
        {
            if (!string.IsNullOrWhiteSpace(options.TemporaryDirectory))
                temporaryPath = await WriteTemporaryFileAsync(workbook, options, cancellationToken,
                        capture?.ProviderName)
                    .ConfigureAwait(false);
            if (temporaryPath == null)
            {
                await WriteSerializedWorkbookAsync(workbook, options, cancellationToken, fileExportCommitter)
                    .ConfigureAwait(false);
                return;
            }

            await using var temporary = File.OpenRead(temporaryPath);
            await WriteSerializedWorkbookAsync(temporary, options, cancellationToken, fileExportCommitter)
                .ConfigureAwait(false);
        }
        catch (Exception exception)
        {
            primaryException = exception;
            throw;
        }
        finally
        {
            DeleteTemporaryFile(temporaryPath, options, primaryException);
        }
    }

    /// <summary>
    /// 将序列化结果写入调用方流或路径提交器。
    /// </summary>
    /// <param name="source">已序列化的来源流。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="fileExportCommitter">路径提交器。</param>
    private static void WriteSerializedWorkbook(Stream source, ExcelImportFailureOptions options,
        CancellationToken cancellationToken, IFileExportCommitter fileExportCommitter)
    {
        if (options.Destination != null)
        {
            source.Position = 0;
            CopyTo(source, options.Destination, cancellationToken);
            return;
        }

        var committer = fileExportCommitter ?? new DefaultFileExportCommitter();
        committer.Commit(options.DestinationPath, destination =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            source.Position = 0;
            CopyTo(source, destination, cancellationToken);
        }, cancellationToken, "xlsx");
    }

    /// <summary>
    /// 异步将序列化结果写入调用方流或路径提交器。
    /// </summary>
    /// <param name="source">已序列化的来源流。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="fileExportCommitter">路径提交器。</param>
    private static async Task WriteSerializedWorkbookAsync(Stream source, ExcelImportFailureOptions options,
        CancellationToken cancellationToken, IFileExportCommitter fileExportCommitter)
    {
        if (options.Destination != null)
        {
            source.Position = 0;
            await source.CopyToAsync(options.Destination, 81920, cancellationToken).ConfigureAwait(false);
            return;
        }

        var committer = fileExportCommitter ?? new DefaultFileExportCommitter();
        await committer.CommitAsync(options.DestinationPath, async (destination, token) =>
        {
            token.ThrowIfCancellationRequested();
            source.Position = 0;
            await source.CopyToAsync(destination, 81920, token).ConfigureAwait(false);
        }, cancellationToken, "xlsx").ConfigureAwait(false);
    }

    /// <summary>
    /// 将序列化结果写入请求指定的临时目录。
    /// </summary>
    /// <param name="source">已序列化的内存流。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="providerName">资源诊断使用的 Provider 名称。</param>
    /// <returns>已创建的临时文件路径。</returns>
    private static string WriteTemporaryFile(MemoryStream source, ExcelImportFailureOptions options,
        CancellationToken cancellationToken, string providerName)
    {
        var directory = options.TemporaryDirectory;
        try
        {
            Directory.CreateDirectory(directory);
        }
        catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException)
        {
            throw new BingOfficesImportException("失败工作簿临时目录创建失败。", exception, Provider(providerName),
                BingOfficesStage.Open);
        }

        var path = Path.Combine(directory, $"bing-offices-exceldatareader-failure-{Guid.NewGuid():N}.tmp");
        try
        {
            using var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None,
                81920, FileOptions.SequentialScan);
            source.Position = 0;
            CopyTo(source, output, cancellationToken);
            output.Flush(true);
            return path;
        }
        catch (OperationCanceledException)
        {
            DeleteTemporaryFile(path, options, null);
            throw;
        }
        catch (BingOfficesException)
        {
            DeleteTemporaryFile(path, options, null);
            throw;
        }
        catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException)
        {
            DeleteTemporaryFile(path, options, exception);
            throw new BingOfficesImportException("失败工作簿临时文件写入失败。", exception,
                Provider(providerName), BingOfficesStage.Serialize);
        }
    }

    /// <summary>
    /// 异步将序列化结果写入请求指定的临时目录。
    /// </summary>
    /// <param name="source">已序列化的内存流。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="providerName">资源诊断使用的 Provider 名称。</param>
    /// <returns>已创建的临时文件路径。</returns>
    private static async Task<string> WriteTemporaryFileAsync(MemoryStream source,
        ExcelImportFailureOptions options, CancellationToken cancellationToken, string providerName)
    {
        var directory = options.TemporaryDirectory;
        try
        {
            Directory.CreateDirectory(directory);
        }
        catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException)
        {
            throw new BingOfficesImportException("失败工作簿临时目录创建失败。", exception, Provider(providerName),
                BingOfficesStage.Open);
        }

        var path = Path.Combine(directory, $"bing-offices-exceldatareader-failure-{Guid.NewGuid():N}.tmp");
        try
        {
            await using var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None,
                81920, FileOptions.SequentialScan);
            source.Position = 0;
            await source.CopyToAsync(output, 81920, cancellationToken).ConfigureAwait(false);
            await output.FlushAsync(cancellationToken).ConfigureAwait(false);
            return path;
        }
        catch (OperationCanceledException)
        {
            DeleteTemporaryFile(path, options, null);
            throw;
        }
        catch (BingOfficesException)
        {
            DeleteTemporaryFile(path, options, null);
            throw;
        }
        catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException)
        {
            DeleteTemporaryFile(path, options, exception);
            throw new BingOfficesImportException("失败工作簿临时文件写入失败。", exception,
                Provider(providerName), BingOfficesStage.Serialize);
        }
    }

    /// <summary>
    /// 删除失败工作簿临时文件，并将清理异常交给诊断接收器。
    /// </summary>
    /// <param name="path">待删除的临时文件路径。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="primaryException">原始异常；没有时为 null。</param>
    private static void DeleteTemporaryFile(string path, ExcelImportFailureOptions options,
        Exception primaryException)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch (Exception cleanupException) when (cleanupException is IOException
            || cleanupException is UnauthorizedAccessException)
        {
            if (primaryException != null)
                primaryException.Data["Bing.Offices.FailureWorkbook.TemporaryCleanupException"] = cleanupException;
            try
            {
                options?.DiagnosticSink?.Invoke(new ExcelImportFailureDiagnostic(
                    "FailureWorkbookTemporaryCleanupFailed", path, cleanupException));
            }
            catch
            {
                // 诊断接收器异常不得覆盖失败工作簿的主异常。
            }
        }
    }

    /// <summary>
    /// 将捕获的数据序列化为位置位于起始处的内存流。
    /// </summary>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">导入期间收集的结构化错误。</param>
    /// <param name="capture">前向读取阶段捕获的表头和失败行。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <param name="sourcePath">暂存的原始 XLSX 路径；AnnotatedOriginal 模式必需。</param>
    /// <param name="summarySheetName">可选的错误摘要工作表名称；为空时使用默认名称。</param>
    /// <returns>包含完整 XLSX 包的内存流。</returns>
    internal static MemoryStream Serialize(ExcelImportFailureOptions options,
        IReadOnlyList<ExcelImportError> errors, ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken, string sourcePath = null,
        string summarySheetName = null)
    {
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        if (capture == null)
            throw new ArgumentNullException(nameof(capture));
        var providerName = capture.ProviderName;
        if (options.Mode == ExcelImportFailureWorkbookMode.AnnotatedOriginal)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
                throw new BingOfficesUnsupportedFeatureException(
                    $"{providerName} AnnotatedOriginal 缺少可保留的暂存源文件。", provider: providerName,
                    operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);
            using var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read,
                64 * 1024, FileOptions.SequentialScan);
            return ExcelDataReaderAnnotatedOriginalWriter.Serialize(source, options, errors, capture,
                cancellationToken);
        }
        if (options.Mode != ExcelImportFailureWorkbookMode.ErrorRowsOnly)
            throw new ArgumentException("ExcelDataReader 失败工作簿仅支持 ErrorRowsOnly 或 AnnotatedOriginal 模式。",
                nameof(options));
        ValidateLimits(options);
        errors ??= Array.Empty<ExcelImportError>();
        ValidateColumnMetadata(capture, providerName, cancellationToken);
        ValidateBudgets(options, errors, capture, cancellationToken, providerName);

        var output = new MemoryStream();
        try
        {
            using (var limited = new LimitedWriteStream(output, options.MaxSerializedBytes, providerName))
            using (var archive = new ZipArchive(limited, ZipArchiveMode.Create, leaveOpen: true))
            {
                WritePackage(archive, errors, capture, cancellationToken, summarySheetName,
                    providerName);
            }
            cancellationToken.ThrowIfCancellationRequested();
            output.Position = 0;
            return output;
        }
        catch
        {
            output.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 将捕获的数据序列化为 XLSX 字节数组。
    /// </summary>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">导入期间收集的结构化错误。</param>
    /// <param name="capture">前向读取阶段捕获的表头和失败行。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <param name="summarySheetName">可选的错误摘要工作表名称；为空时使用默认名称。</param>
    /// <returns>完整 XLSX 包的字节数组。</returns>
    internal static byte[] SerializeToBytes(ExcelImportFailureOptions options,
        IReadOnlyList<ExcelImportError> errors, ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken, string summarySheetName = null)
    {
        using var stream = Serialize(options, errors, capture, cancellationToken,
            summarySheetName: summarySheetName);
        return stream.ToArray();
    }

    /// <summary>
    /// 验证输出限制配置中的数值均为正数。
    /// </summary>
    /// <param name="options">待验证的失败工作簿选项。</param>
    private static void ValidateLimits(ExcelImportFailureOptions options)
    {
        if (options.MaxSerializedBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxSerializedBytes));
        if (options.MaxCandidateErrorRows <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxCandidateErrorRows));
        if (options.MaxCopiedCells <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxCopiedCells));
        if (options.MaxEstimatedTargetObjects <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxEstimatedTargetObjects));
    }

    /// <summary>
    /// 在序列化前检查候选行、复制单元格和目标对象预算。
    /// </summary>
    /// <param name="options">失败工作簿资源限制。</param>
    /// <param name="errors">将写入摘要的错误。</param>
    /// <param name="capture">失败行快照。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <param name="providerName">资源诊断使用的 Provider 名称。</param>
    private static void ValidateBudgets(ExcelImportFailureOptions options,
        IReadOnlyList<ExcelImportError> errors, ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken, string providerName)
    {
        long candidateRows = 0;
        long copiedCells = 0;
        long objects = Add(0, capture.Sheets.Count + 1L, options.MaxEstimatedTargetObjects,
            providerName);
        foreach (var sheet in capture.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            candidateRows = Add(candidateRows, sheet.Rows.Count, options.MaxCandidateErrorRows,
                providerName);
            var width = Math.Max(1, sheet.Header.Count + 1);
            if (width > 16_384 || sheet.Rows.Count + 1L > 1_048_576)
                throw ResourceLimit("失败工作表大小超过 XLSX 格式限制。", providerName: providerName);
            var outputRows = sheet.Header.Count > 0 ? sheet.Rows.Count + 1L : sheet.Rows.Count;
            var cells = Multiply(width, outputRows, providerName);
            copiedCells = Add(copiedCells, cells, options.MaxCopiedCells, providerName);
            objects = Add(objects, outputRows + cells + sheet.ColumnMetadata.Count,
                options.MaxEstimatedTargetObjects, providerName);
            foreach (var row in sheet.Rows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (row.RowIndex < 1 || row.RowIndex > 1_048_576 || row.Values.Count + 1 > 16_384)
                    throw ResourceLimit("失败行坐标超出 XLSX 格式范围。", providerName: providerName);
            }
        }

        foreach (var metadataSheet in capture.Metadata.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            objects = Add(objects, metadataSheet.Validations.Count,
                options.MaxEstimatedTargetObjects, providerName);
            foreach (var validation in metadataSheet.Validations)
            {
                cancellationToken.ThrowIfCancellationRequested();
                objects = Add(objects, validation.Ranges.Count,
                    options.MaxEstimatedTargetObjects, providerName);
            }
            foreach (var conditionalFormatting in metadataSheet.ConditionalFormattings)
            {
                cancellationToken.ThrowIfCancellationRequested();
                objects = Add(objects, 1L + conditionalFormatting.Ranges.Count,
                    options.MaxEstimatedTargetObjects, providerName);
            }
            objects = Add(objects, (long)metadataSheet.ColumnRanges.Count
                + metadataSheet.UnsupportedColumnMetadataCount,
                options.MaxEstimatedTargetObjects, providerName);
        }
        objects = Add(objects, capture.Metadata.DefinedNames.Count,
            options.MaxEstimatedTargetObjects, providerName);

        var summaryRows = errors.Count + 1L;
        var summaryCells = Multiply(SummaryHeaders.Length, summaryRows, providerName);
        copiedCells = Add(copiedCells, summaryCells, options.MaxCopiedCells, providerName);
        objects = Add(objects, summaryRows + summaryCells, options.MaxEstimatedTargetObjects, providerName);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 累加数量并在溢出或超出上限时抛出资源限制异常。
    /// </summary>
    /// <param name="current">当前数量。</param>
    /// <param name="addition">本次增加量。</param>
    /// <param name="maximum">可选最大值。</param>
    /// <param name="providerName">资源诊断使用的 Provider 名称。</param>
    /// <returns>累加后的数量。</returns>
    private static long Add(long current, long addition, long? maximum, string providerName)
    {
        long total;
        try
        {
            total = checked(current + addition);
        }
        catch (OverflowException exception)
        {
            throw ResourceLimit("失败工作簿资源估算超出可表示范围。", exception, providerName);
        }
        if (maximum.HasValue && total > maximum.Value)
            throw ResourceLimit($"失败工作簿资源估算超过限制: {maximum.Value}",
                providerName: providerName);
        return total;
    }

    /// <summary>
    /// 计算单元格网格大小并检查乘法溢出。
    /// </summary>
    /// <param name="width">网格列数。</param>
    /// <param name="height">网格行数。</param>
    /// <param name="providerName">资源诊断使用的 Provider 名称。</param>
    /// <returns>网格单元格数量。</returns>
    private static long Multiply(long width, long height, string providerName)
    {
        try
        {
            return checked(width * height);
        }
        catch (OverflowException exception)
        {
            throw ResourceLimit("失败工作簿单元格估算超出可表示范围。", exception, providerName);
        }
    }

    /// <summary>
    /// 创建带有 ExcelDataReader 导入阶段上下文的资源异常。
    /// </summary>
    /// <param name="message">资源限制说明。</param>
    /// <param name="innerException">导致资源异常的内部异常。</param>
    /// <param name="providerName">资源诊断使用的 Provider 名称。</param>
    /// <returns>结构化资源限制异常。</returns>
    private static BingOfficesResourceLimitException ResourceLimit(string message,
        Exception innerException = null, string providerName = null)
        => new BingOfficesResourceLimitException(message, innerException, Provider(providerName),
            BingOfficesOperation.Import, BingOfficesStage.Preflight);

    /// <summary>
    /// 规范化资源诊断使用的 Provider 名称。
    /// </summary>
    /// <param name="providerName">调用方提供的 Provider 名称。</param>
    /// <returns>可用于异常上下文的 Provider 名称。</returns>
    private static string Provider(string providerName)
        => string.IsNullOrWhiteSpace(providerName) ? "ExcelDataReader" : providerName;

    /// <summary>
    /// 写入最小 XLSX 包的目录、工作簿、样式和工作表部件。
    /// </summary>
    /// <param name="archive">目标 ZIP 包。</param>
    /// <param name="errors">结构化错误。</param>
    /// <param name="capture">失败行快照。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <param name="summarySheetName">可选的错误摘要工作表名称；为空时使用默认名称。</param>
    /// <param name="providerName">异常上下文中的 Provider 名称。</param>
    private static void WritePackage(ZipArchive archive, IReadOnlyList<ExcelImportError> errors,
        ExcelDataReaderFailureWorkbookCapture capture, CancellationToken cancellationToken,
        string summarySheetName, string providerName)
    {
        var sourceSheetNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        ValidateConditionalFormattingCoverage(capture, providerName, cancellationToken);
        var sheetNames = BuildSheetNames(capture.Sheets, summarySheetName, sourceSheetNames);
        ValidateConditionalFormattingSheetNames(capture, sourceSheetNames, providerName,
            cancellationToken);
        var definedNames = BuildRemappedDefinedNames(capture, sourceSheetNames,
            cancellationToken);
        var styleRegistry = FailureStyleRegistry.Build(capture, cancellationToken);
        WriteXml(archive, "[Content_Types].xml", writer =>
        {
            writer.WriteStartElement("Types", ContentTypeNs);
            Empty(writer, "Default", ContentTypeNs, ("Extension", "rels"),
                ("ContentType", "application/vnd.openxmlformats-package.relationships+xml"));
            Empty(writer, "Default", ContentTypeNs, ("Extension", "xml"), ("ContentType", "application/xml"));
            Empty(writer, "Override", ContentTypeNs, ("PartName", "/xl/workbook.xml"),
                ("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"));
            Empty(writer, "Override", ContentTypeNs, ("PartName", "/xl/styles.xml"),
                ("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"));
            for (var i = 0; i < sheetNames.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Empty(writer, "Override", ContentTypeNs, ("PartName", $"/xl/worksheets/sheet{i + 1}.xml"),
                    ("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
            }
            writer.WriteEndElement();
        });

        WriteXml(archive, "_rels/.rels", writer =>
        {
            writer.WriteStartElement("Relationships", PackageRelNs);
            Empty(writer, "Relationship", PackageRelNs, ("Id", "rId1"),
                ("Type", OfficeRelNs + "/officeDocument"), ("Target", "xl/workbook.xml"));
            writer.WriteEndElement();
        });

        WriteXml(archive, "xl/workbook.xml", writer =>
        {
            writer.WriteStartElement("workbook", SpreadsheetNs);
            writer.WriteAttributeString("xmlns", "r", null, OfficeRelNs);
            writer.WriteStartElement("sheets", SpreadsheetNs);
            for (var i = 0; i < sheetNames.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                writer.WriteStartElement("sheet", SpreadsheetNs);
                writer.WriteAttributeString("name", sheetNames[i]);
                writer.WriteAttributeString("sheetId", (i + 1).ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("r", "id", OfficeRelNs, $"rId{i + 1}");
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            if (definedNames.Count > 0)
            {
                writer.WriteStartElement("definedNames", SpreadsheetNs);
                foreach (var definedName in definedNames)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    writer.WriteStartElement("definedName", SpreadsheetNs);
                    writer.WriteAttributeString("name", definedName.Name);
                    if (definedName.LocalSheetId.HasValue)
                        writer.WriteAttributeString("localSheetId",
                            definedName.LocalSheetId.Value.ToString(CultureInfo.InvariantCulture));
                    writer.WriteString(definedName.RefersTo);
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        });

        WriteXml(archive, "xl/_rels/workbook.xml.rels", writer =>
        {
            writer.WriteStartElement("Relationships", PackageRelNs);
            for (var i = 0; i < sheetNames.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Empty(writer, "Relationship", PackageRelNs, ("Id", $"rId{i + 1}"),
                    ("Type", OfficeRelNs + "/worksheet"), ("Target", $"worksheets/sheet{i + 1}.xml"));
            }
            Empty(writer, "Relationship", PackageRelNs, ("Id", $"rId{sheetNames.Count + 1}"),
                ("Type", OfficeRelNs + "/styles"), ("Target", "styles.xml"));
            writer.WriteEndElement();
        });

        WriteStyles(archive, styleRegistry);
        var summaryRows = errors.Select(error => new OutputRow(ToErrorRow(error), null, null, null));
        WriteSheet(archive, "xl/worksheets/sheet1.xml", SummaryHeaders.Cast<object>().ToArray(),
            new Dictionary<int, string>(), new Dictionary<int, bool>(),
            new Dictionary<int, ExcelDataReaderFailureFontStyle>(), true, summaryRows,
            new Dictionary<int, ExcelDataReaderFailureColumnMetadata>(),
            Array.Empty<(int FirstColumn, int LastColumn, double Width, bool HasWidth, bool Hidden)>(),
            Array.Empty<OutputValidation>(),
            Array.Empty<string>(),
            styleRegistry, cancellationToken);
        for (var i = 0; i < capture.Sheets.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = capture.Sheets[i];
            var width = Math.Max(1, sheet.Header.Count);
            foreach (var row in sheet.Rows)
                width = Math.Max(width, row.Values.Count);
            var header = new object[width + 1];
            for (var c = 0; c < width; c++)
                header[c] = c < sheet.Header.Count ? sheet.Header[c] : null;
            header[width] = "__Errors";
            var metadata = capture.Metadata.FindSheet(sheet.Name);
            var columnMetadata = BuildOutputColumnMetadata(sheet, width, cancellationToken);
            var columnRanges = BuildOutputColumnRanges(sheet, metadata, width, cancellationToken);
            var validations = BuildRemappedValidations(sheet, metadata, cancellationToken);
            var conditionalFormattings = BuildRemappedConditionalFormattings(sheet,
                metadata, width, cancellationToken, providerName);
            WriteSheet(archive, $"xl/worksheets/sheet{i + 2}.xml", header,
                sheet.HeaderFormats, sheet.HeaderBoldColumns, sheet.HeaderFontStyles,
                sheet.Header.Count > 0, BuildRows(sheet, errors, width, cancellationToken),
                columnMetadata,
                columnRanges,
                validations,
                conditionalFormattings,
                styleRegistry, cancellationToken);
        }
    }

    /// <summary>
    /// 验证失败工作表的源列元数据属于有限复制范围。
    /// </summary>
    /// <param name="capture">失败行快照。</param>
    /// <param name="providerName">异常上下文中的 Provider 名称。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void ValidateColumnMetadata(ExcelDataReaderFailureWorkbookCapture capture,
        string providerName, CancellationToken cancellationToken)
    {
        foreach (var sheet in capture.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var metadata = capture.Metadata.FindSheet(sheet.Name);
            if (metadata?.UnsupportedColumnMetadataCount > 0)
                throw UnsupportedColumnMetadata(providerName, sheet.Name);
        }
    }

    /// <summary>
    /// 保留源数据列范围内捕获到的列元数据。
    /// </summary>
    /// <param name="sheet">失败行快照。</param>
    /// <param name="sourceWidth">输出中的源数据列数。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>裁剪到源数据列的列元数据快照。</returns>
    private static IReadOnlyDictionary<int, ExcelDataReaderFailureColumnMetadata>
        BuildOutputColumnMetadata(ExcelDataReaderFailureSheet sheet, int sourceWidth,
            CancellationToken cancellationToken)
    {
        if (sheet?.ColumnMetadata == null || sheet.ColumnMetadata.Count == 0 || sourceWidth <= 0)
            return new Dictionary<int, ExcelDataReaderFailureColumnMetadata>();

        var output = new Dictionary<int, ExcelDataReaderFailureColumnMetadata>();
        foreach (var pair in sheet.ColumnMetadata)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (pair.Key >= 0 && pair.Key < sourceWidth)
                output.Add(pair.Key, pair.Value);
        }
        return output;
    }

    /// <summary>
    /// 保留与源数据列相交的列宽和隐藏状态，不延伸到新增错误列。
    /// </summary>
    /// <param name="sheet">失败行快照。</param>
    /// <param name="metadata">源工作表列元数据。</param>
    /// <param name="sourceWidth">输出中的源数据列数。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>裁剪到源数据列的列元数据快照。</returns>
    private static IReadOnlyList<(int FirstColumn, int LastColumn, double Width, bool HasWidth,
        bool Hidden)> BuildOutputColumnRanges(ExcelDataReaderFailureSheet sheet,
        ExcelXlsxFailureSheetMetadata metadata, int sourceWidth,
        CancellationToken cancellationToken)
    {
        if (sheet?.ColumnMetadataCaptured == true || metadata == null
            || metadata.ColumnRanges.Count == 0 || sourceWidth <= 0)
            return Array.Empty<(int FirstColumn, int LastColumn, double Width, bool HasWidth,
                bool Hidden)>();

        var output = new List<(int FirstColumn, int LastColumn, double Width, bool HasWidth,
            bool Hidden)>();
        foreach (var range in metadata.ColumnRanges)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (range.FirstColumn > sourceWidth)
                continue;
            output.Add((range.FirstColumn, Math.Min(range.LastColumn, sourceWidth), range.Width,
                range.HasWidth, range.Hidden));
        }
        return output.AsReadOnly();
    }

    /// <summary>
    /// 验证所有条件格式工作表都有可重映射的失败行快照。
    /// </summary>
    /// <param name="capture">失败行快照。</param>
    /// <param name="providerName">异常上下文中的 Provider 名称。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void ValidateConditionalFormattingCoverage(
        ExcelDataReaderFailureWorkbookCapture capture, string providerName,
        CancellationToken cancellationToken)
    {
        if (capture?.Metadata?.HasConditionalFormatting != true)
            return;
        foreach (var metadata in capture.Metadata.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (metadata?.ConditionalFormattingCount <= 0)
                continue;
            var sheet = capture.Sheets.FirstOrDefault(item => string.Equals(item.Name,
                metadata.Name, StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
                throw UnsupportedConditionalFormatting(providerName);
        }
    }

    /// <summary>
    /// 验证条件格式涉及的工作表名称未因最小输出而变化。
    /// </summary>
    /// <param name="capture">失败行快照。</param>
    /// <param name="sourceSheetNames">源工作表到输出工作表的名称映射。</param>
    /// <param name="providerName">异常上下文中的 Provider 名称。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void ValidateConditionalFormattingSheetNames(
        ExcelDataReaderFailureWorkbookCapture capture,
        IReadOnlyDictionary<string, string> sourceSheetNames, string providerName,
        CancellationToken cancellationToken)
    {
        if (capture?.Metadata?.HasConditionalFormatting != true
            || sourceSheetNames == null)
            return;
        foreach (var pair in sourceSheetNames)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.Equals(pair.Key, pair.Value, StringComparison.Ordinal))
                throw UnsupportedConditionalFormatting(providerName);
        }
    }

    /// <summary>
    /// 验证条件格式区域在低保留输出中仍保持原始坐标。
    /// </summary>
    /// <param name="sheet">失败行工作表快照。</param>
    /// <param name="metadata">源工作表条件格式元数据。</param>
    /// <param name="width">输出源数据列数。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="providerName">异常上下文中的 Provider 名称。</param>
    /// <returns>可直接写回的条件格式 XML。</returns>
    private static IReadOnlyList<string> BuildRemappedConditionalFormattings(
        ExcelDataReaderFailureSheet sheet, ExcelXlsxFailureSheetMetadata metadata, int width,
        CancellationToken cancellationToken, string providerName)
    {
        if (metadata == null || metadata.ConditionalFormattingCount == 0)
            return Array.Empty<string>();
        if (sheet == null || metadata.ConditionalFormattings.Count
            != metadata.ConditionalFormattingCount)
            throw UnsupportedConditionalFormatting(providerName);

        var sourceToOutput = new Dictionary<int, int> { [0] = 0 };
        for (var index = 0; index < sheet.Rows.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sourceRow = sheet.Rows[index].RowIndex - 1;
            if (sourceRow < 0 || !sourceToOutput.TryAdd(sourceRow, index + 1)
                || sourceRow != index + 1)
                throw UnsupportedConditionalFormatting(providerName);
        }

        var output = new List<string>(metadata.ConditionalFormattings.Count);
        foreach (var conditionalFormatting in metadata.ConditionalFormattings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (conditionalFormatting.Ranges.Any(range => range.LastColumn >= width
                || range.FirstColumn < 0
                || RangeRowsOutsideIdentity(range, sourceToOutput)))
                throw UnsupportedConditionalFormatting(providerName);
            output.Add(conditionalFormatting.Xml);
        }
        return output;
    }

    /// <summary>
    /// 判断条件格式覆盖的行是否都保持原始零基坐标。
    /// </summary>
    /// <param name="range">条件格式区域。</param>
    /// <param name="sourceToOutput">源行到输出行的映射。</param>
    /// <returns>存在缺失或错位行时返回 <see langword="true" />。</returns>
    private static bool RangeRowsOutsideIdentity(ExcelXlsxFailureRange range,
        IReadOnlyDictionary<int, int> sourceToOutput)
    {
        for (var row = range.FirstRow; row <= range.LastRow; row++)
            if (!sourceToOutput.TryGetValue(row, out var outputRow) || outputRow != row)
                return true;
        return false;
    }

    /// <summary>
    /// 创建条件格式不支持异常。
    /// </summary>
    /// <param name="providerName">异常上下文中的 Provider 名称。</param>
    /// <returns>结构化不支持异常。</returns>
    private static BingOfficesUnsupportedFeatureException UnsupportedConditionalFormatting(
        string providerName)
        => new BingOfficesUnsupportedFeatureException(
            $"{Provider(providerName)} ErrorRowsOnly 仅支持在源坐标不变时保留 XLSX 条件格式；"
            + "请使用 AnnotatedOriginal。", provider: Provider(providerName),
            operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);

    /// <summary>
    /// 创建超出有限列元数据复制范围的不支持异常。
    /// </summary>
    /// <param name="providerName">异常上下文中的 Provider 名称。</param>
    /// <param name="sheetName">源工作表名称。</param>
    /// <returns>结构化不支持异常。</returns>
    private static BingOfficesUnsupportedFeatureException UnsupportedColumnMetadata(
        string providerName, string sheetName)
        => new BingOfficesUnsupportedFeatureException(
            $"{Provider(providerName)} ErrorRowsOnly 仅保留工作表“{sheetName}”源数据列的显式列宽和隐藏状态；"
            + "其他列元数据暂不支持。", provider: Provider(providerName),
            operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);

    /// <summary>
    /// 按失败行输出顺序重映射原生校验区域。
    /// </summary>
    /// <param name="sheet">失败行工作表快照。</param>
    /// <param name="metadata">源工作表校验元数据。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>可写入输出工作表的校验规则。</returns>
    private static IReadOnlyList<OutputValidation> BuildRemappedValidations(
        ExcelDataReaderFailureSheet sheet, ExcelXlsxFailureSheetMetadata metadata,
        CancellationToken cancellationToken)
    {
        if (sheet == null || metadata == null || metadata.Validations.Count == 0
            || sheet.Rows.Count == 0)
            return Array.Empty<OutputValidation>();

        var output = new List<OutputValidation>(metadata.Validations.Count);
        foreach (var validation in metadata.Validations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var references = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            for (var index = 0; index < sheet.Rows.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var sourceRow = sheet.Rows[index].RowIndex - 1;
                foreach (var range in validation.Ranges)
                {
                    if (sourceRow < range.FirstRow || sourceRow > range.LastRow)
                        continue;
                    var outputRow = index + 2;
                    var first = CellRef(range.FirstColumn, outputRow);
                    var last = CellRef(range.LastColumn, outputRow);
                    var reference = string.Equals(first, last, StringComparison.Ordinal)
                        ? first
                        : first + ":" + last;
                    if (seen.Add(reference))
                        references.Add(reference);
                }
            }
            if (references.Count > 0)
                output.Add(new OutputValidation(validation, references));
        }
        return output;
    }

    /// <summary>
    /// 按失败行输出顺序重映射可安全复制的定义名称。
    /// </summary>
    /// <param name="capture">失败工作簿捕获快照。</param>
    /// <param name="sourceSheetNames">源工作表到输出工作表的名称映射。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>可写入输出工作簿的定义名称。</returns>
    private static IReadOnlyList<OutputDefinedName> BuildRemappedDefinedNames(
        ExcelDataReaderFailureWorkbookCapture capture,
        IReadOnlyDictionary<string, string> sourceSheetNames,
        CancellationToken cancellationToken)
    {
        if (capture == null || capture.Metadata.DefinedNames.Count == 0
            || sourceSheetNames == null || sourceSheetNames.Count == 0)
            return Array.Empty<OutputDefinedName>();

        var output = new List<OutputDefinedName>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var definedName in capture.Metadata.DefinedNames)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var targetIndex = FindSheetIndex(capture.Sheets, definedName.TargetSheetName);
            if (targetIndex < 0
                || !sourceSheetNames.TryGetValue(definedName.TargetSheetName,
                    out var outputSheetName))
                continue;
            var targetSheet = capture.Sheets[targetIndex];
            if (!TryRemapDefinedNameRange(targetSheet, definedName.Range,
                    outputSheetName, cancellationToken, out var refersTo))
                continue;

            int? localSheetId = null;
            if (!string.IsNullOrWhiteSpace(definedName.LocalSheetName))
            {
                var localIndex = FindSheetIndex(capture.Sheets, definedName.LocalSheetName);
                if (localIndex < 0
                    || !sourceSheetNames.ContainsKey(definedName.LocalSheetName))
                    continue;
                // Summary sheet is inserted at index zero in the output workbook.
                localSheetId = localIndex + 1;
            }

            var key = definedName.Name + "\u001f" + (localSheetId?.ToString(
                CultureInfo.InvariantCulture) ?? string.Empty);
            if (seen.Add(key))
                output.Add(new OutputDefinedName(definedName.Name, refersTo, localSheetId));
        }
        return output;
    }

    /// <summary>
    /// 将定义名称的一维源区域映射到失败行输出区域。
    /// </summary>
    /// <param name="sheet">源失败行工作表。</param>
    /// <param name="range">源区域。</param>
    /// <param name="outputSheetName">输出工作表名称。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <param name="refersTo">映射后的区域引用。</param>
    /// <returns>区域完整保留且可安全映射时返回 <see langword="true" />。</returns>
    private static bool TryRemapDefinedNameRange(ExcelDataReaderFailureSheet sheet,
        ExcelXlsxFailureRange range, string outputSheetName,
        CancellationToken cancellationToken, out string refersTo)
    {
        refersTo = null;
        var rowMap = new Dictionary<int, int>();
        for (var index = 0; index < sheet.Rows.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            rowMap[sheet.Rows[index].RowIndex] = index + 2;
        }

        var firstOutputRow = 0;
        var lastOutputRow = 0;
        if (range.FirstRow == range.LastRow)
        {
            if (!rowMap.TryGetValue(range.FirstRow + 1, out firstOutputRow))
                return false;
            lastOutputRow = firstOutputRow;
        }
        else if (range.FirstColumn == range.LastColumn)
        {
            var outputRows = new List<int>(range.LastRow - range.FirstRow + 1);
            for (var sourceRow = range.FirstRow; sourceRow <= range.LastRow; sourceRow++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!rowMap.TryGetValue(sourceRow + 1, out var outputRow))
                    return false;
                outputRows.Add(outputRow);
            }
            if (outputRows.Count == 0)
                return false;
            for (var index = 1; index < outputRows.Count; index++)
                if (outputRows[index] != outputRows[index - 1] + 1)
                    return false;
            firstOutputRow = outputRows[0];
            lastOutputRow = outputRows[outputRows.Count - 1];
        }
        else
        {
            return false;
        }

        var first = AbsoluteCellRef(range.FirstColumn, firstOutputRow);
        var last = AbsoluteCellRef(range.LastColumn, lastOutputRow);
        refersTo = string.Equals(first, last, StringComparison.Ordinal)
            ? QuoteSheet(outputSheetName) + "!" + first
            : QuoteSheet(outputSheetName) + "!" + first + ":" + last;
        return true;
    }

    /// <summary>
    /// 查找捕获工作表的序号。
    /// </summary>
    /// <param name="sheets">失败工作表集合。</param>
    /// <param name="name">工作表名称。</param>
    /// <returns>工作表序号；不存在时返回负数。</returns>
    private static int FindSheetIndex(IReadOnlyList<ExcelDataReaderFailureSheet> sheets,
        string name)
    {
        if (sheets == null || string.IsNullOrWhiteSpace(name))
            return -1;
        for (var index = 0; index < sheets.Count; index++)
            if (string.Equals(sheets[index].Name, name, StringComparison.OrdinalIgnoreCase))
                return index;
        return -1;
    }

    /// <summary>
    /// 将工作表名称编码为定义名称公式中的绝对引用前缀。
    /// </summary>
    /// <param name="name">输出工作表名称。</param>
    /// <returns>带单引号的工作表名称。</returns>
    private static string QuoteSheet(string name)
        => "'" + (name ?? string.Empty).Replace("'", "''", StringComparison.Ordinal) + "'";

    /// <summary>
    /// 创建带绝对行列标记的单元格地址。
    /// </summary>
    /// <param name="column">零基列号。</param>
    /// <param name="row">从一开始的行号。</param>
    /// <returns>绝对 A1 单元格地址。</returns>
    private static string AbsoluteCellRef(int column, int row)
    {
        var reference = CellRef(column, row);
        var digitIndex = reference.IndexOfAny("0123456789".ToCharArray());
        return "$" + reference.Substring(0, digitIndex) + "$" + reference.Substring(digitIndex);
    }

    /// <summary>
    /// 按需生成失败工作簿中的源数据行。
    /// </summary>
    /// <param name="sheet">失败行工作表快照。</param>
    /// <param name="errors">本次导入的结构化错误。</param>
    /// <param name="width">输出源数据列数。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>按顺序排列的输出行。</returns>
    private static IEnumerable<OutputRow> BuildRows(
        ExcelDataReaderFailureSheet sheet, IReadOnlyList<ExcelImportError> errors, int width,
        CancellationToken cancellationToken)
    {
        foreach (var row in sheet.Rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var values = new object[width + 1];
            for (var c = 0; c < width; c++)
                values[c] = c < row.Values.Count ? row.Values[c] : null;
            values[width] = RowErrors(row, errors, sheet.Name);
            yield return new OutputRow(values, row.NumberFormats, row.BoldColumns, row.FontStyles);
        }
    }

    /// <summary>
    /// 创建摘要工作表和源工作表的唯一名称集合。
    /// </summary>
    /// <param name="sheets">捕获的源工作表。</param>
    /// <param name="summarySheetName">可选的错误摘要工作表名称。</param>
    /// <param name="sourceSheetNames">用于回写定义名称的源到输出工作表名称映射。</param>
    /// <returns>摘要工作表优先排列的名称。</returns>
    private static List<string> BuildSheetNames(IReadOnlyList<ExcelDataReaderFailureSheet> sheets,
        string summarySheetName, IDictionary<string, string> sourceSheetNames)
    {
        var result = new List<string>(sheets.Count + 1)
        {
            string.IsNullOrWhiteSpace(summarySheetName) ? SummarySheetName : summarySheetName
        };
        foreach (var sheet in sheets)
        {
            var baseName = SanitizeSheetName(sheet.Name);
            var name = baseName;
            var suffix = 1;
            while (result.Any(existing => string.Equals(existing, name, StringComparison.OrdinalIgnoreCase)))
            {
                var tail = "_" + suffix.ToString(CultureInfo.InvariantCulture);
                suffix++;
                name = baseName.Substring(0, Math.Min(baseName.Length, 31 - tail.Length)) + tail;
            }
            result.Add(name);
            sourceSheetNames?.TryAdd(sheet.Name, name);
        }
        return result;
    }

    /// <summary>
    /// 清理非法字符并限制工作表名称长度。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <returns>符合 XLSX 名称限制的文本。</returns>
    private static string SanitizeSheetName(string name)
    {
        var sanitized = new string((name ?? "Sheet").Select(c => "[]:*?/\\".IndexOf(c) >= 0
            || char.IsControl(c) ? '_' : c).ToArray()).Trim('\'');
        if (string.IsNullOrWhiteSpace(sanitized))
            sanitized = "Sheet";
        return sanitized.Substring(0, Math.Min(31, sanitized.Length));
    }

    /// <summary>
    /// 将结构化错误转换为摘要行。
    /// </summary>
    /// <param name="error">待输出的错误。</param>
    /// <returns>包含错误坐标和值的单元格序列。</returns>
    private static IReadOnlyList<object> ToErrorRow(ExcelImportError error)
        => new object[] { error?.SheetName, error?.RowIndex, error?.ColumnIndex, error?.Code.ToString(),
            error?.Message, error?.PropertyName, error?.ColumnKey, error?.Header, error?.RawValue,
            error?.FirstRowNumber };

    /// <summary>
    /// 生成失败数据行的错误说明文本。
    /// </summary>
    /// <param name="row">失败行快照。</param>
    /// <param name="allErrors">本次导入的全部错误。</param>
    /// <param name="sheetName">源工作表名称。</param>
    /// <returns>当前行的错误代码与消息。</returns>
    private static string RowErrors(ExcelDataReaderFailureRow row,
        IReadOnlyList<ExcelImportError> allErrors, string sheetName)
    {
        var rowErrors = row.Errors?.Where(error => error != null).ToArray();
        if (rowErrors == null || rowErrors.Length == 0)
        {
            rowErrors = allErrors.Where(error => error != null && error.RowIndex == row.RowIndex
                && string.Equals(error.SheetName, sheetName, StringComparison.Ordinal)).ToArray();
        }
        return string.Join("; ", rowErrors.Select(error => $"{error.Code}: {error.Message}"));
    }

    /// <summary>
    /// 写入失败工作簿使用的最小样式表。
    /// </summary>
    /// <param name="archive">目标 ZIP 包。</param>
    /// <param name="registry">失败行样式注册表。</param>
    private static void WriteStyles(ZipArchive archive, FailureStyleRegistry registry)
    {
        WriteXml(archive, "xl/styles.xml", writer => registry.Write(writer));
    }

    /// <summary>
    /// 写入带表头和数据行的 SpreadsheetML 工作表部件。
    /// </summary>
    /// <param name="archive">目标 ZIP 包。</param>
    /// <param name="part">工作表部件路径。</param>
    /// <param name="header">输出表头。</param>
    /// <param name="headerFormats">表头单元格数字格式。</param>
    /// <param name="headerBoldColumns">表头单元格加粗状态。</param>
    /// <param name="headerFontStyles">表头单元格字体及展示状态。</param>
    /// <param name="includeHeader">是否写入表头行。</param>
    /// <param name="rows">输出数据行。</param>
    /// <param name="columnMetadata">源工作表的列展示元数据。</param>
    /// <param name="columnRanges">可安全保留的源列宽和隐藏状态。</param>
    /// <param name="validations">按失败行重映射的原生校验规则。</param>
    /// <param name="conditionalFormattings">源坐标不变时保留的条件格式 XML。</param>
    /// <param name="styleRegistry">输出样式注册表。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void WriteSheet(ZipArchive archive, string part, IReadOnlyList<object> header,
        IReadOnlyDictionary<int, string> headerFormats,
        IReadOnlyDictionary<int, bool> headerBoldColumns,
        IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> headerFontStyles,
        bool includeHeader,
        IEnumerable<OutputRow> rows,
        IReadOnlyDictionary<int, ExcelDataReaderFailureColumnMetadata> columnMetadata,
        IReadOnlyList<(int FirstColumn, int LastColumn, double Width, bool HasWidth, bool Hidden)>
            columnRanges,
        IReadOnlyList<OutputValidation> validations,
        IReadOnlyList<string> conditionalFormattings,
        FailureStyleRegistry styleRegistry,
        CancellationToken cancellationToken)
    {
        WriteXml(archive, part, writer =>
        {
            writer.WriteStartElement("worksheet", SpreadsheetNs);
            WriteColumns(writer, columnMetadata, columnRanges);
            writer.WriteStartElement("sheetData", SpreadsheetNs);
            var rowIndex = 1;
            if (includeHeader)
            {
                WriteRow(writer, rowIndex++, header, headerFormats, headerBoldColumns,
                    headerFontStyles, styleRegistry, cancellationToken);
            }
            foreach (var row in rows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                WriteRow(writer, rowIndex++, row.Values, row.NumberFormats, row.BoldColumns,
                    row.FontStyles, styleRegistry, cancellationToken);
            }
            writer.WriteEndElement();
            WriteConditionalFormattings(writer, conditionalFormattings, cancellationToken);
            WriteDataValidations(writer, validations, cancellationToken);
            writer.WriteEndElement();
        });
    }

    /// <summary>
    /// 写入源坐标不变的条件格式节点。
    /// </summary>
    /// <param name="writer">工作表 XML 写入器。</param>
    /// <param name="conditionalFormattings">条件格式 XML。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void WriteConditionalFormattings(XmlWriter writer,
        IReadOnlyList<string> conditionalFormattings, CancellationToken cancellationToken)
    {
        if (conditionalFormattings == null || conditionalFormattings.Count == 0)
            return;
        foreach (var xml in conditionalFormattings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            writer.WriteRaw(xml);
        }
    }

    /// <summary>
    /// 写入重映射后的原生数据验证规则。
    /// </summary>
    /// <param name="writer">工作表 XML 写入器。</param>
    /// <param name="validations">待写入的校验规则。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void WriteDataValidations(XmlWriter writer,
        IReadOnlyList<OutputValidation> validations, CancellationToken cancellationToken)
    {
        if (validations == null || validations.Count == 0)
            return;
        writer.WriteStartElement("dataValidations", SpreadsheetNs);
        writer.WriteAttributeString("count", validations.Count.ToString(CultureInfo.InvariantCulture));
        foreach (var validation in validations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var source = validation.Source;
            writer.WriteStartElement("dataValidation", SpreadsheetNs);
            WriteOptionalAttribute(writer, "type", source.Type);
            WriteOptionalAttribute(writer, "errorStyle", source.ErrorStyle);
            WriteOptionalAttribute(writer, "operator", source.Operator);
            if (source.AllowBlank)
                writer.WriteAttributeString("allowBlank", "1");
            WriteOptionalAttribute(writer, "showDropDown", source.ShowDropDown);
            if (source.ShowInputMessage)
                writer.WriteAttributeString("showInputMessage", "1");
            if (source.ShowErrorMessage)
                writer.WriteAttributeString("showErrorMessage", "1");
            WriteOptionalAttribute(writer, "errorTitle", source.ErrorTitle);
            WriteOptionalAttribute(writer, "error", source.Error);
            WriteOptionalAttribute(writer, "promptTitle", source.PromptTitle);
            WriteOptionalAttribute(writer, "prompt", source.Prompt);
            writer.WriteAttributeString("sqref", string.Join(" ", validation.Sqref));
            if (source.Formula1 != null)
                writer.WriteElementString("formula1", SpreadsheetNs, source.Formula1);
            if (source.Formula2 != null)
                writer.WriteElementString("formula2", SpreadsheetNs, source.Formula2);
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入非空工作表校验属性。
    /// </summary>
    /// <param name="writer">工作表 XML 写入器。</param>
    /// <param name="name">属性名称。</param>
    /// <param name="value">属性值。</param>
    private static void WriteOptionalAttribute(XmlWriter writer, string name, string value)
    {
        if (!string.IsNullOrEmpty(value))
            writer.WriteAttributeString(name, value);
    }

    /// <summary>
    /// 写入源工作表的列宽和隐藏状态。
    /// </summary>
    /// <param name="writer">工作表 XML 写入器。</param>
    /// <param name="columnMetadata">源工作表的列展示元数据。</param>
    /// <param name="columnRanges">源列范围元数据。</param>
    private static void WriteColumns(XmlWriter writer,
        IReadOnlyDictionary<int, ExcelDataReaderFailureColumnMetadata> columnMetadata,
        IReadOnlyList<(int FirstColumn, int LastColumn, double Width, bool HasWidth, bool Hidden)>
            columnRanges)
    {
        if ((columnMetadata == null || columnMetadata.Count == 0)
            && (columnRanges == null || columnRanges.Count == 0))
            return;

        var columns = new List<(int FirstColumn, int LastColumn, double Width, bool HasWidth,
            bool Hidden, bool CustomWidth)>();
        if (columnMetadata != null)
        {
            foreach (var pair in columnMetadata)
                columns.Add((pair.Key + 1, pair.Key + 1, pair.Value.Width, pair.Value.HasWidth,
                    pair.Value.Hidden, true));
        }
        if (columnRanges != null)
        {
            foreach (var range in columnRanges)
                columns.Add((range.FirstColumn, range.LastColumn, range.Width, range.HasWidth,
                    range.Hidden, range.HasWidth));
        }
        columns.Sort((left, right) => left.FirstColumn.CompareTo(right.FirstColumn));

        writer.WriteStartElement("cols", SpreadsheetNs);
        foreach (var column in columns)
        {
            writer.WriteStartElement("col", SpreadsheetNs);
            writer.WriteAttributeString("min", column.FirstColumn.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("max", column.LastColumn.ToString(CultureInfo.InvariantCulture));
            if (column.HasWidth)
                writer.WriteAttributeString("width", column.Width.ToString("0.########",
                    CultureInfo.InvariantCulture));
            if (column.Hidden)
                writer.WriteAttributeString("hidden", "1");
            if (column.CustomWidth)
                writer.WriteAttributeString("customWidth", "1");
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入一行单元格。
    /// </summary>
    /// <param name="writer">工作表 XML 写入器。</param>
    /// <param name="rowIndex">从一开始的输出行号。</param>
    /// <param name="values">按列排列的值。</param>
    /// <param name="numberFormats">单元格数字格式。</param>
    /// <param name="boldColumns">单元格加粗状态。</param>
    /// <param name="fontStyles">单元格字体及展示状态。</param>
    /// <param name="styleRegistry">输出样式注册表。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void WriteRow(XmlWriter writer, int rowIndex, IReadOnlyList<object> values,
        IReadOnlyDictionary<int, string> numberFormats,
        IReadOnlyDictionary<int, bool> boldColumns,
        IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> fontStyles,
        FailureStyleRegistry styleRegistry, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        writer.WriteStartElement("row", SpreadsheetNs);
        writer.WriteAttributeString("r", rowIndex.ToString(CultureInfo.InvariantCulture));
        for (var i = 0; i < values.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var styleIndex = styleRegistry?.GetStyleIndex(i, numberFormats, boldColumns, fontStyles) ?? 0;
            WriteCell(writer, CellRef(i, rowIndex), values[i] ?? string.Empty, styleIndex);
        }
        writer.WriteEndElement();
    }

    /// <summary>
    /// 根据原始值类型写入单元格 XML。
    /// </summary>
    /// <param name="writer">工作表 XML 写入器。</param>
    /// <param name="reference">A1 格式单元格地址。</param>
    /// <param name="value">原始单元格值。</param>
    /// <param name="styleIndex">样式表中的单元格样式索引。</param>
    private static void WriteCell(XmlWriter writer, string reference, object value, int styleIndex)
    {
        writer.WriteStartElement("c", SpreadsheetNs);
        writer.WriteAttributeString("r", reference);
        if (styleIndex > 0)
            writer.WriteAttributeString("s", styleIndex.ToString(CultureInfo.InvariantCulture));
        switch (value)
        {
            case bool boolean:
                writer.WriteAttributeString("t", "b");
                Value(writer, boolean ? "1" : "0");
                break;
            case DateTime date:
                if (styleIndex <= 0)
                    writer.WriteAttributeString("s", "1");
                Value(writer, date.ToOADate().ToString("R", CultureInfo.InvariantCulture));
                break;
            case DateTimeOffset date:
                if (styleIndex <= 0)
                    writer.WriteAttributeString("s", "1");
                Value(writer, date.UtcDateTime.ToOADate().ToString("R", CultureInfo.InvariantCulture));
                break;
            case string text:
                InlineString(writer, text);
                break;
            case char character:
                InlineString(writer, character.ToString());
                break;
            case byte or sbyte or short or ushort or int or uint or long or ulong or decimal:
                Value(writer, Convert.ToString(value, CultureInfo.InvariantCulture));
                break;
            case float single when float.IsFinite(single):
                Value(writer, single.ToString("R", CultureInfo.InvariantCulture));
                break;
            case double number when double.IsFinite(number):
                Value(writer, number.ToString("R", CultureInfo.InvariantCulture));
                break;
            default:
                InlineString(writer, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                break;
        }
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入 inline string 并保留首尾空白。
    /// </summary>
    /// <param name="writer">工作表 XML 写入器。</param>
    /// <param name="text">单元格文本。</param>
    private static void InlineString(XmlWriter writer, string text)
    {
        writer.WriteAttributeString("t", "inlineStr");
        writer.WriteStartElement("is", SpreadsheetNs);
        writer.WriteStartElement("t", SpreadsheetNs);
        if (text.Length != text.Trim().Length)
            writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
        writer.WriteString(text);
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入数值、日期序列或布尔值节点。
    /// </summary>
    /// <param name="writer">工作表 XML 写入器。</param>
    /// <param name="value">InvariantCulture 格式的单元格值。</param>
    private static void Value(XmlWriter writer, string value)
    {
        writer.WriteStartElement("v", SpreadsheetNs);
        writer.WriteString(value);
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入带属性的 XML 元素。
    /// </summary>
    /// <param name="writer">目标 XML 写入器。</param>
    /// <param name="name">元素本地名称。</param>
    /// <param name="ns">元素命名空间。</param>
    /// <param name="attributes">元素属性。</param>
    private static void Empty(XmlWriter writer, string name, string ns, params (string Name, string Value)[] attributes)
    {
        writer.WriteStartElement(name, ns);
        foreach (var attribute in attributes)
            writer.WriteAttributeString(attribute.Name, attribute.Value);
        writer.WriteEndElement();
    }

    /// <summary>
    /// 创建 ZIP 部件并以 UTF-8 XML 写入内容。
    /// </summary>
    /// <param name="archive">目标 XLSX 压缩包。</param>
    /// <param name="part">ZIP 部件名称。</param>
    /// <param name="write">写入部件 XML 的委托。</param>
    private static void WriteXml(ZipArchive archive, string part, Action<XmlWriter> write)
    {
        var entry = archive.CreateEntry(part, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            CloseOutput = false,
            OmitXmlDeclaration = false,
            Indent = false
        });
        writer.WriteStartDocument();
        write(writer);
        writer.WriteEndDocument();
    }

    /// <summary>
    /// 分块复制序列化结果并在每块之间检查取消令牌。
    /// </summary>
    /// <param name="source">已序列化的内存流。</param>
    /// <param name="destination">调用方流或提交暂存流。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void CopyTo(Stream source, Stream destination, CancellationToken cancellationToken)
    {
        var buffer = new byte[81920];
        int read;
        while ((read = source.Read(buffer, 0, buffer.Length)) != 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            destination.Write(buffer, 0, read);
        }
    }

    /// <summary>
    /// 将零起始列号和一基行号转换为 A1 单元格地址。
    /// </summary>
    /// <param name="column">零起始列号。</param>
    /// <param name="row">一基行号。</param>
    /// <returns>A1 格式的单元格地址。</returns>
    private static string CellRef(int column, int row)
    {
        var value = column + 1;
        var letters = string.Empty;
        while (value > 0)
        {
            value--;
            letters = (char)('A' + value % 26) + letters;
            value /= 26;
        }
        return letters + row.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 对 XLSX ZIP 写入流实施序列化字节限制。
    /// </summary>
    private sealed class LimitedWriteStream : Stream
    {
        /// <summary>
        /// 接收受限输出的底层流。
        /// </summary>
        private readonly Stream _inner;

        /// <summary>
        /// 输出允许的最大字节数。
        /// </summary>
        private readonly long? _maximum;

        /// <summary>
        /// 当前累计写入的字节数。
        /// </summary>
        private long _written;

        /// <summary>
        /// 资源诊断使用的 Provider 名称。
        /// </summary>
        private readonly string _providerName;

        /// <summary>
        /// 初始化一个 <see cref="LimitedWriteStream" /> 类型的实例。
        /// </summary>
        /// <param name="inner">底层内存流。</param>
        /// <param name="maximum">可选序列化字节上限。</param>
        /// <param name="providerName">资源诊断使用的 Provider 名称。</param>
        internal LimitedWriteStream(Stream inner, long? maximum, string providerName)
        {
            _inner = inner;
            _maximum = maximum;
            _providerName = Provider(providerName);
        }

        /// <summary>
        /// 获取流是否可读；此输出流不可读。
        /// </summary>
        public override bool CanRead => false;

        /// <summary>
        /// 获取流是否可查找；此输出流不可查找。
        /// </summary>
        public override bool CanSeek => false;

        /// <summary>
        /// 获取流是否可写；取决于底层流。
        /// </summary>
        public override bool CanWrite => _inner.CanWrite;

        /// <summary>
        /// 获取当前输出字节数。
        /// </summary>
        public override long Length => _written;

        /// <summary>
        /// 获取当前输出位置；此流不支持设置位置。
        /// </summary>
        public override long Position
        {
            get => _written;
            set => throw new NotSupportedException();
        }

        /// <summary>
        /// 刷新底层流。
        /// </summary>
        public override void Flush() => _inner.Flush();

        /// <summary>
        /// 异步刷新底层流。
        /// </summary>
        /// <param name="cancellationToken">取消操作的令牌。</param>
        /// <returns>代表刷新操作的任务。</returns>
        public override Task FlushAsync(CancellationToken cancellationToken)
            => _inner.FlushAsync(cancellationToken);

        /// <summary>
        /// 读取操作不受支持。
        /// </summary>
        /// <param name="buffer">接收数据的缓冲区。</param>
        /// <param name="offset">缓冲区起始索引。</param>
        /// <param name="count">请求读取的字节数。</param>
        /// <returns>此方法不会返回。</returns>
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        /// <summary>
        /// 查找操作不受支持。
        /// </summary>
        /// <param name="offset">偏移量。</param>
        /// <param name="origin">偏移基准。</param>
        /// <returns>此方法不会返回。</returns>
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        /// <summary>
        /// 设置长度操作不受支持。
        /// </summary>
        /// <param name="value">请求的长度。</param>
        public override void SetLength(long value) => throw new NotSupportedException();

        /// <summary>
        /// 写入字节并检查累计序列化上限。
        /// </summary>
        /// <param name="buffer">待写入数据。</param>
        /// <param name="offset">缓冲区起始索引。</param>
        /// <param name="count">待写入字节数。</param>
        public override void Write(byte[] buffer, int offset, int count)
        {
            Check(count);
            _inner.Write(buffer, offset, count);
            _written += count;
        }

        /// <summary>
        /// 写入字节跨度并检查累计序列化上限。
        /// </summary>
        /// <param name="buffer">待写入数据。</param>
        public override void Write(ReadOnlySpan<byte> buffer)
        {
            Check(buffer.Length);
            _inner.Write(buffer);
            _written += buffer.Length;
        }

        /// <summary>
        /// 异步写入字节并检查累计序列化上限。
        /// </summary>
        /// <param name="buffer">待写入数据。</param>
        /// <param name="offset">缓冲区起始索引。</param>
        /// <param name="count">待写入字节数。</param>
        /// <param name="cancellationToken">取消操作的令牌。</param>
        /// <returns>代表写入操作的任务。</returns>
        public override async Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            Check(count);
            await _inner.WriteAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
            _written += count;
        }

        /// <summary>
        /// 异步写入内存字节并检查累计序列化上限。
        /// </summary>
        /// <param name="buffer">待写入数据。</param>
        /// <param name="cancellationToken">取消操作的令牌。</param>
        /// <returns>代表写入操作的任务。</returns>
        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            Check(buffer.Length);
            await _inner.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
            _written += buffer.Length;
        }

        /// <summary>
        /// 检查下一次写入是否超过字节上限。
        /// </summary>
        /// <param name="count">下一次写入的字节数。</param>
        private void Check(int count)
        {
            if (_maximum.HasValue && count > _maximum.Value - _written)
                throw ResourceLimit($"失败工作簿序列化结果超过字节限制: {_maximum.Value}",
                    providerName: _providerName);
        }

        /// <summary>
        /// 释放包装器但保留由调用方持有的底层流。
        /// </summary>
        /// <param name="disposing">是否由 Dispose 调用。</param>
        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}
