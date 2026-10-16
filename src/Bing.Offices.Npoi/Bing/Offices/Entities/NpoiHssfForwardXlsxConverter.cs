using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using Bing.Offices.ExcelDataReader;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Entities;
using NPOI.HSSF.EventUserModel;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record;
using NPOI.HSSF.Util;
using NPOI.POIFS.FileSystem;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Entities;

/// <summary>
/// 将受限 HSSF 记录事件转换为可供 Forward 读取器消费的最小 XLSX 投影。
/// </summary>
/// <remarks>
/// 转换只保留选定工作表的单元格值、行号、日期格式、合并区域和受限命名锚点，行内容先落入删除即清理的临时文件，
/// 不创建 HSSFWorkbook DOM。公式、图片、关系和其他复杂对象由 Forward 布局边界拒绝或忽略。
/// </remarks>
internal static class NpoiHssfForwardXlsxConverter
{
    private const string SpreadsheetNamespace =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipsNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationshipsNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string ContentTypesNamespace =
        "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string PackageRelationshipType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string WorksheetRelationshipType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    private const string StylesRelationshipType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

    /// <summary>
    /// 将 HSSF 输入转换为选定工作表的 XLSX 临时流。
    /// </summary>
    /// <param name="source">HSSF 源流。</param>
    /// <param name="sheetNames">需要保留的工作表名称。</param>
    /// <param name="limits">资源限制。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="failureCapture">可选的低保留失败工作簿捕获器。</param>
    /// <param name="requireNativeDataValidation">是否要求安全投影原生校验；不支持时失败。</param>
    /// <param name="temporaryStorage">与源输入和后续解析共享的临时存储预算。</param>
    /// <returns>位于起始位置、由调用方负责释放的 XLSX 临时文件流。</returns>
    internal static Stream Convert(Stream source, IReadOnlyList<string> sheetNames,
        ExcelResourceLimits limits, CancellationToken cancellationToken,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null,
        bool requireNativeDataValidation = false,
        ExcelTemporaryStorageBudget temporaryStorage = null)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("输入流不可读取。", nameof(source));
        if (sheetNames == null || sheetNames.Count == 0
            || sheetNames.Any(string.IsNullOrWhiteSpace))
            throw new ArgumentException("目标工作表名称不能为空。", nameof(sheetNames));
        if (limits == null)
            throw new ArgumentNullException(nameof(limits));
        limits.Validate();
        cancellationToken.ThrowIfCancellationRequested();

        var output = CreateTempBuffer(temporaryStorage);
        try
        {
            using var listener = new Listener(sheetNames, limits, cancellationToken, failureCapture,
                requireNativeDataValidation, temporaryStorage);
            var tracking = new FormatTrackingHSSFListener(listener);
            listener.SetFormatTracker(tracking);
            var request = new HSSFRequest();
            request.AddListenerForAllRecords(tracking);
            var fileSystem = new POIFSFileSystem(
                new Bing.Offices.Exports.NpoiNonDisposingStream(source, cancellationToken));
            try
            {
                new HSSFEventFactory().ProcessWorkbookEvents(request, fileSystem);
            }
            finally
            {
                fileSystem.Close();
            }

            listener.Complete();
            WritePackage(output, listener);
            output.Flush();
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
    /// 写入最小的 XLSX 包结构。
    /// </summary>
    private static void WritePackage(Stream output, Listener listener)
    {
        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);
        WriteXml(archive, "[Content_Types].xml", writer =>
        {
            writer.WriteStartElement("Types", ContentTypesNamespace);
            writer.WriteStartElement("Default", ContentTypesNamespace);
            writer.WriteAttributeString("Extension", "rels");
            writer.WriteAttributeString("ContentType",
                "application/vnd.openxmlformats-package.relationships+xml");
            writer.WriteEndElement();
            writer.WriteStartElement("Default", ContentTypesNamespace);
            writer.WriteAttributeString("Extension", "xml");
            writer.WriteAttributeString("ContentType", "application/xml");
            writer.WriteEndElement();
            WriteOverride(writer, "/xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            for (var index = 0; index < listener.SheetNames.Count; index++)
                WriteOverride(writer, $"/xl/worksheets/sheet{index + 1}.xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            WriteOverride(writer, "/xl/styles.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            writer.WriteEndElement();
        });
        WriteXml(archive, "_rels/.rels", writer =>
        {
            writer.WriteStartElement("Relationships", PackageRelationshipsNamespace);
            writer.WriteStartElement("Relationship", PackageRelationshipsNamespace);
            writer.WriteAttributeString("Id", "rId1");
            writer.WriteAttributeString("Type", PackageRelationshipType);
            writer.WriteAttributeString("Target", "xl/workbook.xml");
            writer.WriteEndElement();
            writer.WriteEndElement();
        });
        WriteXml(archive, "xl/workbook.xml", writer =>
        {
            writer.WriteStartElement("workbook", SpreadsheetNamespace);
            writer.WriteAttributeString("xmlns", "r", "http://www.w3.org/2000/xmlns/",
                RelationshipsNamespace);
            if (listener.Date1904)
            {
                writer.WriteStartElement("workbookPr", SpreadsheetNamespace);
                writer.WriteAttributeString("date1904", "1");
                writer.WriteEndElement();
            }
            writer.WriteStartElement("sheets", SpreadsheetNamespace);
            for (var index = 0; index < listener.SheetNames.Count; index++)
            {
                writer.WriteStartElement("sheet", SpreadsheetNamespace);
                writer.WriteAttributeString("name", listener.SheetNames[index]);
                writer.WriteAttributeString("sheetId",
                    (index + 1).ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("r", "id", RelationshipsNamespace,
                    $"rId{index + 1}");
                var state = listener.SheetStates[index];
                if (!string.IsNullOrWhiteSpace(state))
                    writer.WriteAttributeString("state", state);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            WriteDefinedNames(writer, listener);
            writer.WriteEndElement();
        });
        WriteXml(archive, "xl/_rels/workbook.xml.rels", writer =>
        {
            writer.WriteStartElement("Relationships", PackageRelationshipsNamespace);
            for (var index = 0; index < listener.SheetNames.Count; index++)
                WriteRelationship(writer, $"rId{index + 1}", WorksheetRelationshipType,
                    $"worksheets/sheet{index + 1}.xml");
            WriteRelationship(writer,
                $"rId{listener.SheetNames.Count + 1}", StylesRelationshipType, "styles.xml");
            writer.WriteEndElement();
        });
        WriteXml(archive, "xl/styles.xml", writer => WriteStyles(writer, listener));
        for (var index = 0; index < listener.SheetNames.Count; index++)
        {
            var sheetEntry = archive.CreateEntry($"xl/worksheets/sheet{index + 1}.xml",
                CompressionLevel.Fastest);
            using var destination = sheetEntry.Open();
            listener.CopySheetXmlTo(index, destination);
        }
    }

    /// <summary>
    /// 写入可供 Forward 读取器解析的命名锚点。
    /// </summary>
    private static void WriteDefinedNames(XmlWriter writer, Listener listener)
    {
        var definedNames = listener.DefinedNames;
        if (definedNames.Count == 0)
            return;
        writer.WriteStartElement("definedNames", SpreadsheetNamespace);
        foreach (var definedName in definedNames)
        {
            writer.WriteStartElement("definedName", SpreadsheetNamespace);
            writer.WriteAttributeString("name", definedName.Name);
            if (definedName.LocalSheetId >= 0)
                writer.WriteAttributeString("localSheetId",
                    definedName.LocalSheetId.ToString(CultureInfo.InvariantCulture));
            writer.WriteString(definedName.Formula);
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入最小可读的工作簿样式表。
    /// </summary>
    private static void WriteStyles(XmlWriter writer, Listener listener)
    {
        writer.WriteStartElement("styleSheet", SpreadsheetNamespace);
        if (listener.CustomFormats.Count > 0)
        {
            writer.WriteStartElement("numFmts", SpreadsheetNamespace);
            writer.WriteAttributeString("count",
                listener.CustomFormats.Count.ToString(CultureInfo.InvariantCulture));
            for (var index = 0; index < listener.CustomFormats.Count; index++)
            {
                writer.WriteStartElement("numFmt", SpreadsheetNamespace);
                writer.WriteAttributeString("numFmtId",
                    (164 + index).ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("formatCode", listener.CustomFormats[index]);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }
        writer.WriteStartElement("fonts", SpreadsheetNamespace);
        writer.WriteAttributeString("count",
            listener.FontDefinitions.Count.ToString(CultureInfo.InvariantCulture));
        foreach (var font in listener.FontDefinitions)
            WriteFont(writer, font);
        writer.WriteEndElement();
        writer.WriteStartElement("fills", SpreadsheetNamespace);
        writer.WriteAttributeString("count",
            (2 + listener.FillDefinitions.Count).ToString(CultureInfo.InvariantCulture));
        WriteEmptyElement(writer, "fill", SpreadsheetNamespace, "patternFill", "patternType", "none");
        WriteEmptyElement(writer, "fill", SpreadsheetNamespace, "patternFill", "patternType", "gray125");
        foreach (var fill in listener.FillDefinitions)
            WriteFill(writer, fill);
        writer.WriteEndElement();
        writer.WriteStartElement("borders", SpreadsheetNamespace);
        writer.WriteAttributeString("count",
            (1 + listener.BorderDefinitions.Count).ToString(CultureInfo.InvariantCulture));
        WriteEmptyBorder(writer);
        foreach (var border in listener.BorderDefinitions)
            WriteBorder(writer, border);
        writer.WriteEndElement();
        writer.WriteStartElement("cellStyleXfs", SpreadsheetNamespace);
        writer.WriteAttributeString("count", "1");
        WriteXf(writer, 0, 0, 0, 0, null, null);
        writer.WriteEndElement();
        writer.WriteStartElement("cellXfs", SpreadsheetNamespace);
        writer.WriteAttributeString("count",
            (1 + listener.StyleDefinitions.Count).ToString(CultureInfo.InvariantCulture));
        WriteXf(writer, 0, 0, 0, 0, null, null);
        foreach (var style in listener.StyleDefinitions)
            WriteXf(writer, style.NumberFormatId, style.FontId, style.FillId, style.BorderId,
                style.Alignment, style.Protection);
        writer.WriteEndElement();
        writer.WriteStartElement("cellStyles", SpreadsheetNamespace);
        writer.WriteAttributeString("count", "1");
        writer.WriteStartElement("cellStyle", SpreadsheetNamespace);
        writer.WriteAttributeString("name", "Normal");
        writer.WriteAttributeString("xfId", "0");
        writer.WriteAttributeString("builtinId", "0");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入最小字体描述。
    /// </summary>
    private static void WriteFont(XmlWriter writer, FontDefinition font)
    {
        writer.WriteStartElement("font", SpreadsheetNamespace);
        if (font.IsBold)
            writer.WriteStartElement("b", SpreadsheetNamespace);
        if (font.IsBold)
            writer.WriteEndElement();
        if (font.IsItalic)
            writer.WriteStartElement("i", SpreadsheetNamespace);
        if (font.IsItalic)
            writer.WriteEndElement();
        if (font.IsStrikeout)
            writer.WriteStartElement("strike", SpreadsheetNamespace);
        if (font.IsStrikeout)
            writer.WriteEndElement();
        if (font.IsUnderline)
        {
            writer.WriteStartElement("u", SpreadsheetNamespace);
            writer.WriteEndElement();
        }
        writer.WriteStartElement("sz", SpreadsheetNamespace);
        writer.WriteAttributeString("val", font.FontHeightInPoints > 0
            ? font.FontHeightInPoints.ToString("R", CultureInfo.InvariantCulture) : "11");
        writer.WriteEndElement();
        writer.WriteStartElement("color", SpreadsheetNamespace);
        if (font.FontColorIndex.HasValue)
            writer.WriteAttributeString("indexed",
                font.FontColorIndex.Value.ToString(CultureInfo.InvariantCulture));
        else
            writer.WriteAttributeString("theme", "1");
        writer.WriteEndElement();
        writer.WriteStartElement("name", SpreadsheetNamespace);
        writer.WriteAttributeString("val", string.IsNullOrWhiteSpace(font.FontName)
            ? "Calibri" : font.FontName);
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入受限的纯色填充。
    /// </summary>
    private static void WriteFill(XmlWriter writer, FillDefinition fill)
    {
        writer.WriteStartElement("fill", SpreadsheetNamespace);
        writer.WriteStartElement("patternFill", SpreadsheetNamespace);
        writer.WriteAttributeString("patternType", fill.Pattern);
        writer.WriteStartElement("fgColor", SpreadsheetNamespace);
        writer.WriteAttributeString("indexed",
            fill.Foreground.Value.ToString(CultureInfo.InvariantCulture));
        writer.WriteEndElement();
        if (fill.Background.HasValue)
        {
            writer.WriteStartElement("bgColor", SpreadsheetNamespace);
            writer.WriteAttributeString("indexed",
                fill.Background.Value.ToString(CultureInfo.InvariantCulture));
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入空边框定义。
    /// </summary>
    private static void WriteEmptyBorder(XmlWriter writer)
    {
        writer.WriteStartElement("border", SpreadsheetNamespace);
        foreach (var side in new[] { "left", "right", "top", "bottom", "diagonal" })
            WriteEmptyElement(writer, side, SpreadsheetNamespace, null, null, null);
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入标准四边框定义。
    /// </summary>
    private static void WriteBorder(XmlWriter writer, BorderDefinition border)
    {
        writer.WriteStartElement("border", SpreadsheetNamespace);
        WriteBorderSide(writer, "left", border.Left);
        WriteBorderSide(writer, "right", border.Right);
        WriteBorderSide(writer, "top", border.Top);
        WriteBorderSide(writer, "bottom", border.Bottom);
        WriteEmptyElement(writer, "diagonal", SpreadsheetNamespace, null, null, null);
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入一侧边框。
    /// </summary>
    private static void WriteBorderSide(XmlWriter writer, string name, BorderSideDefinition side)
    {
        writer.WriteStartElement(name, SpreadsheetNamespace);
        if (side != null && !string.IsNullOrWhiteSpace(side.Style))
        {
            writer.WriteAttributeString("style", side.Style);
            if (side.Color.HasValue)
            {
                writer.WriteStartElement("color", SpreadsheetNamespace);
                writer.WriteAttributeString("indexed",
                    side.Color.Value.ToString(CultureInfo.InvariantCulture));
                writer.WriteEndElement();
            }
        }
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入单个单元格样式。
    /// </summary>
    private static void WriteXf(XmlWriter writer, int numberFormatId, int fontId, int fillId,
        int borderId, AlignmentDefinition alignment, ProtectionDefinition protection)
    {
        writer.WriteStartElement("xf", SpreadsheetNamespace);
        writer.WriteAttributeString("numFmtId", numberFormatId.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("fontId", fontId.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("fillId", fillId.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("borderId", borderId.ToString(CultureInfo.InvariantCulture));
        if (alignment != null)
            writer.WriteAttributeString("applyAlignment", "1");
        if (protection != null)
            writer.WriteAttributeString("applyProtection", "1");
        if (alignment != null)
        {
            writer.WriteStartElement("alignment", SpreadsheetNamespace);
            WriteOptionalAttribute(writer, "horizontal", alignment.Horizontal);
            WriteOptionalAttribute(writer, "vertical", alignment.Vertical);
            if (alignment.WrapText)
                writer.WriteAttributeString("wrapText", "1");
            if (alignment.ShrinkToFit)
                writer.WriteAttributeString("shrinkToFit", "1");
            if (alignment.Rotation.HasValue)
                writer.WriteAttributeString("textRotation",
                    alignment.Rotation.Value.ToString(CultureInfo.InvariantCulture));
            if (alignment.Indent > 0)
                writer.WriteAttributeString("indent",
                    alignment.Indent.ToString(CultureInfo.InvariantCulture));
            writer.WriteEndElement();
        }
        if (protection != null)
        {
            writer.WriteStartElement("protection", SpreadsheetNamespace);
            writer.WriteAttributeString("locked", protection.IsLocked ? "1" : "0");
            writer.WriteAttributeString("hidden", protection.IsHidden ? "1" : "0");
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入非空的样式属性。
    /// </summary>
    private static void WriteOptionalAttribute(XmlWriter writer, string name, string value)
    {
        if (!string.IsNullOrWhiteSpace(value))
            writer.WriteAttributeString(name, value);
    }

    /// <summary>
    /// 写入一个关系元素。
    /// </summary>
    private static void WriteRelationship(XmlWriter writer, string id, string type, string target)
    {
        writer.WriteStartElement("Relationship", PackageRelationshipsNamespace);
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Type", type);
        writer.WriteAttributeString("Target", target);
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入内容类型覆盖项。
    /// </summary>
    private static void WriteOverride(XmlWriter writer, string partName, string contentType)
    {
        writer.WriteStartElement("Override", ContentTypesNamespace);
        writer.WriteAttributeString("PartName", partName);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入空的带属性 XML 元素。
    /// </summary>
    private static void WriteEmptyElement(XmlWriter writer, string elementName, string namespaceName,
        string childName, string attributeName, string attributeValue)
    {
        writer.WriteStartElement(elementName, namespaceName);
        if (!string.IsNullOrWhiteSpace(childName))
        {
            writer.WriteStartElement(childName, namespaceName);
            writer.WriteAttributeString(attributeName, attributeValue);
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    /// <summary>
    /// 写入 XML 部件。
    /// </summary>
    private static void WriteXml(ZipArchive archive, string path, Action<XmlWriter> write)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Fastest);
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            OmitXmlDeclaration = false,
            CloseOutput = false
        });
        write(writer);
    }

    /// <summary>
    /// 创建由调用方负责释放并自动删除的转换输出文件。
    /// </summary>
    private static Stream CreateTempBuffer(ExcelTemporaryStorageBudget temporaryStorage)
    {
        var path = Path.Combine(Path.GetTempPath(),
            $"bing-offices-forward-hssf-{Guid.NewGuid():N}.xlsx");
        return (temporaryStorage ?? new ExcelTemporaryStorageBudget(null, "NPOI")).CreateFile(path,
            FileOptions.Asynchronous | FileOptions.SequentialScan | FileOptions.DeleteOnClose);
    }

    /// <summary>
    /// 保存投影样式需要的最小格式描述。
    /// </summary>
    private sealed class StyleDefinition
    {
        internal StyleDefinition(int numberFormatId, int fontId, int fillId, int borderId,
            AlignmentDefinition alignment = null, ProtectionDefinition protection = null)
        {
            NumberFormatId = numberFormatId;
            FontId = fontId;
            FillId = fillId;
            BorderId = borderId;
            Alignment = alignment;
            Protection = protection;
        }

        /// <summary>
        /// 获取数字格式索引。
        /// </summary>
        internal int NumberFormatId { get; }

        /// <summary>
        /// 获取字体索引。
        /// </summary>
        internal int FontId { get; }

        /// <summary>
        /// 获取填充索引。
        /// </summary>
        internal int FillId { get; }

        /// <summary>
        /// 获取边框索引。
        /// </summary>
        internal int BorderId { get; }

        /// <summary>
        /// 获取对齐状态。
        /// </summary>
        internal AlignmentDefinition Alignment { get; }

        /// <summary>
        /// 获取保护状态。
        /// </summary>
        internal ProtectionDefinition Protection { get; }
    }

    /// <summary>
    /// 保存投影所需的最小单元格保护状态。
    /// </summary>
    private sealed class ProtectionDefinition
    {
        /// <summary>
        /// 初始化一个保护状态描述。
        /// </summary>
        internal ProtectionDefinition(bool isLocked, bool isHidden)
        {
            IsLocked = isLocked;
            IsHidden = isHidden;
        }

        /// <summary>
        /// 获取是否锁定单元格。
        /// </summary>
        internal bool IsLocked { get; }

        /// <summary>
        /// 获取是否隐藏单元格公式。
        /// </summary>
        internal bool IsHidden { get; }

        /// <summary>
        /// 获取用于样式缓存的状态键。
        /// </summary>
        internal string Key => (IsLocked ? "1" : "0") + (IsHidden ? "1" : "0");

        /// <summary>
        /// 获取是否等同于 SpreadsheetML 默认保护状态。
        /// </summary>
        internal bool IsEmpty => IsLocked && !IsHidden;
    }

    /// <summary>
    /// 保存投影所需的最小对齐状态。
    /// </summary>
    private sealed class AlignmentDefinition
    {
        /// <summary>
        /// 初始化一个对齐状态描述。
        /// </summary>
        internal AlignmentDefinition(string horizontal, string vertical, bool wrapText,
            bool shrinkToFit, short? rotation, short indent)
        {
            Horizontal = string.IsNullOrWhiteSpace(horizontal) ? null : horizontal;
            Vertical = string.IsNullOrWhiteSpace(vertical) ? null : vertical;
            WrapText = wrapText;
            ShrinkToFit = shrinkToFit;
            Rotation = rotation >= 0 && rotation <= 180 ? rotation : null;
            Indent = indent >= 0 && indent <= 250 ? indent : (short)0;
        }

        /// <summary>
        /// 获取水平对齐名称。
        /// </summary>
        internal string Horizontal { get; }

        /// <summary>
        /// 获取垂直对齐名称。
        /// </summary>
        internal string Vertical { get; }

        /// <summary>
        /// 获取是否自动换行。
        /// </summary>
        internal bool WrapText { get; }

        /// <summary>
        /// 获取是否缩小字体适应单元格。
        /// </summary>
        internal bool ShrinkToFit { get; }

        /// <summary>
        /// 获取文本旋转角度。
        /// </summary>
        internal short? Rotation { get; }

        /// <summary>
        /// 获取文本缩进量。
        /// </summary>
        internal short Indent { get; }

        /// <summary>
        /// 获取是否没有需要写出的对齐状态。
        /// </summary>
        internal bool IsEmpty => string.IsNullOrWhiteSpace(Horizontal)
            && string.IsNullOrWhiteSpace(Vertical) && !WrapText && !ShrinkToFit
            && !Rotation.HasValue && Indent == 0;

        /// <summary>
        /// 获取用于样式缓存的状态键。
        /// </summary>
        internal string Key => (Horizontal ?? string.Empty) + "\0"
            + (Vertical ?? string.Empty) + "\0" + (WrapText ? "1" : "0")
            + (ShrinkToFit ? "1" : "0") + "\0"
            + (Rotation?.ToString(CultureInfo.InvariantCulture) ?? string.Empty)
            + "\0" + Indent.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 保存投影所需的最小填充状态。
    /// </summary>
    private sealed class FillDefinition
    {
        /// <summary>
        /// 初始化一个填充状态描述。
        /// </summary>
        internal FillDefinition(string pattern, short? foreground, short? background)
        {
            Pattern = string.IsNullOrWhiteSpace(pattern) ? null : pattern.Trim();
            Foreground = foreground >= 0 ? foreground : null;
            Background = background >= 0 ? background : null;
        }

        /// <summary>
        /// 获取填充模式。
        /// </summary>
        internal string Pattern { get; }

        /// <summary>
        /// 获取前景索引色。
        /// </summary>
        internal short? Foreground { get; }

        /// <summary>
        /// 获取背景索引色。
        /// </summary>
        internal short? Background { get; }

        /// <summary>
        /// 获取是否为当前支持的纯色前景填充。
        /// </summary>
        internal bool IsEmpty => !string.Equals(Pattern, "solid", StringComparison.Ordinal)
            || !Foreground.HasValue;

        /// <summary>
        /// 获取用于样式缓存的状态键。
        /// </summary>
        internal string Key => (Pattern ?? string.Empty) + "\0"
            + (Foreground?.ToString(CultureInfo.InvariantCulture) ?? string.Empty) + "\0"
            + (Background?.ToString(CultureInfo.InvariantCulture) ?? string.Empty);
    }

    /// <summary>
    /// 保存投影所需的标准四边框状态。
    /// </summary>
    private sealed class BorderDefinition
    {
        /// <summary>
        /// 初始化一个边框状态描述。
        /// </summary>
        internal BorderDefinition(BorderSideDefinition left, BorderSideDefinition right,
            BorderSideDefinition top, BorderSideDefinition bottom)
        {
            Left = left;
            Right = right;
            Top = top;
            Bottom = bottom;
        }

        /// <summary>
        /// 获取左边框状态。
        /// </summary>
        internal BorderSideDefinition Left { get; }

        /// <summary>
        /// 获取右边框状态。
        /// </summary>
        internal BorderSideDefinition Right { get; }

        /// <summary>
        /// 获取上边框状态。
        /// </summary>
        internal BorderSideDefinition Top { get; }

        /// <summary>
        /// 获取下边框状态。
        /// </summary>
        internal BorderSideDefinition Bottom { get; }

        /// <summary>
        /// 获取是否没有需要写出的边框。
        /// </summary>
        internal bool IsEmpty => Left == null && Right == null && Top == null && Bottom == null;

        /// <summary>
        /// 获取用于样式缓存的状态键。
        /// </summary>
        internal string Key => (Left?.Key ?? string.Empty) + "\0"
            + (Right?.Key ?? string.Empty) + "\0"
            + (Top?.Key ?? string.Empty) + "\0"
            + (Bottom?.Key ?? string.Empty);
    }

    /// <summary>
    /// 保存一侧边框的线型和索引色。
    /// </summary>
    private sealed class BorderSideDefinition
    {
        /// <summary>
        /// 初始化一侧边框状态。
        /// </summary>
        internal BorderSideDefinition(string style, short? color)
        {
            Style = string.IsNullOrWhiteSpace(style) ? null : style.Trim();
            Color = color >= 0 ? color : null;
        }

        /// <summary>
        /// 获取线型名称。
        /// </summary>
        internal string Style { get; }

        /// <summary>
        /// 获取索引色。
        /// </summary>
        internal short? Color { get; }

        /// <summary>
        /// 获取用于边框缓存的状态键。
        /// </summary>
        internal string Key => (Style ?? string.Empty) + "|"
            + (Color?.ToString(CultureInfo.InvariantCulture) ?? string.Empty);
    }

    /// <summary>
    /// 保存投影所需的最小字体状态。
    /// </summary>
    private sealed class FontDefinition
    {
        internal FontDefinition(bool isBold, bool isItalic, bool isStrikeout,
            bool isUnderline, string fontName = null, double fontHeightInPoints = 0,
            short? fontColorIndex = null)
        {
            IsBold = isBold;
            IsItalic = isItalic;
            IsStrikeout = isStrikeout;
            IsUnderline = isUnderline;
            FontName = string.IsNullOrWhiteSpace(fontName) ? null : fontName.Trim();
            FontHeightInPoints = double.IsNaN(fontHeightInPoints)
                || double.IsInfinity(fontHeightInPoints) || fontHeightInPoints <= 0
                ? 0 : fontHeightInPoints;
            FontColorIndex = fontColorIndex >= 0 ? fontColorIndex : null;
        }

        /// <summary>
        /// 获取是否粗体。
        /// </summary>
        internal bool IsBold { get; }

        /// <summary>
        /// 获取是否斜体。
        /// </summary>
        internal bool IsItalic { get; }

        /// <summary>
        /// 获取是否删除线。
        /// </summary>
        internal bool IsStrikeout { get; }

        /// <summary>
        /// 获取是否单下划线。
        /// </summary>
        internal bool IsUnderline { get; }

        /// <summary>
        /// 获取字体名称。
        /// </summary>
        internal string FontName { get; }

        /// <summary>
        /// 获取字号。
        /// </summary>
        internal double FontHeightInPoints { get; }

        /// <summary>
        /// 获取字体索引色。
        /// </summary>
        internal short? FontColorIndex { get; }

        /// <summary>
        /// 获取用于样式缓存的状态键。
        /// </summary>
        internal string Key => (IsBold ? "1" : "0") + (IsItalic ? "1" : "0")
            + (IsStrikeout ? "1" : "0") + (IsUnderline ? "1" : "0")
            + "\0" + (FontName ?? string.Empty) + "\0"
            + FontHeightInPoints.ToString("R", CultureInfo.InvariantCulture) + "\0"
            + (FontColorIndex?.ToString(CultureInfo.InvariantCulture) ?? string.Empty);
    }

    /// <summary>
    /// 收集 HSSF 事件并写入目标工作表 XML。
    /// </summary>
    private sealed class Listener : IHSSFListener, IDisposable
    {
        private readonly HashSet<string> _requestedSheetNames;
        private readonly ExcelResourceLimits _limits;
        private readonly CancellationToken _cancellationToken;
        private readonly ExcelDataReaderFailureWorkbookCapture _failureCapture;
        /// <summary>
        /// 是否要求保留可执行的原生数据校验。
        /// </summary>
        private readonly bool _requireNativeDataValidation;
        private readonly List<string> _sheetNames = new();
        /// <summary>
        /// HSSF 工作表的隐藏状态，按工作表记录顺序保存。
        /// </summary>
        private readonly List<string> _sheetStates = new();
        private readonly List<string> _selectedSheetNames = new();
        /// <summary>
        /// 投影工作表的隐藏状态，按选定工作表顺序保存。
        /// </summary>
        private readonly List<string> _selectedSheetStates = new();
        private readonly List<SheetXmlBuffer> _sheets = new();
        private readonly List<string> _customFormats = new();
        private readonly List<StyleDefinition> _styleDefinitions = new();
        /// <summary>
        /// HSSF 源字体按记录顺序保存的状态。
        /// </summary>
        private readonly List<FontDefinition> _sourceFonts = new();
        private readonly List<(int FormatIndex, int FontIndex, FillDefinition Fill,
            BorderDefinition Border, AlignmentDefinition Alignment,
            ProtectionDefinition Protection)> _xfStyles = new();
        private readonly List<NameRecord> _nameRecords = new();
        /// <summary>
        /// 命名列表名称按 NameRecord 索引缓存的解析结果。
        /// </summary>
        private readonly Dictionary<int, NamedListResolution> _namedListResolutionCache = new();
        /// <summary>
        /// 当前工作表被安全命名列表引用的区域。
        /// </summary>
        private IReadOnlyList<NamedListRange> _namedReferenceRanges =
            Array.Empty<NamedListRange>();
        /// <summary>
        /// 当前工作表命名列表区域中的单元格快照。
        /// </summary>
        private Dictionary<int, Dictionary<int, object>> _namedReferenceCells;
        /// <summary>
        /// 按工作表暂存命名列表来源单元格，等待工作簿校验规则全部解析后筛选。
        /// </summary>
        private readonly Dictionary<string, Dictionary<int, Dictionary<int, object>>>
            _namedReferenceCellSnapshots = new(StringComparer.OrdinalIgnoreCase);
        /// <summary>
        /// 工作簿中原生数据校验实际引用的命名列表。
        /// </summary>
        private readonly List<(string Name, string LocalSheetName)> _referencedNamedLists = new();
        private readonly List<ExternSheetRecord> _externSheetRecords = new();
        private readonly Dictionary<string, int> _formatStyles =
            new(StringComparer.Ordinal);
        /// <summary>
        /// 投影字体状态到字体索引的映射。
        /// </summary>
        private readonly Dictionary<string, int> _fontStyleIds =
            new(StringComparer.Ordinal);
        /// <summary>
        /// 投影填充状态到填充索引的映射。
        /// </summary>
        private readonly Dictionary<string, int> _fillStyleIds =
            new(StringComparer.Ordinal);
        /// <summary>
        /// 投影边框状态到边框索引的映射。
        /// </summary>
        private readonly Dictionary<string, int> _borderStyleIds =
            new(StringComparer.Ordinal);
        /// <summary>
        /// 投影工作簿实际写出的字体状态。
        /// </summary>
        private readonly List<FontDefinition> _fontDefinitions = new()
        {
            new FontDefinition(false, false, false, false)
        };
        /// <summary>
        /// 投影工作簿实际写出的填充状态。
        /// </summary>
        private readonly List<FillDefinition> _fillDefinitions = new();
        /// <summary>
        /// 投影工作簿实际写出的边框状态。
        /// </summary>
        private readonly List<BorderDefinition> _borderDefinitions = new();
        private FormatTrackingHSSFListener _formatTracker;
        private SSTRecord _sharedStrings;
        private SheetXmlBuffer _sheet;
        private int _worksheetIndex = -1;
        private bool _inTarget;
        private long _cellCount;
        private int _maxRow = -1;
        private int _maxColumn = -1;
        /// <summary>
        /// 当前工作表的列宽和隐藏状态。
        /// </summary>
        private Dictionary<int, (double Width, bool HasWidth, bool Hidden)> _columnMetadata;

        /// <summary>
        /// 当前工作表的行高和隐藏状态。
        /// </summary>
        private Dictionary<int, (short Height, bool HasHeight, bool Hidden)> _rowMetadata;

        /// <summary>
        /// 当前工作表的合并区域。
        /// </summary>
        private List<(int FirstRow, int LastRow, int FirstColumn, int LastColumn)> _mergeRegions;

        /// <summary>
        /// 当前工作表是否启用冻结窗格。
        /// </summary>
        private bool _freezePanes;

        /// <summary>
        /// 当前工作表冻结窗格的分割和滚动坐标。
        /// </summary>
        private (int ColumnSplit, int RowSplit, int LeftColumn, int TopRow)? _freezePane;

        /// <summary>
        /// 当前工作表是否包含原生数据校验记录。
        /// </summary>
        private bool _nativeDataValidationDetected;

        /// <summary>
        /// 当前工作表是否包含无法安全投影的原生数据校验记录。
        /// </summary>
        private bool _nativeDataValidationUnsupported;

        /// <summary>
        /// 当前工作表可安全投影的原生数据校验描述。
        /// </summary>
        private readonly List<ExcelDataReaderFailureNativeDataValidation>
            _nativeDataValidations = new List<ExcelDataReaderFailureNativeDataValidation>();

        /// <summary>
        /// 当前工作表是否包含条件格式记录。
        /// </summary>
        private bool _conditionalFormattingDetected;

        /// <summary>
        /// 当前工作表是否包含无法安全重建的条件格式记录。
        /// </summary>
        private bool _conditionalFormattingUnsupported;

        /// <summary>
        /// 当前条件格式区域头。
        /// </summary>
        private CFHeaderRecord _conditionalFormattingHeader;

        /// <summary>
        /// 当前条件格式区域的规则记录。
        /// </summary>
        private readonly List<CFRuleRecord> _conditionalFormattingRules = new();

        /// <summary>
        /// 当前工作表能够安全重建的条件格式描述。
        /// </summary>
        private readonly List<ExcelDataReaderFailureConditionalFormatting>
            _conditionalFormattings = new();

        private readonly ExcelTemporaryStorageBudget _temporaryStorage;

        internal Listener(IReadOnlyList<string> sheetNames, ExcelResourceLimits limits,
            CancellationToken cancellationToken,
            ExcelDataReaderFailureWorkbookCapture failureCapture,
            bool requireNativeDataValidation, ExcelTemporaryStorageBudget temporaryStorage)
        {
            _requestedSheetNames = new HashSet<string>(sheetNames,
                StringComparer.OrdinalIgnoreCase);
            _limits = limits;
            _cancellationToken = cancellationToken;
            _failureCapture = failureCapture;
            _requireNativeDataValidation = requireNativeDataValidation;
            _temporaryStorage = temporaryStorage;
        }

        internal IReadOnlyList<string> SheetNames => _selectedSheetNames;

        /// <summary>
        /// 获取投影工作簿中工作表的隐藏状态。
        /// </summary>
        internal IReadOnlyList<string> SheetStates => _selectedSheetStates;

        internal bool Date1904 { get; private set; }

        internal IReadOnlyList<string> CustomFormats => _customFormats;

        internal IReadOnlyList<StyleDefinition> StyleDefinitions => _styleDefinitions;

        internal IReadOnlyList<FontDefinition> FontDefinitions => _fontDefinitions;

        /// <summary>
        /// 获取投影工作簿需要写出的填充状态。
        /// </summary>
        internal IReadOnlyList<FillDefinition> FillDefinitions => _fillDefinitions;

        /// <summary>
        /// 获取投影工作簿需要写出的边框状态。
        /// </summary>
        internal IReadOnlyList<BorderDefinition> BorderDefinitions => _borderDefinitions;

        internal IReadOnlyList<DefinedNameProjection> DefinedNames =>
            _nameRecords.Select(CreateDefinedNameProjection)
                .Where(item => item != null)
                .ToArray();

        internal void SetFormatTracker(FormatTrackingHSSFListener formatTracker) =>
            _formatTracker = formatTracker ?? throw new ArgumentNullException(nameof(formatTracker));

        /// <inheritdoc />
        public void ProcessRecord(Record record)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            switch (record)
            {
                case FontRecord font:
                    _sourceFonts.Add(new FontDefinition(font.BoldWeight >= 700,
                        font.IsItalic, font.IsStrikeout,
                        System.Convert.ToInt32(font.Underline)
                            == (int)FontUnderlineType.Single,
                        font.FontName, font.FontHeight / 20d,
                        font.ColorPaletteIndex == HSSFColor.Automatic.Index
                            ? (short?)null : font.ColorPaletteIndex));
                    return;
                case ExtendedFormatRecord extendedFormat:
                    _xfStyles.Add((extendedFormat.FormatIndex, extendedFormat.FontIndex,
                        ToFillDefinition(extendedFormat), ToBorderDefinition(extendedFormat),
                        ToAlignmentDefinition(extendedFormat),
                        ToProtectionDefinition(extendedFormat)));
                    return;
                case BoundSheetRecord boundSheet:
                    _sheetNames.Add(boundSheet.Sheetname);
                    _sheetStates.Add(boundSheet.IsVeryHidden
                        ? "veryHidden" : boundSheet.IsHidden ? "hidden" : null);
                    CheckSheetLimit();
                    return;
                case SSTRecord sharedStrings:
                    _sharedStrings = sharedStrings;
                    return;
                case DateWindow1904Record dateWindow:
                    Date1904 = dateWindow.Windowing != 0;
                    return;
                case ExternSheetRecord externSheet:
                    _externSheetRecords.Add(externSheet);
                    return;
                case NameRecord name:
                    _nameRecords.Add(name);
                    return;
                case BOFRecord bof when bof.Type == BOFRecordType.Worksheet:
                    BeginWorksheet();
                    return;
                case EOFRecord:
                    EndWorksheet();
                    return;
            }

            if (!_inTarget)
                return;
            if (record is CFHeaderRecord conditionalHeader)
            {
                FinalizeConditionalFormatting();
                _conditionalFormattingDetected = true;
                _conditionalFormattingHeader = conditionalHeader;
                _conditionalFormattingRules.Clear();
                return;
            }
            if (record is CFRuleRecord conditionalRule)
            {
                _conditionalFormattingDetected = true;
                if (_conditionalFormattingHeader == null)
                    _conditionalFormattingUnsupported = true;
                else
                    _conditionalFormattingRules.Add(conditionalRule);
                return;
            }
            if (IsConditionalFormattingRecord(record))
            {
                FinalizeConditionalFormatting();
                _conditionalFormattingDetected = true;
                _conditionalFormattingUnsupported = true;
                throw Unsupported("复杂条件格式");
            }
            FinalizeConditionalFormatting();
            if (record is WindowTwoRecord windowTwo)
            {
                _freezePanes = windowTwo.FreezePanes;
                return;
            }
            if (record is PaneRecord pane)
            {
                _freezePane = (pane.X, pane.Y, pane.LeftColumn, pane.TopRow);
                return;
            }
            if (record is ColumnInfoRecord columnInfo)
            {
                CaptureColumnInfo(columnInfo);
                return;
            }
            if (record is RowRecord rowRecord)
            {
                CaptureRowInfo(rowRecord);
                return;
            }
            if (record is MergeCellsRecord mergeCells)
            {
                CaptureMergeRegions(mergeCells);
                return;
            }
            if (record is DVRecord)
            {
                _nativeDataValidationDetected = true;
                var validation = (DVRecord)record;
                if (!TryReadNativeDataValidations(validation, out var validations))
                {
                    _nativeDataValidationUnsupported = true;
                    _nativeDataValidations.Clear();
                }
                else if (!_nativeDataValidationUnsupported)
                {
                    _nativeDataValidations.AddRange(validations);
                }
                return;
            }
            if (record is FormulaRecord)
                throw Unsupported("公式单元格");
            switch (record)
            {
                case LabelSSTRecord labelSst:
                    WriteText(labelSst.Row, labelSst.Column, ResolveString(labelSst.SSTIndex),
                        labelSst);
                    break;
                case LabelRecord label:
                    WriteText(label.Row, label.Column, label.Value, label);
                    break;
                case NumberRecord number:
                    WriteNumber(number.Row, number.Column, number.Value, number);
                    break;
                case RKRecord rk:
                    WriteNumber(rk.Row, rk.Column, rk.RKNumber, rk);
                    break;
                case MulRKRecord multiple:
                    for (var offset = 0; offset < multiple.NumColumns; offset++)
                    {
                        var column = multiple.FirstColumn + offset;
                        WriteNumber(multiple.Row, column, multiple.GetRKNumberAt(offset),
                            new CellFormatInfo(multiple.Row, column, multiple.GetXFAt(offset)));
                    }
                    break;
                case BoolErrRecord booleanOrError when booleanOrError.IsBoolean:
                    WriteBoolean(booleanOrError.Row, booleanOrError.Column,
                        booleanOrError.BooleanValue, booleanOrError);
                    break;
                case BoolErrRecord booleanOrError when booleanOrError.IsError:
                    WriteError(booleanOrError.Row, booleanOrError.Column,
                        booleanOrError.ErrorValue.ToString(CultureInfo.InvariantCulture),
                        booleanOrError);
                    break;
            }
        }

        /// <summary>
        /// 完成事件收集并校验目标工作表存在。
        /// </summary>
        internal void Complete()
        {
            _cancellationToken.ThrowIfCancellationRequested();
            EndWorksheet();
            var missing = _requestedSheetNames.FirstOrDefault(name =>
                !_selectedSheetNames.Contains(name, StringComparer.OrdinalIgnoreCase));
            if (missing != null)
                throw new BingOfficesConfigurationException(
                    $"HSSF Forward 实体布局要求 Sheet {missing}，工作簿中不存在该工作表。",
                    stage: BingOfficesStage.Plan);
            CaptureNamedReferenceRows();
        }

        /// <summary>
        /// 将目标工作表 XML 复制到 ZIP 部件。
        /// </summary>
        internal void CopySheetXmlTo(int index, Stream destination)
        {
            if (index < 0 || index >= _sheets.Count)
                throw new ArgumentOutOfRangeException(nameof(index));
            if (_sheets[index] == null)
                throw new InvalidOperationException("目标工作表尚未完成转换。");
            _sheets[index].CopyTo(destination);
        }

        /// <inheritdoc />
        public void Dispose()
        {
            Exception firstError = null;
            foreach (var sheet in _sheets)
            {
                try { sheet.Dispose(); }
                catch (Exception exception) { firstError ??= exception; }
            }
            if (firstError != null)
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(firstError).Throw();
        }

        /// <summary>
        /// 开始处理一个 HSSF 工作表。
        /// </summary>
        private void BeginWorksheet()
        {
            EndWorksheet();
            _worksheetIndex++;
            if (_worksheetIndex == 0)
                IncludeNamedReferenceSheets();
            var sheetName = _worksheetIndex < _sheetNames.Count
                ? _sheetNames[_worksheetIndex] : null;
            _inTarget = sheetName != null && _requestedSheetNames.Contains(sheetName);
            if (_inTarget)
            {
                _selectedSheetNames.Add(sheetName);
                _selectedSheetStates.Add(_worksheetIndex < _sheetStates.Count
                    ? _sheetStates[_worksheetIndex] : null);
                _namedReferenceRanges = GetNamedReferenceRanges(sheetName);
                _namedReferenceCells = _failureCapture != null
                    && _namedReferenceRanges.Count > 0
                    ? new Dictionary<int, Dictionary<int, object>>() : null;
                _sheet = new SheetXmlBuffer(_temporaryStorage);
                _columnMetadata = new Dictionary<int, (double Width, bool HasWidth, bool Hidden)>();
                _rowMetadata = new Dictionary<int, (short Height, bool HasHeight, bool Hidden)>();
                _mergeRegions = new List<(int FirstRow, int LastRow, int FirstColumn, int LastColumn)>();
                _freezePanes = false;
                _freezePane = null;
                _nativeDataValidations.Clear();
                _nativeDataValidationDetected = false;
                _nativeDataValidationUnsupported = false;
                _conditionalFormattingDetected = false;
                _conditionalFormattingUnsupported = false;
                _conditionalFormattingHeader = null;
                _conditionalFormattingRules.Clear();
                _conditionalFormattings.Clear();
                _sheets.Add(_sheet);
            }
        }

        /// <summary>
        /// 预选安全命名区域的来源工作表，确保跨 Sheet 校验在投影中可解析。
        /// </summary>
        private void IncludeNamedReferenceSheets()
        {
            for (var index = 0; index < _nameRecords.Count; index++)
            {
                if (!TryGetNamedListResolution(index, out var resolution))
                    continue;
                _requestedSheetNames.Add(resolution.TargetSheetName);
            }
        }

        /// <summary>
        /// 获取或创建命名列表名称解析结果。
        /// </summary>
        /// <param name="nameIndex">HSSF 名称记录索引。</param>
        /// <param name="resolution">名称解析结果。</param>
        /// <returns>名称可以安全投影时返回 <see langword="true" />。</returns>
        private bool TryGetNamedListResolution(int nameIndex,
            out NamedListResolution resolution)
        {
            if (_namedListResolutionCache.TryGetValue(nameIndex, out resolution))
                return resolution.IsSupported;

            resolution = CreateNamedListResolution(nameIndex);
            _namedListResolutionCache[nameIndex] = resolution;
            return resolution.IsSupported;
        }

        /// <summary>
        /// 创建单个名称记录的安全解析结果。
        /// </summary>
        private NamedListResolution CreateNamedListResolution(int nameIndex)
        {
            if (nameIndex < 0 || nameIndex >= _nameRecords.Count)
                return NamedListResolution.Unsupported;
            var name = _nameRecords[nameIndex];
            if (name == null || name.IsHiddenName || name.IsBuiltInName
                || name.IsFunctionName || name.IsCommandName
                || string.IsNullOrWhiteSpace(name.NameText)
                || name.NameDefinition == null || name.NameDefinition.Length != 1)
                return NamedListResolution.Unsupported;

            var localSheetName = ResolveNameScopeSheet(name);
            if (name.SheetNumber > 0 && string.IsNullOrWhiteSpace(localSheetName))
                return NamedListResolution.Unsupported;
            if (!TryResolveNamedListTarget(name, localSheetName, out var targetSheetName,
                    out var range) || string.IsNullOrWhiteSpace(targetSheetName))
                return NamedListResolution.Unsupported;
            return new NamedListResolution(name.NameText.Trim(), localSheetName,
                targetSheetName, range);
        }

        /// <summary>
        /// 获取当前工作表被安全命名列表引用的去重区域。
        /// </summary>
        /// <param name="sheetName">当前工作表名称。</param>
        /// <returns>命名列表引用的区域集合。</returns>
        private IReadOnlyList<NamedListRange> GetNamedReferenceRanges(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                return Array.Empty<NamedListRange>();
            return _namedListResolutionCache.Values
                .Where(item => item != null && item.IsSupported
                    && string.Equals(item.TargetSheetName, sheetName,
                        StringComparison.OrdinalIgnoreCase)
                    && item.Range != null)
                .Select(item => item.Range)
                .GroupBy(item => (item.FirstRow, item.LastRow,
                    item.FirstColumn, item.LastColumn))
                .Select(group => group.First())
                .ToArray();
        }

        /// <summary>
        /// 解析名称记录的局部作用域工作表。
        /// </summary>
        /// <param name="name">HSSF 名称记录。</param>
        /// <returns>局部作用域工作表名称；全局名称返回空值。</returns>
        private string ResolveNameScopeSheet(NameRecord name)
        {
            if (name == null || name.SheetNumber <= 0)
                return null;
            var sheetIndex = name.SheetNumber - 1;
            return sheetIndex >= 0 && sheetIndex < _sheetNames.Count
                ? _sheetNames[sheetIndex] : null;
        }

        /// <summary>
        /// 结束当前 HSSF 工作表。
        /// </summary>
        private void EndWorksheet()
        {
            if (!_inTarget)
                return;
            FinalizeConditionalFormatting();
            var sheetName = _selectedSheetNames[_selectedSheetNames.Count - 1];
            StoreNamedReferenceRows(sheetName);
            _failureCapture?.CaptureSheetVisibility(sheetName,
                _selectedSheetStates[_selectedSheetStates.Count - 1]);
            _failureCapture?.CaptureSheetColumnMetadata(
                sheetName, _columnMetadata);
            _failureCapture?.CaptureSheetRowMetadata(
                sheetName, _rowMetadata);
            _failureCapture?.CaptureSheetMergeRegions(
                sheetName, _mergeRegions);
            if (_freezePanes && _freezePane.HasValue)
            {
                var pane = _freezePane.Value;
                _failureCapture?.CaptureSheetFreezePane(
                    sheetName, pane.ColumnSplit, pane.RowSplit, pane.LeftColumn, pane.TopRow);
            }
            if (_nativeDataValidationDetected)
            {
                _failureCapture?.CaptureSheetNativeDataValidations(
                    sheetName,
                    _nativeDataValidationUnsupported
                        ? Array.Empty<ExcelDataReaderFailureNativeDataValidation>()
                        : _nativeDataValidations);
                if (_nativeDataValidationUnsupported && _requireNativeDataValidation)
                    throw Unsupported("无法安全投影的原生数据校验");
            }
            _sheet?.Complete(_nativeDataValidationUnsupported
                ? Array.Empty<ExcelDataReaderFailureNativeDataValidation>()
                : _nativeDataValidations, _mergeRegions);
            if (_conditionalFormattingDetected)
            {
                _failureCapture?.CaptureSheetConditionalFormattings(
                    sheetName, _conditionalFormattingUnsupported, _conditionalFormattings);
            }
            _inTarget = false;
            _namedReferenceRanges = Array.Empty<NamedListRange>();
            _namedReferenceCells = null;
            _columnMetadata = null;
            _rowMetadata = null;
            _mergeRegions = null;
            _freezePanes = false;
            _freezePane = null;
            _nativeDataValidationDetected = false;
            _nativeDataValidationUnsupported = false;
            _nativeDataValidations.Clear();
            _conditionalFormattingDetected = false;
            _conditionalFormattingUnsupported = false;
            _conditionalFormattingHeader = null;
            _conditionalFormattingRules.Clear();
            _conditionalFormattings.Clear();
        }

        /// <summary>
        /// 完成当前条件格式区域的低保留快照。
        /// </summary>
        private void FinalizeConditionalFormatting()
        {
            if (_conditionalFormattingHeader == null)
                return;
            if (!_conditionalFormattingUnsupported
                && TryCreateConditionalFormatting(_conditionalFormattingHeader,
                    _conditionalFormattingRules, out var formatting))
                _conditionalFormattings.Add(formatting);
            else
            {
                _conditionalFormattingUnsupported = true;
                throw Unsupported("复杂条件格式");
            }
            _conditionalFormattingHeader = null;
            _conditionalFormattingRules.Clear();
        }

        /// <summary>
        /// 将 HSSF 条件格式区域转换为低保留描述。
        /// </summary>
        /// <param name="header">条件格式区域头。</param>
        /// <param name="rules">区域规则。</param>
        /// <param name="formatting">转换后的条件格式描述。</param>
        /// <returns>区域和规则均在受支持边界内时返回 <see langword="true" />。</returns>
        private static bool TryCreateConditionalFormatting(CFHeaderRecord header,
            IReadOnlyList<CFRuleRecord> rules,
            out ExcelDataReaderFailureConditionalFormatting formatting)
        {
            formatting = null;
            if (header?.CellRanges == null || header.CellRanges.Length == 0
                || rules == null || rules.Count == 0
                || header.NumberOfConditionalFormats != rules.Count)
                return false;
            var ranges = new List<ExcelDataReaderFailureConditionalFormattingRange>();
            foreach (var range in header.CellRanges)
            {
                if (range == null || range.FirstRow < 0 || range.LastRow < range.FirstRow
                    || range.LastRow > 65_535 || range.FirstColumn < 0
                    || range.LastColumn < range.FirstColumn || range.LastColumn > 255)
                    return false;
                ranges.Add(new ExcelDataReaderFailureConditionalFormattingRange(
                    range.FirstRow, range.LastRow, range.FirstColumn, range.LastColumn));
            }
            var convertedRules = new List<ExcelDataReaderFailureConditionalFormattingRule>();
            foreach (var rule in rules)
            {
                if (!TryCreateConditionalFormattingRule(rule, out var convertedRule))
                    return false;
                convertedRules.Add(convertedRule);
            }
            formatting = new ExcelDataReaderFailureConditionalFormatting(ranges, convertedRules);
            return true;
        }

        /// <summary>
        /// 将 HSSF 条件格式规则转换为低保留描述。
        /// </summary>
        /// <param name="rule">HSSF 条件格式规则。</param>
        /// <param name="convertedRule">转换后的规则。</param>
        /// <returns>规则只使用受支持的条件、公式和样式时返回 <see langword="true" />。</returns>
        private static bool TryCreateConditionalFormattingRule(CFRuleRecord rule,
            out ExcelDataReaderFailureConditionalFormattingRule convertedRule)
        {
            convertedRule = null;
            if (rule == null || (rule.ConditionType != 1 && rule.ConditionType != 2)
                || !TryFormatConditionalFormula(rule.ParsedExpression1, out var formula1))
                return false;
            string formula2 = null;
            if (rule.ParsedExpression2 != null && rule.ParsedExpression2.Length > 0
                && !TryFormatConditionalFormula(rule.ParsedExpression2, out formula2))
                return false;
            if (rule.ConditionType == 1
                && (rule.ComparisonOperation < 1 || rule.ComparisonOperation > 8
                    || ((rule.ComparisonOperation == 1 || rule.ComparisonOperation == 2)
                        && string.IsNullOrWhiteSpace(formula2))))
                return false;
            if (!TryCreateConditionalFormattingStyle(rule, out var style))
                return false;
            convertedRule = new ExcelDataReaderFailureConditionalFormattingRule(
                rule.ConditionType, rule.ComparisonOperation, formula1, formula2, style);
            return true;
        }

        /// <summary>
        /// 将条件格式公式令牌转换为安全的公式文本。
        /// </summary>
        /// <param name="expression">HSSF 公式令牌。</param>
        /// <param name="formula">转换后的公式。</param>
        /// <returns>公式不含外部、名称或动态表达式时返回 <see langword="true" />。</returns>
        private static bool TryFormatConditionalFormula(Ptg[] expression, out string formula)
        {
            formula = null;
            if (expression == null || expression.Length == 0)
                return false;
            foreach (var token in expression)
            {
                var tokenName = token?.GetType().Name ?? string.Empty;
                if (tokenName.IndexOf("Name", StringComparison.OrdinalIgnoreCase) >= 0
                    || tokenName.IndexOf("3D", StringComparison.OrdinalIgnoreCase) >= 0
                    || tokenName.IndexOf("External", StringComparison.OrdinalIgnoreCase) >= 0
                    || tokenName.IndexOf("Function", StringComparison.OrdinalIgnoreCase) >= 0
                    || tokenName.IndexOf("Array", StringComparison.OrdinalIgnoreCase) >= 0
                    || tokenName.IndexOf("Mem", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(tokenName, "AttrPtg", StringComparison.Ordinal)
                    || tokenName.IndexOf("Deleted", StringComparison.OrdinalIgnoreCase) >= 0
                    || tokenName.IndexOf("Err", StringComparison.OrdinalIgnoreCase) >= 0)
                    return false;
            }
            try
            {
                formula = HSSFFormulaParser.ToFormulaString(null, expression)?.Trim();
            }
            catch
            {
                return false;
            }
            return !string.IsNullOrWhiteSpace(formula) && formula.Length <= 255
                && formula.IndexOf('!') < 0 && formula.IndexOf('[') < 0
                && formula.IndexOf(']') < 0 && formula.IndexOf("#REF!",
                    StringComparison.OrdinalIgnoreCase) < 0;
        }

        /// <summary>
        /// 转换条件格式使用的字体、填充和边框状态。
        /// </summary>
        /// <param name="rule">HSSF 条件格式规则。</param>
        /// <param name="style">转换后的样式。</param>
        /// <returns>样式只使用常规状态时返回 <see langword="true" />。</returns>
        private static bool TryCreateConditionalFormattingStyle(CFRuleRecord rule,
            out ExcelDataReaderFailureFontStyle style)
        {
            style = null;
            if (rule.ContainsFontFormattingBlock)
            {
                var font = rule.FontFormatting;
                if (font == null || font.IsOutlineOn || font.IsShadowOn
                    || font.EscapementType != FontSuperScript.None)
                    return false;
            }
            if (rule.ContainsPatternFormattingBlock)
            {
                var pattern = rule.PatternFormatting;
                if (pattern == null || (pattern.FillPattern != FillPattern.NoFill
                    && pattern.FillPattern != FillPattern.SolidForeground))
                    return false;
            }
            if (rule.ContainsBorderFormattingBlock)
            {
                var border = rule.BorderFormatting;
                if (border == null || border.BorderDiagonal != BorderStyle.None
                    || border.IsForwardDiagonalOn || border.IsBackwardDiagonalOn)
                    return false;
            }
            if (!rule.ContainsFontFormattingBlock && !rule.ContainsPatternFormattingBlock
                && !rule.ContainsBorderFormattingBlock)
                return true;

            var fontFormatting = rule.FontFormatting;
            var patternFormatting = rule.PatternFormatting;
            var borderFormatting = rule.BorderFormatting;
            style = new ExcelDataReaderFailureFontStyle(
                fontFormatting?.IsBold == true,
                fontFormatting?.IsItalic == true,
                fontFormatting?.IsStruckout == true,
                fontFormatting?.UnderlineType != FontUnderlineType.None,
                fontHeightInPoints: fontFormatting?.FontHeight > 0
                    ? fontFormatting.FontHeight / 20d : 0,
                fontColorIndex: fontFormatting?.FontColorIndex >= 0
                    ? fontFormatting.FontColorIndex : null,
                fillPattern: patternFormatting?.FillPattern == FillPattern.SolidForeground
                    ? "solid" : null,
                fillForegroundColorIndex: patternFormatting?.FillForegroundColor >= 0
                    ? patternFormatting.FillForegroundColor : null,
                fillBackgroundColorIndex: patternFormatting?.FillBackgroundColor >= 0
                    ? patternFormatting.FillBackgroundColor : null,
                borderLeft: ToConditionalBorderSide(borderFormatting?.BorderLeft,
                    borderFormatting?.LeftBorderColor),
                borderRight: ToConditionalBorderSide(borderFormatting?.BorderRight,
                    borderFormatting?.RightBorderColor),
                borderTop: ToConditionalBorderSide(borderFormatting?.BorderTop,
                    borderFormatting?.TopBorderColor),
                borderBottom: ToConditionalBorderSide(borderFormatting?.BorderBottom,
                    borderFormatting?.BottomBorderColor));
            return true;
        }

        /// <summary>
        /// 将 HSSF 条件格式边框转换为低保留边框状态。
        /// </summary>
        private static ExcelDataReaderFailureBorderSide ToConditionalBorderSide(
            BorderStyle? style, short? color)
            => !style.HasValue || style.Value == BorderStyle.None ? null
                : new ExcelDataReaderFailureBorderSide(ToBorderStyleName(style.Value),
                    color.HasValue && color.Value >= 0 ? color : null);

        /// <summary>
        /// 将 HSSF 边框枚举转换为低保留边框名称。
        /// </summary>
        private static string ToBorderStyleName(BorderStyle style)
            => style switch
            {
                BorderStyle.Thin => "thin",
                BorderStyle.Medium => "medium",
                BorderStyle.Dashed => "dashed",
                BorderStyle.Dotted => "dotted",
                BorderStyle.Thick => "thick",
                BorderStyle.Double => "double",
                BorderStyle.Hair => "hair",
                BorderStyle.MediumDashed => "mediumDashed",
                BorderStyle.DashDot => "dashDot",
                BorderStyle.MediumDashDot => "mediumDashDot",
                BorderStyle.DashDotDot => "dashDotDot",
                BorderStyle.MediumDashDotDot => "mediumDashDotDot",
                BorderStyle.SlantedDashDot => "slantDashDot",
                _ => null
            };

        /// <summary>
        /// 读取可以在稀疏 HSSF 工作簿中安全重建的原生校验。
        /// </summary>
        /// <param name="record">HSSF 原生校验记录。</param>
        /// <param name="validations">按区域展开的校验描述。</param>
        /// <returns>校验只使用受支持的常量、显式列表或命名列表时返回 <see langword="true" />。</returns>
        private bool TryReadNativeDataValidations(DVRecord record,
            out IReadOnlyList<ExcelDataReaderFailureNativeDataValidation> validations)
        {
            validations = Array.Empty<ExcelDataReaderFailureNativeDataValidation>();
            if (record == null || record.CellRangeAddress?.CellRangeAddresses == null)
                return false;

            string formula1 = null;
            string formula2 = null;
            string listSourceName = null;
            string listSourceLocalSheetName = null;
            IReadOnlyList<string> listValues = null;
            if (record.DataType == NPOI.SS.UserModel.ValidationType.LIST)
            {
                if (record.ListExplicitFormula)
                {
                    if (!TryReadExplicitListFormula(record.Formula1, out listValues,
                        _requireNativeDataValidation))
                        return false;
                }
                else if (!TryReadNamedListFormula(record.Formula1, out listSourceName,
                    out listSourceLocalSheetName))
                    return false;
            }
            else if (record.ListExplicitFormula
                || (record.DataType != NPOI.SS.UserModel.ValidationType.INTEGER
                    && record.DataType != NPOI.SS.UserModel.ValidationType.DECIMAL
                    && record.DataType != NPOI.SS.UserModel.ValidationType.TEXT_LENGTH)
                || !TryReadSimpleFormula(record.Formula1, out formula1)
                || !TryReadSimpleFormula(record.Formula2, out formula2))
                return false;

            var result = new List<ExcelDataReaderFailureNativeDataValidation>();
            foreach (var region in record.CellRangeAddress.CellRangeAddresses)
            {
                if (region == null || region.FirstRow < 0 || region.LastRow < region.FirstRow
                    || region.LastColumn < region.FirstColumn || region.FirstColumn < 0
                    || region.LastRow > 65_535 || region.LastColumn > 255)
                    return false;
                result.Add(new ExcelDataReaderFailureNativeDataValidation(
                    region.FirstRow, region.LastRow, region.FirstColumn, region.LastColumn,
                    record.DataType, record.ConditionOperator, formula1, formula2,
                    record.EmptyCellAllowed, record.SuppressDropdownArrow,
                    record.ShowPromptOnCellSelected, NormalizeValidationText(record.PromptTitle),
                    NormalizeValidationText(record.PromptText), record.ShowErrorOnInvalidValue,
                    NormalizeValidationText(record.ErrorTitle), NormalizeValidationText(record.ErrorText),
                    record.ErrorStyle, listSourceName, listSourceLocalSheetName, listValues));
            }
            validations = result;
            return result.Count > 0;
        }

        /// <summary>
        /// 读取 HSSF 记录中的显式列表值。
        /// </summary>
        /// <param name="formula">显式列表公式令牌。</param>
        /// <param name="values">分隔后的列表值。</param>
        /// <param name="requireSafeInlineFormula">是否要求值可以安全写入 OOXML 内联公式。</param>
        /// <returns>列表只包含非空、可在 HSSF 中重建的值时返回 <see langword="true" />。</returns>
        private static bool TryReadExplicitListFormula(Ptg[] formula,
            out IReadOnlyList<string> values, bool requireSafeInlineFormula = false)
        {
            values = Array.Empty<string>();
            if (formula == null || formula.Length != 1 || !(formula[0] is StringPtg text)
                || string.IsNullOrEmpty(text.Value))
                return false;
            var items = text.Value.Split('\0', StringSplitOptions.None);
            if (items.Length == 0 || items.Any(string.IsNullOrEmpty)
                || items.Any(item => item.Length > 255)
                || (requireSafeInlineFormula && items.Any(item => item.IndexOf(',') >= 0
                    || item.IndexOf('"') >= 0
                    || item.IndexOf('\r') >= 0
                    || item.IndexOf('\n') >= 0))
                || items.Sum(item => item.Length) + items.Length - 1 > 255)
                return false;
            values = items;
            return true;
        }

        /// <summary>
        /// 读取指向安全一维名称区域的命名列表公式。
        /// </summary>
        /// <param name="formula">HSSF 列表公式令牌。</param>
        /// <param name="name">命名列表名称。</param>
        /// <param name="localSheetName">局部名称作用域工作表；全局名称为 null。</param>
        /// <returns>名称、作用域和来源均可投影时返回 <see langword="true" />。</returns>
        private bool TryReadNamedListFormula(Ptg[] formula, out string name,
            out string localSheetName)
        {
            name = null;
            localSheetName = null;
            if (formula == null || formula.Length != 1 || !(formula[0] is NamePtg nameToken)
                || nameToken.Index < 0 || nameToken.Index >= _nameRecords.Count)
                return false;

            if (!TryGetNamedListResolution(nameToken.Index, out var resolution))
                return false;

            if (!string.IsNullOrWhiteSpace(resolution.LocalSheetName))
            {
                localSheetName = resolution.LocalSheetName;
                if (_selectedSheetNames.Count == 0
                    || !string.Equals(localSheetName, _selectedSheetNames[^1],
                        StringComparison.OrdinalIgnoreCase))
                    return false;
            }

            if (!_requestedSheetNames.Contains(resolution.TargetSheetName))
                return false;

            name = resolution.Name;
            var resolvedName = name;
            var resolvedLocalSheetName = localSheetName;
            if (!_referencedNamedLists.Any(item =>
                string.Equals(item.Name, resolvedName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(item.LocalSheetName, resolvedLocalSheetName,
                    StringComparison.OrdinalIgnoreCase)))
                _referencedNamedLists.Add((resolvedName, resolvedLocalSheetName));
            return true;
        }

        /// <summary>
        /// 解析命名列表的一维绝对区域所在工作表。
        /// </summary>
        /// <param name="name">HSSF 名称记录。</param>
        /// <param name="localSheetName">局部名称作用域工作表。</param>
        /// <param name="targetSheetName">名称区域所在工作表。</param>
        /// <param name="range">名称区域的零基坐标。</param>
        /// <returns>名称引用为单格或一维绝对区域时返回 <see langword="true" />。</returns>
        private bool TryResolveNamedListTarget(NameRecord name, string localSheetName,
            out string targetSheetName, out NamedListRange range)
        {
            targetSheetName = null;
            range = null;
            var definition = name?.NameDefinition;
            if (definition == null || definition.Length != 1)
                return false;
            switch (definition[0])
            {
                case Ref3DPtg reference3D:
                    if (reference3D.IsRowRelative || reference3D.IsColRelative)
                        return false;
                    targetSheetName = ResolveReferenceSheetName(reference3D.ExternSheetIndex);
                    range = new NamedListRange(reference3D.Row, reference3D.Row,
                        reference3D.Column, reference3D.Column);
                    return !string.IsNullOrWhiteSpace(targetSheetName) && range.IsValid;
                case Area3DPtg area3D:
                    if (area3D.IsFirstRowRelative || area3D.IsFirstColRelative
                        || area3D.IsLastRowRelative || area3D.IsLastColRelative
                        || area3D.FirstRow > area3D.LastRow
                        || area3D.FirstColumn > area3D.LastColumn
                        || (area3D.FirstRow != area3D.LastRow
                            && area3D.FirstColumn != area3D.LastColumn))
                        return false;
                    targetSheetName = ResolveReferenceSheetName(area3D.ExternSheetIndex);
                    range = new NamedListRange(area3D.FirstRow, area3D.LastRow,
                        area3D.FirstColumn, area3D.LastColumn);
                    return !string.IsNullOrWhiteSpace(targetSheetName) && range.IsValid;
                case RefPtg reference:
                    if (string.IsNullOrWhiteSpace(localSheetName)
                        || reference.IsRowRelative || reference.IsColRelative)
                        return false;
                    targetSheetName = localSheetName;
                    range = new NamedListRange(reference.Row, reference.Row,
                        reference.Column, reference.Column);
                    return range.IsValid;
                case AreaPtg area:
                    if (string.IsNullOrWhiteSpace(localSheetName)
                        || area.IsFirstRowRelative || area.IsFirstColRelative
                        || area.IsLastRowRelative || area.IsLastColRelative
                        || area.FirstRow > area.LastRow || area.FirstColumn > area.LastColumn
                        || (area.FirstRow != area.LastRow && area.FirstColumn != area.LastColumn))
                        return false;
                    targetSheetName = localSheetName;
                    range = new NamedListRange(area.FirstRow, area.LastRow,
                        area.FirstColumn, area.LastColumn);
                    return range.IsValid;
                default:
                    return false;
            }
        }

        /// <summary>
        /// 将外部 Sheet 索引解析为单一源工作表名称。
        /// </summary>
        /// <param name="externSheetIndex">名称令牌中的外部 Sheet 索引。</param>
        /// <returns>能够解析时返回工作表名称。</returns>
        private string ResolveReferenceSheetName(int externSheetIndex)
        {
            var sheetIndex = ResolveReferenceSheetIndex(externSheetIndex);
            return sheetIndex >= 0 && sheetIndex < _sheetNames.Count
                ? _sheetNames[sheetIndex] : null;
        }

        /// <summary>
        /// 将简单 HSSF 常量公式转换为不带等号的边界文本。
        /// </summary>
        /// <param name="formula">HSSF 公式令牌。</param>
        /// <param name="value">转换后的常量文本。</param>
        /// <returns>公式只包含一个数字常量时返回 <see langword="true" />。</returns>
        private static bool TryReadSimpleFormula(Ptg[] formula, out string value)
        {
            value = null;
            if (formula == null || formula.Length == 0)
                return true;
            if (formula.Length != 1)
                return false;
            switch (formula[0])
            {
                case NumberPtg number:
                    value = number.Value.ToString("R", CultureInfo.InvariantCulture);
                    return true;
                case IntPtg integer:
                    value = integer.Value.ToString(CultureInfo.InvariantCulture);
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// 规范化 HSSF 记录中的空提示文本。
        /// </summary>
        /// <param name="value">原始提示文本。</param>
        /// <returns>有效文本或空值。</returns>
        private static string NormalizeValidationText(string value)
            => string.IsNullOrEmpty(value) || value == "\0" ? null : value;

        /// <summary>
        /// 保存 HSSF 列宽和隐藏状态，并限制到 XLSX 可表示的列范围。
        /// </summary>
        /// <param name="record">HSSF 列信息记录。</param>
        private void CaptureColumnInfo(ColumnInfoRecord record)
        {
            if (record == null || _columnMetadata == null)
                return;
            var first = Math.Max(0, record.FirstColumn);
            var last = Math.Min(16_383, record.LastColumn);
            if (last < first)
                return;
            var width = record.ColumnWidth / 256d;
            var hasWidth = width > 0 && !double.IsNaN(width) && !double.IsInfinity(width);
            for (var column = first; column <= last; column++)
                _columnMetadata[column] = (width, hasWidth, record.IsHidden);
        }

        /// <summary>
        /// 保存 HSSF 行高和隐藏状态。
        /// </summary>
        /// <param name="record">HSSF 行记录。</param>
        private void CaptureRowInfo(RowRecord record)
        {
            if (record == null || _rowMetadata == null || record.RowNumber < 0)
                return;
            var hasHeight = record.Height > 0;
            _rowMetadata[record.RowNumber] = (record.Height, hasHeight, record.ZeroHeight);
        }

        /// <summary>
        /// 保存 HSSF 合并区域。
        /// </summary>
        /// <param name="record">HSSF 合并区域记录。</param>
        private void CaptureMergeRegions(MergeCellsRecord record)
        {
            if (record == null || _mergeRegions == null)
                return;
            for (var index = 0; index < record.NumAreas; index++)
            {
                var area = record.GetAreaAt(index);
                if (area == null)
                    continue;
                var firstRow = Math.Max(0, area.FirstRow);
                var lastRow = Math.Min(65_535, area.LastRow);
                var firstColumn = Math.Max(0, area.FirstColumn);
                var lastColumn = Math.Min(255, area.LastColumn);
                if (lastRow < firstRow || lastColumn < firstColumn)
                    continue;
                var region = (firstRow, lastRow, firstColumn, lastColumn);
                if (!_mergeRegions.Contains(region))
                    _mergeRegions.Add(region);
            }
        }

        /// <summary>
        /// 解析共享字符串。
        /// </summary>
        private string ResolveString(int index)
        {
            if (_sharedStrings == null || index < 0 || index >= _sharedStrings.NumUniqueStrings)
                throw new BingOfficesImportException("HSSF shared string 索引无效。", provider: "NPOI",
                    stage: BingOfficesStage.Read);
            return _sharedStrings.GetString(index)?.String ?? string.Empty;
        }

        /// <summary>
        /// 写入文本单元格。
        /// </summary>
        private void WriteText(int row, int column, string value, CellValueRecordInterface cell)
        {
            CaptureNamedReferenceCell(row, column, value ?? string.Empty);
            WriteCell(row, column, GetProjectionStyleIndex(cell), writer =>
            {
                writer.WriteAttributeString("t", "inlineStr");
                writer.WriteStartElement("is", SpreadsheetNamespace);
                writer.WriteStartElement("t", SpreadsheetNamespace);
                writer.WriteAttributeString("xml", "space",
                    "http://www.w3.org/XML/1998/namespace", "preserve");
                writer.WriteString(value ?? string.Empty);
                writer.WriteEndElement();
                writer.WriteEndElement();
            });
        }

        /// <summary>
        /// 写入数值单元格。
        /// </summary>
        private void WriteNumber(int row, int column, double value, CellValueRecordInterface cell)
        {
            CaptureNamedReferenceCell(row, column, value);
            WriteCell(row, column, GetProjectionStyleIndex(cell), writer =>
            {
                writer.WriteStartElement("v", SpreadsheetNamespace);
                writer.WriteString(value.ToString("R", CultureInfo.InvariantCulture));
                writer.WriteEndElement();
            });
        }

        /// <summary>
        /// 将 HSSF 数字格式映射为投影 XLSX 的样式索引。
        /// </summary>
        private int GetProjectionStyleIndex(CellValueRecordInterface cell)
        {
            var formatIndex = _formatTracker?.GetFormatIndex(cell) ?? 0;
            var formatString = _formatTracker?.GetFormatString(formatIndex) ?? string.Empty;
            var xfIndex = cell?.XFIndex ?? -1;
            return GetStyleIndex(formatString, GetFontDefinition(xfIndex),
                GetFillDefinition(xfIndex), GetBorderDefinition(xfIndex),
                GetAlignmentDefinition(xfIndex), GetProtectionDefinition(xfIndex));
        }

        /// <summary>
        /// 根据格式和字体状态创建投影样式索引。
        /// </summary>
        private int GetStyleIndex(string formatString, FontDefinition font, FillDefinition fill,
            BorderDefinition border, AlignmentDefinition alignment,
            ProtectionDefinition protection)
        {
            font ??= _fontDefinitions[0];
            fill = fill?.IsEmpty == true ? null : fill;
            border = border?.IsEmpty == true ? null : border;
            alignment = alignment?.IsEmpty == true ? null : alignment;
            protection = protection?.IsEmpty == true ? null : protection;
            if (string.IsNullOrWhiteSpace(formatString)
                || string.Equals(formatString, "General", StringComparison.OrdinalIgnoreCase))
            {
                if (font.Key == _fontDefinitions[0].Key && fill == null && border == null
                    && alignment == null && protection == null)
                    return 0;
            }
            var key = (formatString ?? string.Empty) + "\0" + font.Key + "\0"
                + (fill?.Key ?? string.Empty) + "\0" + (border?.Key ?? string.Empty)
                + "\0" + (alignment?.Key ?? string.Empty)
                + "\0" + (protection?.Key ?? string.Empty);
            if (_formatStyles.TryGetValue(key, out var styleIndex))
                return styleIndex;
            var numberFormatId = 0;
            if (!string.IsNullOrWhiteSpace(formatString)
                && !string.Equals(formatString, "General", StringComparison.OrdinalIgnoreCase))
            {
                if (!_customFormats.Contains(formatString, StringComparer.Ordinal))
                    _customFormats.Add(formatString);
                numberFormatId = 164 + _customFormats.IndexOf(formatString);
            }
            styleIndex = _styleDefinitions.Count + 1;
            _styleDefinitions.Add(new StyleDefinition(numberFormatId, GetFontId(font),
                GetFillId(fill), GetBorderId(border), alignment, protection));
            _formatStyles.Add(key, styleIndex);
            return styleIndex;
        }

        /// <summary>
        /// 获取 HSSF 单元格使用的最小字体状态。
        /// </summary>
        private FontDefinition GetFontDefinition(short xfIndex)
        {
            if (xfIndex < 0 || xfIndex >= _xfStyles.Count)
                return _fontDefinitions[0];
            var fontIndex = _xfStyles[xfIndex].FontIndex;
            // HSSF 在字体索引 4 处保留一个未发送到事件流的槽位。
            if (fontIndex >= 5 && fontIndex - 1 < _sourceFonts.Count)
                return _sourceFonts[fontIndex - 1];
            if (fontIndex >= 0 && fontIndex < _sourceFonts.Count)
                return _sourceFonts[fontIndex];
            if (fontIndex > 0 && fontIndex - 1 < _sourceFonts.Count)
                return _sourceFonts[fontIndex - 1];
            return _fontDefinitions[0];
        }

        /// <summary>
        /// 获取 HSSF 单元格使用的最小对齐状态。
        /// </summary>
        private AlignmentDefinition GetAlignmentDefinition(short xfIndex)
        {
            if (xfIndex < 0 || xfIndex >= _xfStyles.Count)
                return null;
            return _xfStyles[xfIndex].Alignment;
        }

        /// <summary>
        /// 获取 HSSF 单元格使用的保护状态。
        /// </summary>
        private ProtectionDefinition GetProtectionDefinition(short xfIndex)
        {
            if (xfIndex < 0 || xfIndex >= _xfStyles.Count)
                return null;
            return _xfStyles[xfIndex].Protection;
        }

        /// <summary>
        /// 获取 HSSF 单元格使用的最小填充状态。
        /// </summary>
        private FillDefinition GetFillDefinition(short xfIndex)
        {
            if (xfIndex < 0 || xfIndex >= _xfStyles.Count)
                return null;
            return _xfStyles[xfIndex].Fill;
        }

        /// <summary>
        /// 获取 HSSF 单元格使用的标准四边框状态。
        /// </summary>
        private BorderDefinition GetBorderDefinition(short xfIndex)
        {
            if (xfIndex < 0 || xfIndex >= _xfStyles.Count)
                return null;
            return _xfStyles[xfIndex].Border;
        }

        /// <summary>
        /// 将 HSSF XF 对齐状态转换为 XLSX 对齐描述。
        /// </summary>
        private static AlignmentDefinition ToAlignmentDefinition(ExtendedFormatRecord record)
        {
            if (record == null)
                return null;
            var alignment = new AlignmentDefinition(
                ToHorizontalAlignment(record.Alignment),
                ToVerticalAlignment(record.VerticalAlignment),
                record.WrapText,
                record.ShrinkToFit,
                record.Rotation,
                record.Indent);
            return alignment.IsEmpty ? null : alignment;
        }

        /// <summary>
        /// 将 HSSF XF 保护状态转换为投影描述。
        /// </summary>
        private static ProtectionDefinition ToProtectionDefinition(ExtendedFormatRecord record)
        {
            if (record == null)
                return null;
            var protection = new ProtectionDefinition(record.IsLocked, record.IsHidden);
            return protection.IsEmpty ? null : protection;
        }

        /// <summary>
    /// 将 HSSF XF 纯色填充转换为投影描述。
        /// </summary>
        private static FillDefinition ToFillDefinition(ExtendedFormatRecord record)
        {
            if (record == null || record.AdtlFillPattern != 1
                || record.FillForeground < 0
                || record.FillForeground == HSSFColor.Automatic.Index)
                return null;
            return new FillDefinition("solid", record.FillForeground, record.FillBackground);
        }

        /// <summary>
        /// 将 HSSF XF 标准四边框转换为投影描述。
        /// </summary>
        private static BorderDefinition ToBorderDefinition(ExtendedFormatRecord record)
        {
            if (record == null)
                return null;
            var border = new BorderDefinition(
                ToBorderSide(record.BorderLeft, record.LeftBorderPaletteIdx),
                ToBorderSide(record.BorderRight, record.RightBorderPaletteIdx),
                ToBorderSide(record.BorderTop, record.TopBorderPaletteIdx),
                ToBorderSide(record.BorderBottom, record.BottomBorderPaletteIdx));
            return border.IsEmpty ? null : border;
        }

        /// <summary>
        /// 将 HSSF 边框线型和索引色转换为 OOXML 状态。
        /// </summary>
        private static BorderSideDefinition ToBorderSide(short style, short color)
        {
            var name = style switch
            {
                1 => "thin",
                2 => "medium",
                3 => "dashed",
                4 => "dotted",
                5 => "thick",
                6 => "double",
                7 => "hair",
                8 => "mediumDashed",
                9 => "dashDot",
                10 => "mediumDashDot",
                11 => "dashDotDot",
                12 => "mediumDashDotDot",
                13 => "slantDashDot",
                _ => null
            };
            if (name == null)
                return null;
            var colorIndex = color >= 0 && color != HSSFColor.Automatic.Index
                ? (short?)color : null;
            return new BorderSideDefinition(name, colorIndex);
        }

        private static string ToHorizontalAlignment(short value) => value switch
        {
            1 => "left",
            2 => "center",
            3 => "right",
            4 => "fill",
            5 => "justify",
            6 => "centerContinuous",
            7 => "distributed",
            _ => null
        };

        private static string ToVerticalAlignment(short value) => value switch
        {
            1 => "center",
            2 => "top",
            3 => "justify",
            4 => "distributed",
            _ => null
        };

        /// <summary>
        /// 获取或创建投影字体索引。
        /// </summary>
        private int GetFontId(FontDefinition font)
        {
            if (font == null || font.Key == _fontDefinitions[0].Key)
                return 0;
            if (_fontStyleIds.TryGetValue(font.Key, out var fontId))
                return fontId;
            fontId = _fontDefinitions.Count;
            _fontDefinitions.Add(font);
            _fontStyleIds.Add(font.Key, fontId);
            return fontId;
        }

        /// <summary>
        /// 获取或创建投影填充索引。
        /// </summary>
        private int GetFillId(FillDefinition fill)
        {
            if (fill == null || fill.IsEmpty)
                return 0;
            if (_fillStyleIds.TryGetValue(fill.Key, out var fillId))
                return fillId;
            fillId = _fillDefinitions.Count + 2;
            _fillDefinitions.Add(fill);
            _fillStyleIds.Add(fill.Key, fillId);
            return fillId;
        }

        /// <summary>
        /// 获取或创建投影边框索引。
        /// </summary>
        private int GetBorderId(BorderDefinition border)
        {
            if (border == null || border.IsEmpty)
                return 0;
            if (_borderStyleIds.TryGetValue(border.Key, out var borderId))
                return borderId;
            borderId = _borderDefinitions.Count + 1;
            _borderDefinitions.Add(border);
            _borderStyleIds.Add(border.Key, borderId);
            return borderId;
        }

        /// <summary>
        /// 写入布尔单元格。
        /// </summary>
        private void WriteBoolean(int row, int column, bool value, CellValueRecordInterface cell)
        {
            CaptureNamedReferenceCell(row, column, value);
            WriteCell(row, column, GetProjectionStyleIndex(cell), writer =>
            {
                writer.WriteAttributeString("t", "b");
                writer.WriteStartElement("v", SpreadsheetNamespace);
                writer.WriteString(value ? "1" : "0");
                writer.WriteEndElement();
            });
        }

        /// <summary>
        /// 写入错误单元格。
        /// </summary>
        private void WriteError(int row, int column, string value, CellValueRecordInterface cell)
        {
            CaptureNamedReferenceCell(row, column, value ?? string.Empty);
            WriteCell(row, column, GetProjectionStyleIndex(cell), writer =>
            {
                writer.WriteAttributeString("t", "e");
                writer.WriteStartElement("v", SpreadsheetNamespace);
                writer.WriteString(value ?? string.Empty);
                writer.WriteEndElement();
            });
        }

        /// <summary>
        /// 暂存安全命名区域中的源单元格，供低保留失败工作簿恢复。
        /// </summary>
        /// <param name="row">零基行号。</param>
        /// <param name="column">零基列号。</param>
        /// <param name="value">单元格值。</param>
        private void CaptureNamedReferenceCell(int row, int column, object value)
        {
            if (_namedReferenceCells == null
                || !_namedReferenceRanges.Any(range => range.Contains(row, column)))
                return;
            if (!_namedReferenceCells.TryGetValue(row, out var cells))
            {
                cells = new Dictionary<int, object>();
                _namedReferenceCells[row] = cells;
            }
            cells[column] = value;
        }

        /// <summary>
        /// 将当前工作表的命名区域单元格写入失败快照。
        /// </summary>
        /// <param name="sheetName">当前工作表名称。</param>
        private void StoreNamedReferenceRows(string sheetName)
        {
            if (_failureCapture == null || _namedReferenceCells == null
                || string.IsNullOrWhiteSpace(sheetName))
                return;
            _namedReferenceCellSnapshots[sheetName] = _namedReferenceCells;
        }

        /// <summary>
        /// 将原生命名列表实际引用的来源行写入失败快照。
        /// </summary>
        private void CaptureNamedReferenceRows()
        {
            if (_failureCapture == null || _namedReferenceCellSnapshots.Count == 0)
                return;
            foreach (var snapshot in _namedReferenceCellSnapshots)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                var referencedRanges = GetReferencedNamedReferenceRanges(snapshot.Key);
                if (referencedRanges.Count == 0)
                    continue;
                foreach (var row in snapshot.Value.OrderBy(item => item.Key))
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    var cells = row.Value.Where(cell => referencedRanges.Any(range =>
                        range.Contains(row.Key, cell.Key))).ToArray();
                    if (cells.Length == 0)
                        continue;
                    var width = cells.Max(cell => cell.Key) + 1;
                    var values = new object[width];
                    foreach (var cell in cells)
                        values[cell.Key] = cell.Value;
                    _failureCapture.CaptureContextRow(snapshot.Key, row.Key, values);
                }
            }
            _namedReferenceCellSnapshots.Clear();
        }

        /// <summary>
        /// 获取当前工作表被原生命名列表校验实际引用的区域。
        /// </summary>
        /// <param name="sheetName">当前工作表名称。</param>
        /// <returns>命名列表实际引用的区域集合。</returns>
        private IReadOnlyList<NamedListRange> GetReferencedNamedReferenceRanges(
            string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName)
                || _referencedNamedLists.Count == 0)
                return Array.Empty<NamedListRange>();
            var ranges = new List<NamedListRange>();
            foreach (var reference in _referencedNamedLists)
            {
                foreach (var resolution in _namedListResolutionCache.Values)
                {
                    if (resolution == null || !resolution.IsSupported || resolution.Range == null
                        || !string.Equals(resolution.Name, reference.Name,
                            StringComparison.OrdinalIgnoreCase)
                        || !string.Equals(resolution.LocalSheetName,
                            reference.LocalSheetName,
                            StringComparison.OrdinalIgnoreCase)
                        || !string.Equals(resolution.TargetSheetName, sheetName,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (!ranges.Any(range => range.FirstRow == resolution.Range.FirstRow
                        && range.LastRow == resolution.Range.LastRow
                        && range.FirstColumn == resolution.Range.FirstColumn
                        && range.LastColumn == resolution.Range.LastColumn))
                        ranges.Add(resolution.Range);
                }
            }
            return ranges;
        }

        /// <summary>
        /// 写入单元格并检查单元格资源限制。
        /// </summary>
        private void WriteCell(int row, int column, int styleIndex, Action<XmlWriter> writeValue)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            if (_limits.MaxColumnsPerSheet.HasValue
                && column + 1 > _limits.MaxColumnsPerSheet.Value)
                throw ResourceLimit("工作表列数超过限制。");
            if (_limits.MaxCells.HasValue && ++_cellCount > _limits.MaxCells.Value)
                throw ResourceLimit("工作簿物理单元格数量超过限制。");
            _maxRow = Math.Max(_maxRow, row);
            _maxColumn = Math.Max(_maxColumn, column);
            _sheet.WriteCell(row, column, styleIndex, writeValue);
        }

        /// <summary>
        /// 检查工作表数量限制。
        /// </summary>
        private void CheckSheetLimit()
        {
            if (_limits.MaxSheets.HasValue && _sheetNames.Count > _limits.MaxSheets.Value)
                throw ResourceLimit("工作簿工作表数量超过限制。");
        }

        /// <summary>
        /// 创建资源限制异常。
        /// </summary>
        private static BingOfficesResourceLimitException ResourceLimit(string message) =>
            new(message, provider: "NPOI", operation: BingOfficesOperation.Import,
                stage: BingOfficesStage.Read);

        /// <summary>
        /// 创建不支持功能异常。
        /// </summary>
        private static BingOfficesUnsupportedFeatureException Unsupported(string feature) =>
            new($"NPOI HSSF Forward 实体批次暂不支持{feature}。", provider: "NPOI",
                operation: BingOfficesOperation.Import, stage: BingOfficesStage.Read);

        /// <summary>
        /// 判断事件记录是否属于条件格式定义。
        /// </summary>
        private static bool IsConditionalFormattingRecord(Record record)
        {
            var name = record?.GetType().Name;
            return name != null
                && (name.StartsWith("CFHeader", StringComparison.Ordinal)
                    || name.StartsWith("CFRule", StringComparison.Ordinal));
        }

        /// <summary>
        /// 将 HSSF 名称记录转换为临时 XLSX 的名称描述。
        /// </summary>
        private DefinedNameProjection CreateDefinedNameProjection(NameRecord name)
        {
            if (string.IsNullOrWhiteSpace(name.NameText))
                return null;
            var localSheetId = ResolveProjectedLocalSheetId(name.SheetNumber);
            var formula = CreateNameFormula(name, localSheetId);
            return new DefinedNameProjection(name.NameText, localSheetId, formula);
        }

        /// <summary>
        /// 解析局部名称作用域在投影工作簿中的 Sheet 序号。
        /// </summary>
        private int ResolveProjectedLocalSheetId(int sheetNumber)
        {
            if (sheetNumber <= 0)
                return -1;
            var sourceIndex = sheetNumber - 1;
            if (sourceIndex < 0 || sourceIndex >= _sheetNames.Count)
                return -2;
            var sheetName = _sheetNames[sourceIndex];
            for (var index = 0; index < _selectedSheetNames.Count; index++)
            {
                if (string.Equals(_selectedSheetNames[index], sheetName,
                    StringComparison.OrdinalIgnoreCase))
                    return index;
            }
            return -2;
        }

        /// <summary>
        /// 将受限名称 token 转换为单格或区域公式。
        /// </summary>
        private string CreateNameFormula(NameRecord name, int localSheetId)
        {
            var definition = name.NameDefinition;
            if (definition == null || definition.Length != 1)
                return "#REF!";
            switch (definition[0])
            {
                case Ref3DPtg reference3D:
                    return CreateReferenceFormula(reference3D);
                case Area3DPtg area3D:
                    return CreateAreaFormula(area3D);
                case RefPtg reference:
                    return CreateLocalReferenceFormula(reference, localSheetId);
                case AreaPtg area:
                    return CreateLocalAreaFormula(area, localSheetId);
                default:
                    return "#REF!";
            }
        }

        /// <summary>
        /// 将跨工作表单格 token 转换为 A1 公式。
        /// </summary>
        private string CreateReferenceFormula(Ref3DPtg reference)
        {
            if (reference.IsRowRelative || reference.IsColRelative)
                return "#REF!";
            var sheetIndex = ResolveReferenceSheetIndex(reference.ExternSheetIndex);
            if (sheetIndex < 0 || sheetIndex >= _sheetNames.Count)
                return "#REF!";
            return CreateCellFormula(_sheetNames[sheetIndex], reference.Row, reference.Column);
        }

        /// <summary>
        /// 将跨工作表区域 token 转换为 A1 公式。
        /// </summary>
        private string CreateAreaFormula(Area3DPtg area)
        {
            if (area.IsFirstRowRelative || area.IsFirstColRelative
                || area.IsLastRowRelative || area.IsLastColRelative)
                return "#REF!";
            var sheetIndex = ResolveReferenceSheetIndex(area.ExternSheetIndex);
            if (sheetIndex < 0 || sheetIndex >= _sheetNames.Count)
                return "#REF!";
            return CreateAreaFormula(_sheetNames[sheetIndex], area.FirstRow, area.FirstColumn,
                area.LastRow, area.LastColumn);
        }

        /// <summary>
        /// 将局部名称单格 token 转换为 A1 公式。
        /// </summary>
        private string CreateLocalReferenceFormula(RefPtg reference, int localSheetId)
        {
            if (localSheetId < 0 || reference.IsRowRelative || reference.IsColRelative
                || localSheetId >= _selectedSheetNames.Count)
                return "#REF!";
            return CreateCellFormula(_selectedSheetNames[localSheetId], reference.Row,
                reference.Column);
        }

        /// <summary>
        /// 将局部名称区域 token 转换为 A1 公式。
        /// </summary>
        private string CreateLocalAreaFormula(AreaPtg area, int localSheetId)
        {
            if (localSheetId < 0 || localSheetId >= _selectedSheetNames.Count
                || area.IsFirstRowRelative || area.IsFirstColRelative
                || area.IsLastRowRelative || area.IsLastColRelative)
                return "#REF!";
            return CreateAreaFormula(_selectedSheetNames[localSheetId], area.FirstRow,
                area.FirstColumn, area.LastRow, area.LastColumn);
        }

        /// <summary>
        /// 按外部 Sheet 索引解析源工作表序号。
        /// </summary>
        private int ResolveReferenceSheetIndex(int externSheetIndex)
        {
            var resolved = -1;
            foreach (var record in _externSheetRecords)
            {
                if (externSheetIndex < 0 || externSheetIndex >= record.NumOfREFRecords)
                    continue;
                var first = record.GetFirstSheetIndexFromRefIndex(externSheetIndex);
                var last = record.GetLastSheetIndexFromRefIndex(externSheetIndex);
                if (first < 0 || first != last)
                    continue;
                if (resolved >= 0 && resolved != first)
                    return -1;
                resolved = first;
            }
            return resolved;
        }

        /// <summary>
        /// 创建单格公式。
        /// </summary>
        private static string CreateCellFormula(string sheetName, int row, int column) =>
            $"'{EscapeSheetName(sheetName)}'!{ToAbsoluteCellReference(row, column)}";

        /// <summary>
        /// 创建区域公式。
        /// </summary>
        private static string CreateAreaFormula(string sheetName, int firstRow, int firstColumn,
            int lastRow, int lastColumn) =>
            $"'{EscapeSheetName(sheetName)}'!{ToAbsoluteCellReference(firstRow, firstColumn)}"
            + $":{ToAbsoluteCellReference(lastRow, lastColumn)}";

        /// <summary>
        /// 转义工作表名称中的单引号。
        /// </summary>
        private static string EscapeSheetName(string value) => (value ?? string.Empty).Replace("'", "''");

        /// <summary>
        /// 将零基坐标转换为绝对 A1 引用。
        /// </summary>
        private static string ToAbsoluteCellReference(int row, int column) =>
            "$" + ToColumnReference(column) + "$" + (row + 1).ToString(CultureInfo.InvariantCulture);

        /// <summary>
        /// 将零基列转换为列字母。
        /// </summary>
        private static string ToColumnReference(int column)
        {
            var value = column + 1;
            var builder = new StringBuilder();
            while (value > 0)
            {
                var remainder = (value - 1) % 26;
                builder.Insert(0, (char)('A' + remainder));
                value = (value - 1) / 26;
            }
            return builder.ToString();
        }
    }

    /// <summary>
    /// 写入可供 Forward 读取器执行的原生数据校验。
    /// </summary>
    private static void WriteNativeDataValidations(XmlWriter writer,
        IReadOnlyList<ExcelDataReaderFailureNativeDataValidation> validations)
    {
        if (validations == null || validations.Count == 0)
            return;
        writer.WriteStartElement("dataValidations", SpreadsheetNamespace);
        writer.WriteAttributeString("count", validations.Count.ToString(CultureInfo.InvariantCulture));
        foreach (var validation in validations)
        {
            if (validation == null)
                continue;
            var type = ToOpenXmlValidationType(validation.ValidationType);
            if (type == null)
                continue;
            writer.WriteStartElement("dataValidation", SpreadsheetNamespace);
            writer.WriteAttributeString("type", type);
            if (type != "list")
                WriteOptionalAttribute(writer, "operator",
                    ToOpenXmlValidationOperator(validation.ComparisonOperator));
            if (validation.EmptyCellAllowed)
                writer.WriteAttributeString("allowBlank", "1");
            if (validation.SuppressDropdownArrow)
                writer.WriteAttributeString("showDropDown", "1");
            if (validation.ShowPromptBox)
                writer.WriteAttributeString("showInputMessage", "1");
            if (validation.ShowErrorBox)
                writer.WriteAttributeString("showErrorMessage", "1");
            WriteOptionalAttribute(writer, "errorStyle",
                ToOpenXmlErrorStyle(validation.ErrorStyle));
            WriteOptionalAttribute(writer, "promptTitle", validation.PromptTitle);
            WriteOptionalAttribute(writer, "prompt", validation.PromptText);
            WriteOptionalAttribute(writer, "errorTitle", validation.ErrorTitle);
            WriteOptionalAttribute(writer, "error", validation.ErrorText);
            writer.WriteAttributeString("sqref", ToAbsoluteCellReference(validation.FirstRow,
                validation.FirstColumn) + ":"
                + ToAbsoluteCellReference(validation.LastRow, validation.LastColumn));

            var formula1 = validation.ValidationType == NPOI.SS.UserModel.ValidationType.LIST
                ? validation.ListValues != null && validation.ListValues.Count > 0
                    ? "\"" + string.Join(",", validation.ListValues) + "\""
                    : validation.ListSourceName
                : validation.Formula1;
            if (!string.IsNullOrWhiteSpace(formula1))
                writer.WriteElementString("formula1", SpreadsheetNamespace, formula1);
            if (!string.IsNullOrWhiteSpace(validation.Formula2))
                writer.WriteElementString("formula2", SpreadsheetNamespace, validation.Formula2);
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    /// <summary>
    /// 将 NPOI 校验类型转换为 OOXML 类型名称。
    /// </summary>
    private static string ToOpenXmlValidationType(int type) => type switch
    {
        NPOI.SS.UserModel.ValidationType.INTEGER => "whole",
        NPOI.SS.UserModel.ValidationType.DECIMAL => "decimal",
        NPOI.SS.UserModel.ValidationType.LIST => "list",
        NPOI.SS.UserModel.ValidationType.TEXT_LENGTH => "textLength",
        _ => null
    };

    /// <summary>
    /// 将 NPOI 比较运算符转换为 OOXML 名称。
    /// </summary>
    private static string ToOpenXmlValidationOperator(int operation) =>
        ((NPOI.SS.UserModel.ComparisonOperator)operation) switch
    {
        NPOI.SS.UserModel.ComparisonOperator.NoComparison => "between",
        NPOI.SS.UserModel.ComparisonOperator.Between => "between",
        NPOI.SS.UserModel.ComparisonOperator.NotBetween => "notBetween",
        NPOI.SS.UserModel.ComparisonOperator.Equal => "equal",
        NPOI.SS.UserModel.ComparisonOperator.NotEqual => "notEqual",
        NPOI.SS.UserModel.ComparisonOperator.GreaterThan => "greaterThan",
        NPOI.SS.UserModel.ComparisonOperator.LessThan => "lessThan",
        NPOI.SS.UserModel.ComparisonOperator.GreaterThanOrEqual => "greaterThanOrEqual",
        NPOI.SS.UserModel.ComparisonOperator.LessThanOrEqual => "lessThanOrEqual",
        _ => null
    };

    /// <summary>
    /// 将 HSSF 错误样式转换为 OOXML 名称。
    /// </summary>
    private static string ToOpenXmlErrorStyle(int errorStyle) => errorStyle switch
    {
        1 => "warning",
        2 => "information",
        _ => "stop"
    };

    /// <summary>
    /// 保存命名列表引用的一维 HSSF 区域。
    /// </summary>
    private sealed class NamedListRange
    {
        internal NamedListRange(int firstRow, int lastRow, int firstColumn,
            int lastColumn)
        {
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
        }

        internal int FirstRow { get; }

        internal int LastRow { get; }

        internal int FirstColumn { get; }

        internal int LastColumn { get; }

        internal bool IsValid => FirstRow >= 0 && LastRow >= FirstRow
            && LastRow <= 65_535 && FirstColumn >= 0
            && LastColumn >= FirstColumn && LastColumn <= 255
            && (FirstRow == LastRow || FirstColumn == LastColumn);

        internal bool Contains(int row, int column) => IsValid
            && row >= FirstRow && row <= LastRow
            && column >= FirstColumn && column <= LastColumn;
    }

    /// <summary>
    /// 保存命名列表名称的安全解析结果。
    /// </summary>
    private sealed class NamedListResolution
    {
        internal static readonly NamedListResolution Unsupported =
            new NamedListResolution(null, null, null, null, false);

        internal NamedListResolution(string name, string localSheetName,
            string targetSheetName, NamedListRange range, bool isSupported = true)
        {
            Name = name;
            LocalSheetName = localSheetName;
            TargetSheetName = targetSheetName;
            Range = range;
            IsSupported = isSupported;
        }

        internal string Name { get; }

        internal string LocalSheetName { get; }

        internal string TargetSheetName { get; }

        internal NamedListRange Range { get; }

        internal bool IsSupported { get; }
    }

    /// <summary>
    /// 表示投影工作簿中的命名区域。
    /// </summary>
    private sealed class DefinedNameProjection
    {
        internal DefinedNameProjection(string name, int localSheetId, string formula)
        {
            Name = name;
            LocalSheetId = localSheetId;
            Formula = formula;
        }

        internal string Name { get; }

        internal int LocalSheetId { get; }

        internal string Formula { get; }
    }

    /// <summary>
    /// 为 MulRKRecord 提供可查询格式的单元格视图。
    /// </summary>
    private sealed class CellFormatInfo : CellValueRecordInterface
    {
        internal CellFormatInfo(int row, int column, short xfIndex)
        {
            Row = row;
            Column = column;
            XFIndex = xfIndex;
        }

        public int Row { get; set; }
        public int Column { get; set; }
        public short XFIndex { get; set; }
    }

    /// <summary>
    /// 将目标工作表行单元格写入磁盘暂存文件。
    /// </summary>
    private sealed class SheetXmlBuffer : IDisposable
    {
        private readonly Stream _stream;
        private readonly XmlWriter _writer;
        private int _currentRow = -1;
        private bool _rowOpen;
        private bool _completed;

        internal SheetXmlBuffer(ExcelTemporaryStorageBudget temporaryStorage)
        {
            var path = Path.Combine(Path.GetTempPath(),
                $"bing-offices-forward-hssf-sheet-{Guid.NewGuid():N}.xml");
            _stream = (temporaryStorage ?? new ExcelTemporaryStorageBudget(null, "NPOI")).CreateFile(path,
                FileOptions.Asynchronous | FileOptions.SequentialScan | FileOptions.DeleteOnClose);
            _writer = XmlWriter.Create(_stream, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                OmitXmlDeclaration = false,
                CloseOutput = false
            });
            _writer.WriteStartDocument();
            _writer.WriteStartElement("worksheet", SpreadsheetNamespace);
            _writer.WriteStartElement("sheetData", SpreadsheetNamespace);
        }

        internal void WriteCell(int row, int column, int styleIndex, Action<XmlWriter> writeValue)
        {
            if (_completed)
                throw new InvalidOperationException("目标工作表已经完成写入。");
            if (row < _currentRow)
                throw new BingOfficesImportException("HSSF 工作表行顺序无效。", provider: "NPOI",
                    stage: BingOfficesStage.Read);
            if (row != _currentRow)
            {
                CloseRow();
                _currentRow = row;
                _writer.WriteStartElement("row", SpreadsheetNamespace);
                _writer.WriteAttributeString("r", (row + 1).ToString(CultureInfo.InvariantCulture));
                _rowOpen = true;
            }
            _writer.WriteStartElement("c", SpreadsheetNamespace);
            _writer.WriteAttributeString("r", ToCellReference(row, column));
            if (styleIndex != 0)
                _writer.WriteAttributeString("s", styleIndex.ToString(CultureInfo.InvariantCulture));
            writeValue(_writer);
            _writer.WriteEndElement();
        }

        /// <summary>
        /// 完成工作表 XML 写入并追加可安全投影的工作表元数据。
        /// </summary>
        /// <param name="validations">可安全投影的原生校验描述。</param>
        /// <param name="mergeRegions">可安全投影的合并区域。</param>
        internal void Complete(
            IReadOnlyList<ExcelDataReaderFailureNativeDataValidation> validations,
            IReadOnlyList<(int FirstRow, int LastRow, int FirstColumn, int LastColumn)> mergeRegions)
        {
            if (_completed)
                return;
            CloseRow();
            _writer.WriteEndElement();
            WriteMergeRegions(_writer, mergeRegions);
            WriteNativeDataValidations(_writer, validations);
            _writer.WriteEndElement();
            _writer.WriteEndDocument();
            _writer.Flush();
            _stream.Flush();
            _stream.Position = 0;
            _completed = true;
        }

        /// <summary>
        /// 写入可安全投影的合并区域。
        /// </summary>
        /// <param name="writer">工作表 XML 写入器。</param>
        /// <param name="mergeRegions">零基合并区域。</param>
        private static void WriteMergeRegions(XmlWriter writer,
            IReadOnlyList<(int FirstRow, int LastRow, int FirstColumn, int LastColumn)> mergeRegions)
        {
            if (mergeRegions == null || mergeRegions.Count == 0)
                return;

            var validRegions = mergeRegions.Where(region =>
                region.FirstRow >= 0 && region.LastRow >= region.FirstRow
                && region.LastRow <= 65_535 && region.FirstColumn >= 0
                && region.LastColumn >= region.FirstColumn && region.LastColumn <= 255)
                .ToArray();
            if (validRegions.Length == 0)
                return;

            writer.WriteStartElement("mergeCells", SpreadsheetNamespace);
            writer.WriteAttributeString("count", validRegions.Length.ToString(
                CultureInfo.InvariantCulture));
            foreach (var region in validRegions)
            {
                writer.WriteStartElement("mergeCell", SpreadsheetNamespace);
                writer.WriteAttributeString("ref", ToCellReference(region.FirstRow,
                    region.FirstColumn) + ":" + ToCellReference(region.LastRow,
                    region.LastColumn));
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        internal void CopyTo(Stream destination)
        {
            if (!_completed)
                throw new InvalidOperationException("目标工作表尚未完成写入。");
            _stream.Position = 0;
            _stream.CopyTo(destination);
        }

        /// <inheritdoc />
        public void Dispose()
        {
            try { _writer.Dispose(); }
            finally { _stream.Dispose(); }
        }

        private void CloseRow()
        {
            if (!_rowOpen)
                return;
            _writer.WriteEndElement();
            _rowOpen = false;
        }
    }

    /// <summary>
    /// 将零基列和行转换为 A1 引用。
    /// </summary>
    private static string ToCellReference(int row, int column) =>
        ToColumnReference(column) + (row + 1).ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// 将零基坐标转换为绝对 A1 引用。
    /// </summary>
    private static string ToAbsoluteCellReference(int row, int column) =>
        "$" + ToColumnReference(column) + "$"
        + (row + 1).ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// 将零基列转换为列字母。
    /// </summary>
    private static string ToColumnReference(int column)
    {
        var value = column + 1;
        var builder = new StringBuilder();
        while (value > 0)
        {
            var remainder = (value - 1) % 26;
            builder.Insert(0, (char)('A' + remainder));
            value = (value - 1) / 26;
        }
        return builder.ToString();
    }
}
