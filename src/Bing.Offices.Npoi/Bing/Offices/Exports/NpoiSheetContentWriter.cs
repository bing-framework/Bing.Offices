using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using System.Text;

namespace Bing.Offices.Exports;

/// <summary>
/// 写入 BIFF/OOXML 原生图片及校验规则。
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
    {
        if (request.Format != ExcelFormat.Xlsx || !request.Sheets.Any(sheet => sheet.DataValidations.Any(rule => rule.Type == ExcelDataValidationType.List)))
        {
            workbook.Write(destination, false);
            return;
        }
        // NPOI 2.7.4 将 CDATA 内的引号编码为字面量 &#34;，导致列表无效。
        // 暂存后流式重写 worksheet，只载入小型 dataValidations 节点，不载入数据行。
        var path = Path.Combine(Path.GetTempPath(), "bing-npoi-validations-" + Guid.NewGuid().ToString("N") + ".tmp");
        using var staged = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 81920, FileOptions.DeleteOnClose);
        workbook.Write(staged, true);
        staged.Position = 0;
        using var input = new ZipArchive(staged, ZipArchiveMode.Read, true);
        using var output = new ZipArchive(destination, ZipArchiveMode.Create, true);
        var rulesByPart = request.Sheets.Where(sheet => sheet.DataValidations.Any(rule => rule.Type == ExcelDataValidationType.List))
            .ToDictionary(sheet => ((NPOI.XSSF.UserModel.XSSFSheet)workbook.GetSheet(sheet.Name)).GetPackagePart().PartName.Name.TrimStart('/'),
                sheet => sheet.DataValidations.Where(rule => rule.Type == ExcelDataValidationType.List).ToArray());
        foreach (var entry in input.Entries)
        {
            token.ThrowIfCancellationRequested();
            using var source = entry.Open();
            using var target = output.CreateEntry(entry.FullName).Open();
            if (rulesByPart.TryGetValue(entry.FullName, out var definitions))
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
                foreach (var definition in definitions)
                {
                    var address = ExcelSheetContent.Address(definition.Range);
                    var rule = rules.Elements().Last(element => (string)element.Attribute("type") == "list"
                        && ((string)element.Attribute("sqref") == address || (string)element.Attribute("sqref") == address.Split(':')[0] && definition.Range.StartRow == definition.Range.EndRow && definition.Range.StartColumn == definition.Range.EndColumn));
                    rule.Elements().Single(element => element.Name.LocalName == "formula1").Value = "\"" + string.Join(",", definition.Values) + "\"";
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
        var helper = sheet.GetDataValidationHelper();
        foreach (var definition in request.DataValidations)
        {
            token.ThrowIfCancellationRequested();
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
                _ => helper.CreateExplicitListConstraint(definition.Values.ToArray())
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
