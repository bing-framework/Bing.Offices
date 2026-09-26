using SpreadCheetah;
using SpreadCheetah.Images;
using SpreadCheetah.Validations;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace Bing.Offices.Exports;

/// <summary>
/// 公共图片和数据校验的前向 XLSX 写入器。
/// </summary>
internal static class SpreadCheetahSheetContentWriter
{
    /// <summary>
    /// 嵌入 JPEG 前用于建立图片关系的单像素 PNG 占位内容。
    /// </summary>
    private static readonly byte[] JpegPlaceholder = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=");
    // 1.28.0 原生 drawing 使用不同的像素换算；只对有图片请求暂存并修正小型 drawing，
    // 保留数据行的前向写入，不将工作簿实体或图片副本累计到内存。
    /// <summary>
    /// 异步暂存导出结果并修正图片尺寸和编码。
    /// </summary>
    /// <param name="request">工作簿或工作表导出请求。</param>
    /// <param name="destination">接收输出的目标流。</param>
    /// <param name="write">将工作簿异步写入暂存流的回调。</param>
    /// <param name="token">取消操作的令牌。</param>
    internal static async Task ExportStagedAsync(ExcelWorkbookExportRequest request, Stream destination,
        Func<Stream, Task> write, CancellationToken token)
    {
        var path = Path.Combine(Path.GetTempPath(), "bing-spread-images-" + Guid.NewGuid().ToString("N") + ".tmp");
        await using var staged = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None,
            81920, FileOptions.Asynchronous | FileOptions.DeleteOnClose);
        await write(staged).ConfigureAwait(false);
        staged.Position = 0;
        using (var archive = new ZipArchive(staged, ZipArchiveMode.Update, leaveOpen: true))
        {
            XNamespace drawingNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
            var placeholders = new HashSet<string>(StringComparer.Ordinal);
            var pngParts = new HashSet<string>(StringComparer.Ordinal);
            for (var sheetIndex = 0; sheetIndex < request.Sheets.Count; sheetIndex++)
            {
                var images = request.Sheets[sheetIndex].Images;
                if (images.Count == 0) continue;
                token.ThrowIfCancellationRequested();
                var relationsPath = $"xl/worksheets/_rels/sheet{sheetIndex + 1}.xml.rels";
                using var relationsInput = archive.GetEntry(relationsPath).Open();
                var relations = XDocument.Load(relationsInput);
                var target = relations.Root.Elements().Single(element => ((string)element.Attribute("Type"))?.EndsWith("/drawing", StringComparison.Ordinal) == true).Attribute("Target").Value;
                var drawingPath = new Uri(new Uri($"https://package/xl/worksheets/sheet{sheetIndex + 1}.xml"), target).AbsolutePath.TrimStart('/');
                var entry = archive.GetEntry(drawingPath);
                XDocument document;
                using (var input = entry.Open()) document = XDocument.Load(input);
                var drawingRelationsPath = drawingPath.Insert(drawingPath.LastIndexOf('/') + 1, "_rels/") + ".rels";
                var drawingRelationsEntry = archive.GetEntry(drawingRelationsPath);
                XDocument drawingRelations;
                using (var input = drawingRelationsEntry.Open()) drawingRelations = XDocument.Load(input);
                var anchors = document.Root.Elements().ToArray();
                if (anchors.Length != images.Count) throw new InvalidDataException("图片锚点数量与请求不一致。");
                for (var i = 0; i < images.Count; i++)
                {
                    var image = images[i];
                    var anchor = anchors[i];
                    var relationId = anchor.Descendants().Single(element => element.Name.LocalName == "blip")
                        .Attributes().Single(attribute => attribute.Name.LocalName == "embed").Value;
                    var relation = drawingRelations.Root.Elements().Single(element => (string)element.Attribute("Id") == relationId);
                    var originalPart = new Uri(new Uri("https://package/" + drawingPath), (string)relation.Attribute("Target")).AbsolutePath.TrimStart('/');
                    if (image.Content[0] == 255)
                    {
                        // 固定版本仅原生嵌入 PNG；直接封装 JPEG 原始字节，不解码重编码。
                        var jpegPart = $"xl/media/bing-image-{sheetIndex}-{i}.jpeg";
                        using var jpegOutput = archive.CreateEntry(jpegPart).Open();
                        await jpegOutput.WriteAsync(image.Content, 0, image.Content.Length, token).ConfigureAwait(false);
                        relation.SetAttributeValue("Target", "../media/" + Path.GetFileName(jpegPart));
                        placeholders.Add(originalPart);
                    }
                    else pngParts.Add(originalPart);
                    anchor.ReplaceWith(new XElement(drawingNs + "oneCellAnchor",
                        new XElement(drawingNs + "from",
                            new XElement(drawingNs + "col", image.Column),
                            new XElement(drawingNs + "colOff", (long)image.OffsetX * 9525),
                            new XElement(drawingNs + "row", image.Row),
                            new XElement(drawingNs + "rowOff", (long)image.OffsetY * 9525)),
                        new XElement(drawingNs + "ext", new XAttribute("cx", (long)image.Width * 9525), new XAttribute("cy", (long)image.Height * 9525)),
                        anchor.Element(drawingNs + "pic"), new XElement(drawingNs + "clientData")));
                }
                using var output = entry.Open();
                output.SetLength(0);
                using var writer = XmlWriter.Create(output, new XmlWriterSettings { Async = true, Encoding = new System.Text.UTF8Encoding(false), CloseOutput = false });
                await document.SaveAsync(writer, token).ConfigureAwait(false);
                await writer.FlushAsync().ConfigureAwait(false);
                using var relOutput = drawingRelationsEntry.Open();
                relOutput.SetLength(0);
                using var relWriter = XmlWriter.Create(relOutput, new XmlWriterSettings { Async = true, Encoding = new System.Text.UTF8Encoding(false), CloseOutput = false });
                await drawingRelations.SaveAsync(relWriter, token).ConfigureAwait(false);
                await relWriter.FlushAsync().ConfigureAwait(false);
            }
            foreach (var part in placeholders.Except(pngParts)) archive.GetEntry(part)?.Delete();
            if (placeholders.Count > 0)
            {
                var typesEntry = archive.GetEntry("[Content_Types].xml");
                XDocument types;
                using (var input = typesEntry.Open()) types = XDocument.Load(input);
                if (!types.Root.Elements().Any(element => (string)element.Attribute("Extension") == "jpeg"))
                    types.Root.Add(new XElement(types.Root.Name.Namespace + "Default", new XAttribute("Extension", "jpeg"), new XAttribute("ContentType", "image/jpeg")));
                using var output = typesEntry.Open();
                output.SetLength(0);
                using var writer = XmlWriter.Create(output, new XmlWriterSettings { Async = true, Encoding = new System.Text.UTF8Encoding(false), CloseOutput = false });
                await types.SaveAsync(writer, token).ConfigureAwait(false);
                await writer.FlushAsync().ConfigureAwait(false);
            }
        }
        token.ThrowIfCancellationRequested();
        staged.Position = 0;
        await staged.CopyToAsync(destination, 81920, token).ConfigureAwait(false);
    }

    /// <summary>
    /// 异步嵌入工作簿请求中的图片资源。
    /// </summary>
    /// <param name="spreadsheet">当前输出工作簿。</param>
    /// <param name="request">工作簿或工作表导出请求。</param>
    /// <param name="token">取消操作的令牌。</param>
    /// <returns>图片定义与已嵌入图片资源的映射。</returns>
    internal static async Task<IReadOnlyDictionary<ExcelSheetImageDefinition, EmbeddedImage>> EmbedAsync(
        Spreadsheet spreadsheet, ExcelWorkbookExportRequest request, CancellationToken token)
    {
        var result = new Dictionary<ExcelSheetImageDefinition, EmbeddedImage>();
        foreach (var image in request.Sheets.SelectMany(sheet => sheet.Images))
        {
            token.ThrowIfCancellationRequested();
            using var source = new MemoryStream(image.Content[0] == 255 ? JpegPlaceholder : image.Content, writable: false);
            var embedded = await spreadsheet.EmbedImageAsync(source, token).ConfigureAwait(false);
            result.Add(image, embedded);
        }
        return result;
    }

    /// <summary>
    /// 向当前工作表添加图片和数据校验。
    /// </summary>
    /// <param name="spreadsheet">当前输出工作簿。</param>
    /// <param name="request">工作簿或工作表导出请求。</param>
    /// <param name="images">图片定义到已嵌入资源的映射。</param>
    /// <param name="token">取消操作的令牌。</param>
    internal static void Apply(Spreadsheet spreadsheet, ExcelSheetExportRequest request,
        IReadOnlyDictionary<ExcelSheetImageDefinition, EmbeddedImage> images, CancellationToken token)
    {
        foreach (var image in request.Images)
        {
            token.ThrowIfCancellationRequested();
            // 精确尺寸与偏移在暂存包的 drawing 修正阶段写入。
            spreadsheet.AddImage(ImageCanvas.Dimensions(ExcelSheetContent.Address(image.Row, image.Column), 1, 1), images[image]);
        }
        foreach (var rule in request.DataValidations)
        {
            token.ThrowIfCancellationRequested();
            var native = Create(rule);
            native.IgnoreBlank = rule.IgnoreBlanks;
            native.ShowErrorAlert = rule.ShowErrorMessage;
            native.ShowInputMessage = !string.IsNullOrEmpty(rule.InputTitle) || !string.IsNullOrEmpty(rule.InputMessage);
            native.InputTitle = rule.InputTitle;
            native.InputMessage = rule.InputMessage;
            native.ErrorTitle = rule.ErrorTitle;
            native.ErrorMessage = rule.ErrorMessage;
            spreadsheet.AddDataValidation(ExcelSheetContent.Address(rule.Range), native);
        }
    }

    /// <summary>
    /// 创建原生数据校验规则。
    /// </summary>
    /// <param name="rule">公共数据校验定义。</param>
    /// <returns>与公共校验条件对应的原生校验规则。</returns>
    private static DataValidation Create(ExcelDataValidationDefinition rule) => rule.Type switch
    {
            ExcelDataValidationType.Integer => rule.Operator switch
            {
                ExcelConditionalComparisonOperator.Equal => DataValidation.IntegerEqualTo((int)rule.Value1),
                ExcelConditionalComparisonOperator.NotEqual => DataValidation.IntegerNotEqualTo((int)rule.Value1),
                ExcelConditionalComparisonOperator.GreaterThan => DataValidation.IntegerGreaterThan((int)rule.Value1),
                ExcelConditionalComparisonOperator.GreaterThanOrEqual => DataValidation.IntegerGreaterThanOrEqualTo((int)rule.Value1),
                ExcelConditionalComparisonOperator.LessThan => DataValidation.IntegerLessThan((int)rule.Value1),
                ExcelConditionalComparisonOperator.LessThanOrEqual => DataValidation.IntegerLessThanOrEqualTo((int)rule.Value1),
                ExcelConditionalComparisonOperator.Between => DataValidation.IntegerBetween((int)rule.Value1, (int)rule.Value2),
                ExcelConditionalComparisonOperator.NotBetween => DataValidation.IntegerNotBetween((int)rule.Value1, (int)rule.Value2),
                _ => throw new ArgumentOutOfRangeException(nameof(rule.Operator))
            },
            ExcelDataValidationType.Decimal => rule.Operator switch
            {
                ExcelConditionalComparisonOperator.Equal => DataValidation.DecimalEqualTo(rule.Value1),
                ExcelConditionalComparisonOperator.NotEqual => DataValidation.DecimalNotEqualTo(rule.Value1),
                ExcelConditionalComparisonOperator.GreaterThan => DataValidation.DecimalGreaterThan(rule.Value1),
                ExcelConditionalComparisonOperator.GreaterThanOrEqual => DataValidation.DecimalGreaterThanOrEqualTo(rule.Value1),
                ExcelConditionalComparisonOperator.LessThan => DataValidation.DecimalLessThan(rule.Value1),
                ExcelConditionalComparisonOperator.LessThanOrEqual => DataValidation.DecimalLessThanOrEqualTo(rule.Value1),
                ExcelConditionalComparisonOperator.Between => DataValidation.DecimalBetween(rule.Value1, rule.Value2),
                ExcelConditionalComparisonOperator.NotBetween => DataValidation.DecimalNotBetween(rule.Value1, rule.Value2),
                _ => throw new ArgumentOutOfRangeException(nameof(rule.Operator))
            },
            ExcelDataValidationType.TextLength => rule.Operator switch
            {
                ExcelConditionalComparisonOperator.Equal => DataValidation.TextLengthEqualTo((int)rule.Value1),
                ExcelConditionalComparisonOperator.NotEqual => DataValidation.TextLengthNotEqualTo((int)rule.Value1),
                ExcelConditionalComparisonOperator.GreaterThan => DataValidation.TextLengthGreaterThan((int)rule.Value1),
                ExcelConditionalComparisonOperator.GreaterThanOrEqual => DataValidation.TextLengthGreaterThanOrEqualTo((int)rule.Value1),
                ExcelConditionalComparisonOperator.LessThan => DataValidation.TextLengthLessThan((int)rule.Value1),
                ExcelConditionalComparisonOperator.LessThanOrEqual => DataValidation.TextLengthLessThanOrEqualTo((int)rule.Value1),
                ExcelConditionalComparisonOperator.Between => DataValidation.TextLengthBetween((int)rule.Value1, (int)rule.Value2),
                ExcelConditionalComparisonOperator.NotBetween => DataValidation.TextLengthNotBetween((int)rule.Value1, (int)rule.Value2),
                _ => throw new ArgumentOutOfRangeException(nameof(rule.Operator))
            },
            ExcelDataValidationType.Date => rule.Operator switch
            {
                ExcelConditionalComparisonOperator.Equal => DataValidation.DateTimeEqualTo(rule.Date1.Value),
                ExcelConditionalComparisonOperator.NotEqual => DataValidation.DateTimeNotEqualTo(rule.Date1.Value),
                ExcelConditionalComparisonOperator.GreaterThan => DataValidation.DateTimeGreaterThan(rule.Date1.Value),
                ExcelConditionalComparisonOperator.GreaterThanOrEqual => DataValidation.DateTimeGreaterThanOrEqualTo(rule.Date1.Value),
                ExcelConditionalComparisonOperator.LessThan => DataValidation.DateTimeLessThan(rule.Date1.Value),
                ExcelConditionalComparisonOperator.LessThanOrEqual => DataValidation.DateTimeLessThanOrEqualTo(rule.Date1.Value),
                ExcelConditionalComparisonOperator.Between => DataValidation.DateTimeBetween(rule.Date1.Value, rule.Date2.GetValueOrDefault()),
                ExcelConditionalComparisonOperator.NotBetween => DataValidation.DateTimeNotBetween(rule.Date1.Value, rule.Date2.GetValueOrDefault()),
                _ => throw new ArgumentOutOfRangeException(nameof(rule.Operator))
            },
        ExcelDataValidationType.List => DataValidation.ListValues(rule.Values),
        _ => throw new ArgumentOutOfRangeException(nameof(rule.Type))
    };
}
