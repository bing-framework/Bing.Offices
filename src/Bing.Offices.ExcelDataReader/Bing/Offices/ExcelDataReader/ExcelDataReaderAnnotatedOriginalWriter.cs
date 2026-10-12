using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.ExcelDataReader;

/// <summary>
/// 在 XLSX 原工作簿副本中加入失败摘要与单元格批注。
/// </summary>
/// <remarks>
/// 原始 ZIP 部件默认逐项复制；仅工作簿索引、目标工作表关系和批注部件会被改写。
/// </remarks>
internal static class ExcelDataReaderAnnotatedOriginalWriter
{
    /// <summary>
    /// SpreadsheetML 命名空间。
    /// </summary>
    private static readonly XNamespace SpreadsheetNs =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// Office 文档关系命名空间。
    /// </summary>
    private static readonly XNamespace OfficeRelNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// OPC 包关系命名空间。
    /// </summary>
    private static readonly XNamespace PackageRelNs =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>
    /// OPC 内容类型命名空间。
    /// </summary>
    private static readonly XNamespace ContentTypeNs =
        "http://schemas.openxmlformats.org/package/2006/content-types";

    /// <summary>
    /// VML 批注绘图命名空间。
    /// </summary>
    private static readonly XNamespace VmlNs = "urn:schemas-microsoft-com:vml";

    /// <summary>
    /// VML Office 扩展命名空间。
    /// </summary>
    private static readonly XNamespace VmlOfficeNs = "urn:schemas-microsoft-com:office:office";

    /// <summary>
    /// VML Excel 客户端数据命名空间。
    /// </summary>
    private static readonly XNamespace VmlExcelNs = "urn:schemas-microsoft-com:office:excel";

    /// <summary>
    /// 关系类型 URI 前缀。
    /// </summary>
    private const string OfficeRelationshipBase =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    /// <summary>
    /// 失败摘要工作表基础名称。
    /// </summary>
    private const string SummarySheetBaseName = "__BingErrors";

    /// <summary>
    /// 失败摘要工作表列标题。
    /// </summary>
    private static readonly string[] SummaryHeaders =
    {
        "Sheet", "Row", "Column", "Code", "Message", "Property", "ColumnKey", "Header", "RawValue",
        "FirstRow"
    };

    /// <summary>
    /// 在原始 XLSX 包副本中加入错误摘要和可定位单元格批注。
    /// </summary>
    /// <param name="stagedXlsx">可读取、可查找的 staged XLSX 流；调用方仍拥有该流。</param>
    /// <param name="options">失败工作簿策略与资源预算。</param>
    /// <param name="errors">导入期间产生的结构化错误。</param>
    /// <param name="capture">前向读取捕获；本模式使用其中已登记的源工作表信息。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>包含原有部件及失败信息的 XLSX 内存流，位置位于起始处。</returns>
    internal static MemoryStream Serialize(Stream stagedXlsx, ExcelImportFailureOptions options,
        IReadOnlyList<ExcelImportError> errors, ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken)
    {
        if (stagedXlsx == null)
            throw new ArgumentNullException(nameof(stagedXlsx));
        if (!stagedXlsx.CanRead)
            throw new ArgumentException("源工作簿流不可读取。", nameof(stagedXlsx));
        if (!stagedXlsx.CanSeek)
            throw new ArgumentException("AnnotatedOriginal 要求可查找的 staged XLSX 流。", nameof(stagedXlsx));
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        if (capture == null)
            throw new ArgumentNullException(nameof(capture));
        if (options.Mode != ExcelImportFailureWorkbookMode.AnnotatedOriginal)
            throw new ArgumentException("AnnotatedOriginal 写入器仅支持 AnnotatedOriginal 模式。",
                nameof(options));
        options.Validate();
        cancellationToken.ThrowIfCancellationRequested();

        var sourcePosition = stagedXlsx.Position;
        try
        {
            stagedXlsx.Position = 0;
            using var source = OpenPackage(stagedXlsx, options, cancellationToken);
            var package = ReadPackage(source, options, errors ?? Array.Empty<ExcelImportError>(), capture,
                cancellationToken);
            ApplyAnnotations(package, options, errors ?? Array.Empty<ExcelImportError>(), cancellationToken);
            AddSummarySheet(package, errors ?? Array.Empty<ExcelImportError>(), cancellationToken);
            ValidateBudgets(package, options, errors ?? Array.Empty<ExcelImportError>(), cancellationToken);
            return WritePackage(source, package, options, cancellationToken);
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw InvalidPackage("AnnotatedOriginal 无法更新输入 XLSX 包。", exception);
        }
        finally
        {
            stagedXlsx.Position = sourcePosition;
        }
    }

    /// <summary>
    /// 将 AnnotatedOriginal 结果序列化为 XLSX 字节数组。
    /// </summary>
    /// <param name="stagedXlsx">可读取、可查找的 staged XLSX 流；调用方仍拥有该流。</param>
    /// <param name="options">失败工作簿策略与资源预算。</param>
    /// <param name="errors">导入期间产生的结构化错误。</param>
    /// <param name="capture">前向读取捕获。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>完整 XLSX 包的字节数组。</returns>
    internal static byte[] SerializeToBytes(Stream stagedXlsx, ExcelImportFailureOptions options,
        IReadOnlyList<ExcelImportError> errors, ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken)
    {
        using var result = Serialize(stagedXlsx, options, errors, capture, cancellationToken);
        return result.ToArray();
    }

    /// <summary>
    /// 以安全读取设置打开一个 ZIP 包并验证 XLSX 必需部件。
    /// </summary>
    /// <param name="source">源工作簿流。</param>
    /// <param name="options">资源限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>保持源流打开的 XLSX ZIP 包。</returns>
    private static ZipArchive OpenPackage(Stream source, ExcelImportFailureOptions options,
        CancellationToken cancellationToken)
    {
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in archive.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ValidatePartName(entry.FullName);
                if (!names.Add(entry.FullName))
                {
                    archive.Dispose();
                    throw InvalidPackage("XLSX 包包含重复部件名称。");
                }
                if (options.MaxSerializedBytes.HasValue && entry.Length > options.MaxSerializedBytes.Value)
                {
                    archive.Dispose();
                    throw ResourceLimit("源部件超过 AnnotatedOriginal 序列化预算。");
                }
            }
            if (!names.Contains("[Content_Types].xml") || !names.Contains("_rels/.rels"))
            {
                archive.Dispose();
                throw InvalidPackage("输入不是包含 OPC 内容类型和根关系部件的 XLSX 包。");
            }
            return archive;
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is InvalidDataException || exception is IOException
            || exception is NotSupportedException)
        {
            throw InvalidPackage("输入不是可读取的 XLSX ZIP 包。", exception);
        }
    }

    /// <summary>
    /// 解析工作簿清单并统计源单元格和行对象。
    /// </summary>
    /// <param name="archive">已验证的源 ZIP 包。</param>
    /// <param name="options">资源预算。</param>
    /// <param name="errors">结构化错误。</param>
    /// <param name="capture">捕获的源工作表信息。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>待更新的工作簿包模型。</returns>
    private static PackageState ReadPackage(ZipArchive archive, ExcelImportFailureOptions options,
        IReadOnlyList<ExcelImportError> errors, ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken)
    {
        var entries = archive.Entries.ToDictionary(entry => entry.FullName,
            StringComparer.OrdinalIgnoreCase);
        var contentTypes = LoadXml(RequiredEntry(entries, "[Content_Types].xml"), options,
            cancellationToken, "内容类型部件");
        ValidateRelationshipParts(entries, options, cancellationToken);
        var rootRelationships = LoadXml(RequiredEntry(entries, "_rels/.rels"), options,
            cancellationToken, "根关系部件");
        var officeRelationship = rootRelationships.Root?.Elements(PackageRelNs + "Relationship")
            .Where(element => (string)element.Attribute("Type") == OfficeRelationshipBase + "officeDocument")
            .ToArray() ?? Array.Empty<XElement>();
        if (officeRelationship.Length != 1)
            throw InvalidPackage("XLSX 包缺少唯一的工作簿根关系。");
        var workbookPart = ResolveTarget("", RequiredInternalTarget(officeRelationship[0]));
        var workbookEntry = RequiredEntry(entries, workbookPart);
        var workbook = LoadXml(workbookEntry, options, cancellationToken, "工作簿部件");
        var workbookContentType = contentTypes.Root?.Elements(ContentTypeNs + "Override")
            .SingleOrDefault(element => string.Equals((string)element.Attribute("PartName"),
                "/" + workbookPart, StringComparison.OrdinalIgnoreCase));
        if (!string.Equals((string)workbookContentType?.Attribute("ContentType"),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                StringComparison.OrdinalIgnoreCase))
            throw InvalidPackage("AnnotatedOriginal 仅支持普通 XLSX 工作簿，不支持 XLS/XLSB 或宏工作簿。");
        var workbookRelationshipsPart = RelationshipPart(workbookPart);
        var workbookRelationships = LoadXml(RequiredEntry(entries, workbookRelationshipsPart), options,
            cancellationToken, "工作簿关系部件");
        var relationships = ParseRelationships(workbookRelationships, workbookPart);
        var sheets = ParseSheets(workbook, relationships, entries, options, cancellationToken);
        var cellCount = 0L;
        var rowCount = 0L;
        foreach (var sheet in sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            CountWorksheetCells(RequiredEntry(entries, sheet.Part), options, cancellationToken,
                ref cellCount, ref rowCount);
        }
        var pictureCount = 0;
        var pictureBytes = 0L;
        foreach (var entry in entries.Values)
        {
            if (!IsPicturePart(entry.FullName))
                continue;
            cancellationToken.ThrowIfCancellationRequested();
            pictureCount = checked(pictureCount + 1);
            pictureBytes = checked(pictureBytes + entry.Length);
            if (options.MaxCopiedPictures.HasValue
                && pictureCount > options.MaxCopiedPictures.Value)
                throw ResourceLimit($"AnnotatedOriginal 图片数量超过限制: {options.MaxCopiedPictures.Value}");
            if (options.MaxCopiedPictureBytes.HasValue
                && pictureBytes > options.MaxCopiedPictureBytes.Value)
                throw ResourceLimit($"AnnotatedOriginal 图片字节数超过限制: {options.MaxCopiedPictureBytes.Value}");
        }
        var candidateRows = errors.Where(error => error != null && error.RowIndex > 0
                && error.ColumnIndex > 0 && !string.IsNullOrWhiteSpace(error.SheetName))
            .Select(error => (error.SheetName, error.RowIndex))
            .Distinct(new SheetRowComparer()).LongCount();
        if (options.MaxCandidateErrorRows.HasValue && candidateRows > options.MaxCandidateErrorRows.Value)
            throw ResourceLimit($"AnnotatedOriginal 错误候选行数超过限制: {options.MaxCandidateErrorRows.Value}");
        if (options.MaxCopiedCells.HasValue && cellCount > options.MaxCopiedCells.Value)
            throw ResourceLimit($"AnnotatedOriginal 保留的源单元格数超过限制: {options.MaxCopiedCells.Value}");
        var errorSheetName = UniqueSheetName(workbook, SummarySheetBaseName);
        return new PackageState(entries, contentTypes, workbook, workbookRelationships, workbookPart,
            workbookRelationshipsPart, sheets, errorSheetName, cellCount, rowCount, capture,
            pictureCount, options, cancellationToken);
    }

    /// <summary>
    /// 判断部件是否为需要保留的工作簿图片资源。
    /// </summary>
    /// <param name="part">ZIP 部件路径。</param>
    /// <returns>部件位于标准图片目录时返回 <see langword="true" />。</returns>
    private static bool IsPicturePart(string part) =>
        part.StartsWith("xl/media/", StringComparison.OrdinalIgnoreCase)
        && !part.EndsWith("/", StringComparison.Ordinal);

    /// <summary>
    /// 将错误坐标转换为原工作表批注，并按冲突策略处理已有批注。
    /// </summary>
    /// <param name="package">待更新的工作簿包模型。</param>
    /// <param name="options">批注冲突策略和资源限制。</param>
    /// <param name="errors">结构化错误。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void ApplyAnnotations(PackageState package, ExcelImportFailureOptions options,
        IReadOnlyList<ExcelImportError> errors, CancellationToken cancellationToken)
    {
        var located = errors.Where(error => error != null && error.RowIndex > 0
                && error.RowIndex <= 1_048_576 && error.ColumnIndex > 0
                && error.ColumnIndex <= 16_384 && !string.IsNullOrWhiteSpace(error.SheetName))
            .GroupBy(error => (error.SheetName, error.RowIndex, error.ColumnIndex), new ErrorCellComparer())
            .ToArray();
        foreach (var group in located)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = package.Sheets.FirstOrDefault(item =>
                string.Equals(item.Name, group.Key.SheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
                continue;
            var text = string.Join(Environment.NewLine, group.Select(error =>
                $"{error.Code}: {error.Message}"));
            AddCellComment(package, sheet, CellReference(group.Key.ColumnIndex, group.Key.RowIndex), text,
                options.CommentConflictPolicy, options, cancellationToken);
        }
    }

    /// <summary>
    /// 为一个单元格追加或更新批注，并维护工作表关系和 VML 绘图。
    /// </summary>
    /// <param name="package">待更新的工作簿包模型。</param>
    /// <param name="sheet">错误坐标所属工作表。</param>
    /// <param name="cellReference">目标 A1 单元格坐标。</param>
    /// <param name="failureText">待写入的错误说明。</param>
    /// <param name="policy">已有批注冲突策略。</param>
    /// <param name="options">输出资源预算。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void AddCellComment(PackageState package, SheetState sheet, string cellReference,
        string failureText, ExcelImportCommentConflictPolicy policy, ExcelImportFailureOptions options,
        CancellationToken cancellationToken)
    {
        var sheetRelationshipsPart = RelationshipPart(sheet.Part);
        var sheetRelationships = package.GetOrCreateRelationships(sheetRelationshipsPart);
        var commentRelationship = FindRelationship(sheetRelationships, "comments");
        var vmlRelationship = FindRelationship(sheetRelationships, "vmlDrawing");
        var commentsPart = commentRelationship == null
            ? package.UniquePart("xl/comments", ".xml")
            : ResolveTarget(sheet.Part, RequiredInternalTarget(commentRelationship));
        var vmlPart = vmlRelationship == null
            ? package.UniquePart("xl/drawings/commentsDrawing", ".vml")
            : ResolveTarget(sheet.Part, RequiredInternalTarget(vmlRelationship));
        var commentsEntry = package.Entries.TryGetValue(commentsPart, out var existingCommentsEntry)
            ? existingCommentsEntry : null;
        if (commentRelationship != null && commentsEntry == null)
            throw InvalidPackage("工作表批注关系指向不存在的部件。");
        var comments = commentsEntry == null
            ? NewCommentsDocument()
            : LoadXml(commentsEntry, options, cancellationToken, "批注部件");
        var root = comments.Root;
        if (root?.Name != SpreadsheetNs + "comments")
            throw InvalidPackage("批注部件的根元素无效。");
        var authors = root.Element(SpreadsheetNs + "authors");
        if (authors == null)
        {
            authors = new XElement(SpreadsheetNs + "authors");
            root.AddFirst(authors);
        }
        var authorNames = authors.Elements(SpreadsheetNs + "author").ToList();
        var authorIndex = authorNames.FindIndex(author => string.Equals(author.Value, "Bing.Offices",
            StringComparison.Ordinal));
        if (authorIndex < 0)
        {
            authorIndex = authorNames.Count;
            authors.Add(new XElement(SpreadsheetNs + "author", "Bing.Offices"));
        }
        var commentList = root.Element(SpreadsheetNs + "commentList");
        if (commentList == null)
        {
            commentList = new XElement(SpreadsheetNs + "commentList");
            root.Add(commentList);
        }
        var existing = commentList.Elements(SpreadsheetNs + "comment").FirstOrDefault(comment =>
            string.Equals((string)comment.Attribute("ref"), cellReference, StringComparison.OrdinalIgnoreCase));
        if (existing != null && policy == ExcelImportCommentConflictPolicy.Fail)
            throw new BingOfficesConfigurationException(
                $"目标单元格已有批注: {sheet.Name}!{cellReference}", stage: BingOfficesStage.Plan);
        if (existing != null && policy == ExcelImportCommentConflictPolicy.Preserve)
            return;
        var combinedText = existing != null && policy == ExcelImportCommentConflictPolicy.Append
            ? string.Join(Environment.NewLine, new[] { CommentText(existing), failureText }
                .Where(value => !string.IsNullOrEmpty(value)))
            : failureText;
        existing?.Remove();
        commentList.Add(CreateComment(cellReference, authorIndex, combinedText));
        var existingVmlEntry = package.Entries.TryGetValue(vmlPart, out var vmlEntry) ? vmlEntry : null;
        var vml = existingVmlEntry == null ? NewVmlDocument() : LoadXml(existingVmlEntry, options,
            cancellationToken, "批注 VML 部件");
        EnsureOnlyNoteShapes(vml);
        var commentsByCell = commentList.Elements(SpreadsheetNs + "comment")
            .Select(comment => (Reference: (string)comment.Attribute("ref"), Comment: comment))
            .Where(item => !string.IsNullOrWhiteSpace(item.Reference))
            .ToDictionary(item => item.Reference, item => item.Comment, StringComparer.OrdinalIgnoreCase);
        ReplaceNoteShapes(vml, commentsByCell.Keys, cancellationToken);
        sheet.Document ??= LoadXml(RequiredEntry(package.Entries, sheet.Part), package.Options,
            cancellationToken, "工作表部件");
        package.ReplaceXml(commentsPart, comments);
        package.ReplaceXml(vmlPart, vml);
        EnsureRelationship(sheetRelationships, commentRelationship, "comments", commentsPart, sheet.Part);
        EnsureRelationship(sheetRelationships, vmlRelationship, "vmlDrawing", vmlPart, sheet.Part);
        package.ReplaceXml(sheetRelationshipsPart, sheetRelationships);
        EnsureCommentContentTypes(package, commentsPart);
        AddLegacyDrawing(sheet, vmlRelationship?.Attribute("Id")?.Value
            ?? FindRelationship(sheetRelationships, "vmlDrawing")?.Attribute("Id")?.Value);
        package.ReplaceXml(sheet.Part, sheet.Document);
        sheet.CommentCount = commentsByCell.Count;
    }

    /// <summary>
    /// 在原工作簿中追加保存全部结构化错误的摘要工作表。
    /// </summary>
    /// <param name="package">待更新的工作簿包模型。</param>
    /// <param name="errors">结构化错误。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void AddSummarySheet(PackageState package, IReadOnlyList<ExcelImportError> errors,
        CancellationToken cancellationToken)
    {
        var sheets = package.Workbook.Root?.Element(SpreadsheetNs + "sheets");
        if (sheets == null)
            throw InvalidPackage("工作簿部件缺少 sheets 元素。");
        var relationshipId = NextRelationshipId(package.WorkbookRelationships);
        var sheetId = package.Sheets.Count == 0 ? 1 : package.Sheets.Max(sheet => sheet.SheetId) + 1;
        var summaryPart = package.UniquePart("xl/worksheets/bingErrors", ".xml");
        var summarySheet = CreateSummaryWorksheet(errors, cancellationToken);
        sheets.Add(new XElement(SpreadsheetNs + "sheet",
            new XAttribute("name", package.ErrorSheetName),
            new XAttribute("sheetId", sheetId.ToString(CultureInfo.InvariantCulture)),
            new XAttribute(OfficeRelNs + "id", relationshipId)));
        package.WorkbookRelationships.Root.Add(new XElement(PackageRelNs + "Relationship",
            new XAttribute("Id", relationshipId),
            new XAttribute("Type", OfficeRelationshipBase + "worksheet"),
            new XAttribute("Target", RelativeTarget(package.WorkbookPart, summaryPart))));
        package.ReplaceXml(package.WorkbookPart, package.Workbook);
        package.ReplaceXml(package.WorkbookRelationshipsPart, package.WorkbookRelationships);
        package.ReplaceXml(summaryPart, summarySheet);
        EnsureWorksheetContentType(package, summaryPart);
    }

    /// <summary>
    /// 验证候选行、复制单元格和目标对象预算。
    /// </summary>
    /// <param name="package">已解析的包模型。</param>
    /// <param name="options">资源预算。</param>
    /// <param name="errors">结构化错误。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void ValidateBudgets(PackageState package, ExcelImportFailureOptions options,
        IReadOnlyList<ExcelImportError> errors, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var summaryCells = checked((long)SummaryHeaders.Length * (errors.Count + 1L));
        var addedCells = checked(package.CellCount + summaryCells);
        if (options.MaxCopiedCells.HasValue && addedCells > options.MaxCopiedCells.Value)
            throw ResourceLimit($"AnnotatedOriginal 单元格估算超过限制: {options.MaxCopiedCells.Value}");
        var objects = checked(package.RowCount + package.CellCount + package.Sheets.Count + 1L
            + errors.Count + summaryCells + package.PictureCount
            + package.Sheets.Sum(sheet => sheet.CommentCount));
        if (options.MaxEstimatedTargetObjects.HasValue && objects > options.MaxEstimatedTargetObjects.Value)
            throw ResourceLimit($"AnnotatedOriginal 目标对象估算超过限制: {options.MaxEstimatedTargetObjects.Value}");
    }

    /// <summary>
    /// 写入保留原部件并应用更新的 XLSX 包。
    /// </summary>
    /// <param name="source">原始 ZIP 包。</param>
    /// <param name="package">更新后的包模型。</param>
    /// <param name="options">输出资源预算。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>完整 XLSX 内存流。</returns>
    private static MemoryStream WritePackage(ZipArchive source, PackageState package,
        ExcelImportFailureOptions options, CancellationToken cancellationToken)
    {
        var output = new MemoryStream();
        try
        {
            using (var limited = new LimitedWriteStream(output, options.MaxSerializedBytes))
            using (var destination = new ZipArchive(limited, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (var entry in source.Entries)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (package.Replacements.TryGetValue(entry.FullName, out var replacement))
                        WriteEntry(destination, entry.FullName, replacement, cancellationToken);
                    else
                        CopyEntry(entry, destination, cancellationToken);
                }
                foreach (var entry in package.Additions)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (package.Entries.ContainsKey(entry.Key))
                        throw InvalidPackage($"无法追加已存在的 XLSX 部件: {entry.Key}");
                    WriteEntry(destination, entry.Key, entry.Value, cancellationToken);
                }
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
    /// 逐块复制 ZIP 部件并响应取消。
    /// </summary>
    /// <param name="source">源部件。</param>
    /// <param name="destination">目标 ZIP 包。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void CopyEntry(ZipArchiveEntry source, ZipArchive destination,
        CancellationToken cancellationToken)
    {
        var target = destination.CreateEntry(source.FullName, CompressionLevel.Optimal);
        using var input = source.Open();
        using var output = target.Open();
        Copy(input, output, cancellationToken);
    }

    /// <summary>
    /// 将替换后的 XML 部件写入 ZIP 包。
    /// </summary>
    /// <param name="archive">目标 ZIP 包。</param>
    /// <param name="part">部件名称。</param>
    /// <param name="content">部件内容。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void WriteEntry(ZipArchive archive, string part, byte[] content,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var entry = archive.CreateEntry(part, CompressionLevel.Optimal);
        using var output = entry.Open();
        output.Write(content, 0, content.Length);
    }

    /// <summary>
    /// 将流分块复制到目标流并检查取消。
    /// </summary>
    /// <param name="source">读取流。</param>
    /// <param name="destination">写入流。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void Copy(Stream source, Stream destination, CancellationToken cancellationToken)
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
    /// 解析关系部件并校验关系标识唯一。
    /// </summary>
    /// <param name="document">关系部件 XML。</param>
    /// <param name="sourcePart">关系所属源部件。</param>
    /// <returns>关系标识到 XML 节点的映射。</returns>
    private static Dictionary<string, XElement> ParseRelationships(XDocument document, string sourcePart)
    {
        if (document.Root?.Name != PackageRelNs + "Relationships")
            throw InvalidPackage($"关系部件根元素无效: {sourcePart}");
        var result = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var relation in document.Root.Elements(PackageRelNs + "Relationship"))
        {
            var id = (string)relation.Attribute("Id");
            if (string.IsNullOrWhiteSpace(id) || !result.TryAdd(id, relation))
                throw InvalidPackage($"关系部件包含空或重复的关系标识: {sourcePart}");
        }
        return result;
    }

    /// <summary>
    /// 验证包内所有关系均为唯一、内部且指向现有部件。
    /// </summary>
    /// <param name="entries">已验证的 ZIP 部件映射。</param>
    /// <param name="options">资源限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void ValidateRelationshipParts(
        IReadOnlyDictionary<string, ZipArchiveEntry> entries,
        ExcelImportFailureOptions options, CancellationToken cancellationToken)
    {
        foreach (var entry in entries.Values.Where(item => IsRelationshipPart(item.FullName)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sourcePart = SourcePartFromRelationshipPart(entry.FullName);
            var document = LoadXml(entry, options, cancellationToken, "关系部件");
            var relationships = ParseRelationships(document, entry.FullName);
            foreach (var relationship in relationships.Values)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var target = ResolveTarget(sourcePart, RequiredInternalTarget(relationship));
                if (!entries.ContainsKey(target))
                    throw InvalidPackage($"关系目标部件不存在: {entry.FullName} -> {target}");
            }
        }
    }

    /// <summary>
    /// 判断部件是否为 OPC 关系部件。
    /// </summary>
    /// <param name="part">ZIP 部件路径。</param>
    /// <returns>关系部件返回 <see langword="true" />。</returns>
    private static bool IsRelationshipPart(string part) =>
        string.Equals(part, "_rels/.rels", StringComparison.OrdinalIgnoreCase)
        || (part.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
            && part.Contains("/_rels/", StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// 根据关系部件路径取得其源部件路径。
    /// </summary>
    /// <param name="relationshipPart">关系部件路径。</param>
    /// <returns>关系源部件路径。</returns>
    private static string SourcePartFromRelationshipPart(string relationshipPart)
    {
        if (string.Equals(relationshipPart, "_rels/.rels", StringComparison.OrdinalIgnoreCase))
            return string.Empty;
        const string marker = "/_rels/";
        var markerIndex = relationshipPart.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (markerIndex <= 0 || !relationshipPart.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            throw InvalidPackage($"关系部件路径无效: {relationshipPart}");
        var directory = relationshipPart.Substring(0, markerIndex + 1);
        var fileNameStart = markerIndex + marker.Length;
        var fileName = relationshipPart.Substring(fileNameStart,
            relationshipPart.Length - fileNameStart - ".rels".Length);
        if (string.IsNullOrWhiteSpace(fileName))
            throw InvalidPackage($"关系部件缺少源部件名称: {relationshipPart}");
        return directory + fileName;
    }

    /// <summary>
    /// 解析工作表清单并确认每个关系都指向存在的工作表部件。
    /// </summary>
    /// <param name="workbook">工作簿 XML。</param>
    /// <param name="relationships">工作簿关系映射。</param>
    /// <param name="entries">ZIP 部件映射。</param>
    /// <param name="options">资源限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>按工作簿顺序排列的工作表。</returns>
    private static List<SheetState> ParseSheets(XDocument workbook,
        IReadOnlyDictionary<string, XElement> relationships,
        IReadOnlyDictionary<string, ZipArchiveEntry> entries, ExcelImportFailureOptions options,
        CancellationToken cancellationToken)
    {
        var sheetsElement = workbook.Root?.Element(SpreadsheetNs + "sheets");
        if (sheetsElement == null)
            throw InvalidPackage("工作簿部件缺少 sheets 元素。");
        var result = new List<SheetState>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in sheetsElement.Elements(SpreadsheetNs + "sheet"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var name = (string)sheet.Attribute("name");
            var relationshipId = (string)sheet.Attribute(OfficeRelNs + "id");
            if (string.IsNullOrWhiteSpace(name) || !names.Add(name)
                || string.IsNullOrWhiteSpace(relationshipId)
                || !relationships.TryGetValue(relationshipId, out var relationship)
                || (string)relationship.Attribute("Type") != OfficeRelationshipBase + "worksheet")
                throw InvalidPackage("工作簿包含无效或重复的工作表声明。");
            var part = ResolveTarget("xl/workbook.xml", RequiredInternalTarget(relationship));
            if (!entries.ContainsKey(part))
                throw InvalidPackage($"工作表关系目标不存在: {part}");
            var sheetId = ParsePositiveInt((string)sheet.Attribute("sheetId"), "工作表 sheetId");
            result.Add(new SheetState(name, part, sheetId));
        }
        if (result.Count == 0)
            throw InvalidPackage("XLSX 工作簿没有可用工作表。");
        return result;
    }

    /// <summary>
    /// 流式统计工作表中的行和单元格节点。
    /// </summary>
    /// <param name="entry">工作表 ZIP 部件。</param>
    /// <param name="options">资源限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <param name="cellCount">累计单元格数量。</param>
    /// <param name="rowCount">累计行数量。</param>
    private static void CountWorksheetCells(ZipArchiveEntry entry, ExcelImportFailureOptions options,
        CancellationToken cancellationToken, ref long cellCount, ref long rowCount)
    {
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, SecureReaderSettings(options));
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element)
                continue;
            if (reader.LocalName == "row")
                rowCount = AddWithinLimit(rowCount, 1, options.MaxEstimatedTargetObjects,
                    "AnnotatedOriginal 工作表行数超过对象预算。");
            else if (reader.LocalName == "c")
                cellCount = AddWithinLimit(cellCount, 1, options.MaxCopiedCells,
                    "AnnotatedOriginal 源单元格数量超过复制预算。");
        }
    }

    /// <summary>
    /// 以禁止 DTD 和外部解析器的设置读取 XML 部件。
    /// </summary>
    /// <param name="entry">XML ZIP 部件。</param>
    /// <param name="options">资源限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <param name="description">用于结构化错误消息的部件说明。</param>
    /// <returns>解析后的 XML 文档。</returns>
    private static XDocument LoadXml(ZipArchiveEntry entry, ExcelImportFailureOptions options,
        CancellationToken cancellationToken, string description)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (options.MaxSerializedBytes.HasValue && entry.Length > options.MaxSerializedBytes.Value)
            throw ResourceLimit($"{description}超过 AnnotatedOriginal 字节预算。");
        try
        {
            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, SecureReaderSettings(options));
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is XmlException || exception is IOException
            || exception is InvalidDataException)
        {
            throw InvalidPackage($"{description}不是有效 XML。", exception);
        }
    }

    /// <summary>
    /// 创建限制 DTD 和 XML 实体扩展的读取设置。
    /// </summary>
    /// <param name="options">工作簿资源限制。</param>
    /// <returns>安全 XML 读取设置。</returns>
    private static XmlReaderSettings SecureReaderSettings(ExcelImportFailureOptions options) => new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = false,
        IgnoreWhitespace = false,
        MaxCharactersInDocument = options.MaxSerializedBytes ?? 0
    };

    /// <summary>
    /// 获取必须存在的 ZIP 部件。
    /// </summary>
    /// <param name="entries">ZIP 部件映射。</param>
    /// <param name="part">部件名称。</param>
    /// <returns>目标部件。</returns>
    private static ZipArchiveEntry RequiredEntry(IReadOnlyDictionary<string, ZipArchiveEntry> entries,
        string part) => entries.TryGetValue(part, out var entry)
        ? entry : throw InvalidPackage($"XLSX 部件不存在: {part}");

    /// <summary>
    /// 获取内部 OPC 关系的目标路径并拒绝外部关系。
    /// </summary>
    /// <param name="relationship">关系 XML 节点。</param>
    /// <returns>相对或根相对目标路径。</returns>
    private static string RequiredInternalTarget(XElement relationship)
    {
        if (string.Equals((string)relationship.Attribute("TargetMode"), "External",
            StringComparison.OrdinalIgnoreCase))
            throw InvalidPackage("AnnotatedOriginal 不接受指向外部部件的关系。");
        var target = (string)relationship.Attribute("Target");
        if (string.IsNullOrWhiteSpace(target) || target.Contains('?') || target.Contains('#'))
            throw InvalidPackage("XLSX 关系目标为空或包含不支持的 URI 后缀。");
        return target;
    }

    /// <summary>
    /// 将 OPC 关系目标解析为包内标准路径并拒绝越界路径。
    /// </summary>
    /// <param name="sourcePart">关系源部件；根关系传入空字符串。</param>
    /// <param name="target">关系目标。</param>
    /// <returns>无前导斜杠的包内部件路径。</returns>
    private static string ResolveTarget(string sourcePart, string target)
    {
        var decoded = Uri.UnescapeDataString(target.Replace('\\', '/'));
        var parts = (decoded.StartsWith('/') ? decoded.TrimStart('/')
            : (string.IsNullOrEmpty(sourcePart) ? decoded
                : GetDirectory(sourcePart) + decoded)).Split('/');
        var stack = new List<string>();
        foreach (var segment in parts)
        {
            if (segment.Length == 0 || segment == ".")
                continue;
            if (segment == "..")
            {
                if (stack.Count == 0)
                    throw InvalidPackage("XLSX 关系目标越出包根目录。");
                stack.RemoveAt(stack.Count - 1);
                continue;
            }
            stack.Add(segment);
        }
        if (stack.Count == 0)
            throw InvalidPackage("XLSX 关系目标为空。");
        return string.Join("/", stack);
    }

    /// <summary>
    /// 生成源部件对应的 OPC 关系部件名称。
    /// </summary>
    /// <param name="part">源部件路径。</param>
    /// <returns>关系部件路径。</returns>
    private static string RelationshipPart(string part)
    {
        var slash = part.LastIndexOf('/');
        var directory = slash < 0 ? string.Empty : part.Substring(0, slash + 1);
        var fileName = slash < 0 ? part : part.Substring(slash + 1);
        return directory + "_rels/" + fileName + ".rels";
    }

    /// <summary>
    /// 取得部件所在目录路径并保留末尾斜杠。
    /// </summary>
    /// <param name="part">部件路径。</param>
    /// <returns>目录路径。</returns>
    private static string GetDirectory(string part)
    {
        var slash = part.LastIndexOf('/');
        return slash < 0 ? string.Empty : part.Substring(0, slash + 1);
    }

    /// <summary>
    /// 生成相对于关系源部件目录的 OPC 目标路径。
    /// </summary>
    /// <param name="sourcePart">关系源部件路径。</param>
    /// <param name="targetPart">目标部件路径。</param>
    /// <returns>相对目标路径。</returns>
    private static string RelativeTarget(string sourcePart, string targetPart)
    {
        var sourceDirectory = GetDirectory(sourcePart).TrimEnd('/').Split('/',
            StringSplitOptions.RemoveEmptyEntries);
        var target = targetPart.Split('/', StringSplitOptions.RemoveEmptyEntries);
        var shared = 0;
        while (shared < sourceDirectory.Length && shared < target.Length
            && string.Equals(sourceDirectory[shared], target[shared], StringComparison.OrdinalIgnoreCase))
            shared++;
        return string.Join("/", Enumerable.Repeat("..", sourceDirectory.Length - shared)
            .Concat(target.Skip(shared)));
    }

    /// <summary>
    /// 验证 ZIP 部件路径不含根路径或目录穿越。
    /// </summary>
    /// <param name="part">ZIP 部件路径。</param>
    private static void ValidatePartName(string part)
    {
        if (string.IsNullOrWhiteSpace(part) || part.StartsWith('/') || part.Contains('\\')
            || part.Split('/').Any(segment => segment == ".."))
            throw InvalidPackage("XLSX 包含无效的部件路径。");
    }

    /// <summary>
    /// 解析必须为正数的 XML 整数属性。
    /// </summary>
    /// <param name="value">属性值。</param>
    /// <param name="description">错误说明。</param>
    /// <returns>正整数。</returns>
    private static int ParsePositiveInt(string value, string description)
    {
        if (int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed)
            && parsed > 0)
            return parsed;
        throw InvalidPackage($"{description}无效。");
    }

    /// <summary>
    /// 在累加期间检查对象预算并避免整数溢出。
    /// </summary>
    /// <param name="current">当前数量。</param>
    /// <param name="increment">增加数量。</param>
    /// <param name="maximum">最大数量。</param>
    /// <param name="message">超限消息。</param>
    /// <returns>累加后的数量。</returns>
    private static long AddWithinLimit(long current, long increment, long? maximum, string message)
    {
        long total;
        try
        {
            total = checked(current + increment);
        }
        catch (OverflowException exception)
        {
            throw ResourceLimit(message, exception);
        }
        if (maximum.HasValue && total > maximum.Value)
            throw ResourceLimit(message);
        return total;
    }

    /// <summary>
    /// 创建包含错误摘要行的工作表 XML。
    /// </summary>
    /// <param name="errors">结构化错误。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>错误摘要工作表文档。</returns>
    private static XDocument CreateSummaryWorksheet(IReadOnlyList<ExcelImportError> errors,
        CancellationToken cancellationToken)
    {
        var data = new XElement(SpreadsheetNs + "sheetData");
        data.Add(CreateSummaryRow(1, SummaryHeaders.Cast<object>(), cancellationToken));
        for (var index = 0; index < errors.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var error = errors[index];
            data.Add(CreateSummaryRow(index + 2, new object[]
            {
                error?.SheetName, error?.RowIndex, error?.ColumnIndex, error?.Code.ToString(), error?.Message,
                error?.PropertyName, error?.ColumnKey, error?.Header, FormatSummaryValue(error?.RawValue),
                error?.FirstRowNumber
            }, cancellationToken));
        }
        var lastRow = Math.Max(1, errors.Count + 1);
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(SpreadsheetNs + "worksheet",
                new XElement(SpreadsheetNs + "dimension",
                    new XAttribute("ref", $"A1:J{lastRow.ToString(CultureInfo.InvariantCulture)}")),
                data));
    }

    /// <summary>
    /// 创建一行只含内联文本的错误摘要单元格。
    /// </summary>
    /// <param name="rowIndex">一基输出行号。</param>
    /// <param name="values">行单元格值。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>工作表行 XML。</returns>
    private static XElement CreateSummaryRow(int rowIndex, IEnumerable<object> values,
        CancellationToken cancellationToken)
    {
        var row = new XElement(SpreadsheetNs + "row",
            new XAttribute("r", rowIndex.ToString(CultureInfo.InvariantCulture)));
        var column = 1;
        foreach (var value in values)
        {
            cancellationToken.ThrowIfCancellationRequested();
            row.Add(new XElement(SpreadsheetNs + "c",
                new XAttribute("r", CellReference(column++, rowIndex)),
                new XAttribute("t", "inlineStr"),
                new XElement(SpreadsheetNs + "is",
                    new XElement(SpreadsheetNs + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        FormatSummaryValue(value)))));
        }
        return row;
    }

    /// <summary>
    /// 将摘要值转换为不受公式解释的显示文本。
    /// </summary>
    /// <param name="value">错误属性值。</param>
    /// <returns>区域性无关的摘要文本。</returns>
    private static string FormatSummaryValue(object value)
    {
        if (value == null)
            return string.Empty;
        if (value is DateTime date)
            return date.ToString("O", CultureInfo.InvariantCulture);
        if (value is DateTimeOffset offset)
            return offset.ToString("O", CultureInfo.InvariantCulture);
        if (value is bool boolean)
            return boolean ? "TRUE" : "FALSE";
        return value is IFormattable formattable
            ? formattable.ToString(null, CultureInfo.InvariantCulture)
            : value.ToString();
    }

    /// <summary>
    /// 创建空的 SpreadsheetML 批注部件。
    /// </summary>
    /// <returns>批注 XML 文档。</returns>
    private static XDocument NewCommentsDocument() => new(
        new XDeclaration("1.0", "UTF-8", "yes"),
        new XElement(SpreadsheetNs + "comments",
            new XElement(SpreadsheetNs + "authors"),
            new XElement(SpreadsheetNs + "commentList")));

    /// <summary>
    /// 创建一个带纯文本的单元格批注。
    /// </summary>
    /// <param name="reference">目标 A1 单元格坐标。</param>
    /// <param name="authorIndex">作者表索引。</param>
    /// <param name="text">批注文本。</param>
    /// <returns>SpreadsheetML comment 元素。</returns>
    private static XElement CreateComment(string reference, int authorIndex, string text) =>
        new(SpreadsheetNs + "comment",
            new XAttribute("ref", reference),
            new XAttribute("authorId", authorIndex.ToString(CultureInfo.InvariantCulture)),
            new XElement(SpreadsheetNs + "text",
                new XElement(SpreadsheetNs + "t",
                    new XAttribute(XNamespace.Xml + "space", "preserve"), text ?? string.Empty)));

    /// <summary>
    /// 拼接现有批注中的文本运行。
    /// </summary>
    /// <param name="comment">既有批注 XML。</param>
    /// <returns>按文档顺序合并的批注文本。</returns>
    private static string CommentText(XElement comment) => string.Concat(comment
        .Descendants(SpreadsheetNs + "t").Select(element => element.Value));

    /// <summary>
    /// 创建包含标准 Office 批注形状类型的空 VML 文档。
    /// </summary>
    /// <returns>批注 VML 文档。</returns>
    private static XDocument NewVmlDocument() => new(
        new XDeclaration("1.0", "UTF-8", "yes"),
        new XElement("xml",
            new XAttribute(XNamespace.Xmlns + "v", VmlNs),
            new XAttribute(XNamespace.Xmlns + "o", VmlOfficeNs),
            new XAttribute(XNamespace.Xmlns + "x", VmlExcelNs),
            new XElement(VmlOfficeNs + "shapelayout",
                new XAttribute(VmlNs + "ext", "edit"),
                new XElement(VmlOfficeNs + "idmap", new XAttribute(VmlNs + "ext", "edit"),
                    new XAttribute("data", "1"))),
            new XElement(VmlNs + "shapetype",
                new XAttribute("id", "_x0000_t202"),
                new XAttribute("coordsize", "21600,21600"),
                new XAttribute(VmlOfficeNs + "spt", "202"),
                new XAttribute("path", "m,l,21600r21600,l21600,xe"),
                new XElement(VmlNs + "stroke", new XAttribute("joinstyle", "miter")),
                new XElement(VmlNs + "path", new XAttribute(VmlOfficeNs + "connecttype", "rect"),
                    new XAttribute("gradientshapeok", "t")))));

    /// <summary>
    /// 确保现有 VML 只包含可安全重建的单元格批注形状。
    /// </summary>
    /// <param name="document">现有或新建的 VML 文档。</param>
    private static void EnsureOnlyNoteShapes(XDocument document)
    {
        if (document.Root?.Name.LocalName != "xml")
            throw InvalidPackage("批注 VML 根元素无效。");
        var foreignShape = document.Root.Elements(VmlNs + "shape").FirstOrDefault(shape =>
            !shape.Elements(VmlExcelNs + "ClientData").Any(data => (string)data.Attribute("ObjectType") == "Note"));
        if (foreignShape != null)
            throw InvalidPackage("工作表 VML 包含非批注对象，无法安全重建批注绘图。");
    }

    /// <summary>
    /// 按批注集合重新建立 VML 注释形状，同时保留其他 VML 元素。
    /// </summary>
    /// <param name="document">待更新的 VML 文档。</param>
    /// <param name="references">全部批注单元格坐标。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void ReplaceNoteShapes(XDocument document, IEnumerable<string> references,
        CancellationToken cancellationToken)
    {
        var root = document.Root;
        foreach (var oldShape in root.Elements(VmlNs + "shape").Where(shape =>
            shape.Elements(VmlExcelNs + "ClientData").Any(data => (string)data.Attribute("ObjectType") == "Note"))
            .ToArray())
            oldShape.Remove();
        var id = 1025;
        foreach (var reference in references.OrderBy(reference => CellRow(reference))
            .ThenBy(reference => CellColumn(reference)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var shape = CreateVmlShape(id++, reference);
            var shapeType = root.Elements(VmlNs + "shapetype").LastOrDefault();
            if (shapeType == null)
                root.Add(shape);
            else
                shapeType.AddAfterSelf(shape);
        }
    }

    /// <summary>
    /// 创建 Excel 批注使用的 VML shape 和零基坐标。
    /// </summary>
    /// <param name="shapeId">VML 形状 ID。</param>
    /// <param name="reference">目标单元格坐标。</param>
    /// <returns>VML 注释形状。</returns>
    private static XElement CreateVmlShape(int shapeId, string reference)
    {
        var row = CellRow(reference) - 1;
        var column = CellColumn(reference) - 1;
        return new XElement(VmlNs + "shape",
            new XAttribute("id", $"_x0000_s{shapeId.ToString(CultureInfo.InvariantCulture)}"),
            new XAttribute("type", "#_x0000_t202"),
            new XAttribute("style", "position:absolute;visibility:hidden"),
            new XAttribute("fillcolor", "#ffffe1"),
            new XAttribute(VmlOfficeNs + "insetmode", "auto"),
            new XElement(VmlNs + "fill", new XAttribute("color2", "#ffffe1")),
            new XElement(VmlNs + "shadow", new XAttribute("on", "t"), new XAttribute("color", "black"),
                new XAttribute("obscured", "t")),
            new XElement(VmlNs + "path", new XAttribute(VmlOfficeNs + "connecttype", "none")),
            new XElement(VmlNs + "textbox", new XAttribute("style", "mso-direction-alt:auto"),
                new XElement("div", new XAttribute("style", "text-align:left"))),
            new XElement(VmlExcelNs + "ClientData", new XAttribute("ObjectType", "Note"),
                new XElement(VmlExcelNs + "MoveWithCells"), new XElement(VmlExcelNs + "SizeWithCells"),
                new XElement(VmlExcelNs + "AutoFill", "False"),
                new XElement(VmlExcelNs + "Row", row.ToString(CultureInfo.InvariantCulture)),
                new XElement(VmlExcelNs + "Column", column.ToString(CultureInfo.InvariantCulture))));
    }

    /// <summary>
    /// 设置工作表 legacyDrawing 关系，拒绝覆盖非批注 VML 绘图。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="relationshipId">VML 绘图关系 ID。</param>
    private static void AddLegacyDrawing(SheetState sheet, string relationshipId)
    {
        if (string.IsNullOrWhiteSpace(relationshipId))
            throw InvalidPackage("批注 VML 关系 ID 缺失。");
        var legacy = sheet.Document.Root?.Element(SpreadsheetNs + "legacyDrawing");
        if (legacy != null)
        {
            var existingId = (string)legacy.Attribute(OfficeRelNs + "id");
            if (!string.Equals(existingId, relationshipId, StringComparison.Ordinal))
                throw InvalidPackage("工作表已经引用其他 legacyDrawing，无法安全添加批注。");
            return;
        }
        var extension = sheet.Document.Root?.Element(SpreadsheetNs + "extLst");
        var drawing = new XElement(SpreadsheetNs + "legacyDrawing",
            new XAttribute(OfficeRelNs + "id", relationshipId));
        if (extension == null)
            sheet.Document.Root?.Add(drawing);
        else
            extension.AddBeforeSelf(drawing);
    }

    /// <summary>
    /// 在工作表关系中找到指定类型的关系。
    /// </summary>
    /// <param name="relationships">工作表关系文档。</param>
    /// <param name="suffix">关系类型末尾。</param>
    /// <returns>匹配关系；不存在时为 null。</returns>
    private static XElement FindRelationship(XDocument relationships, string suffix)
    {
        var type = OfficeRelationshipBase + suffix;
        var matches = relationships.Root?.Elements(PackageRelNs + "Relationship")
            .Where(element => (string)element.Attribute("Type") == type).ToArray()
            ?? Array.Empty<XElement>();
        if (matches.Length > 1)
            throw InvalidPackage($"工作表包含重复的 {suffix} 关系。");
        return matches.SingleOrDefault();
    }

    /// <summary>
    /// 确保关系文档包含指向指定部件的关系并返回其稳定节点。
    /// </summary>
    /// <param name="relationships">待更新的关系文档。</param>
    /// <param name="existing">已存在关系；不存在时为 null。</param>
    /// <param name="suffix">关系类型末尾。</param>
    /// <param name="targetPart">目标部件路径。</param>
    /// <param name="sourcePart">关系源部件路径。</param>
    private static void EnsureRelationship(XDocument relationships, XElement existing, string suffix,
        string targetPart, string sourcePart)
    {
        if (existing != null)
        {
            if (!string.Equals(ResolveTarget(sourcePart, RequiredInternalTarget(existing)), targetPart,
                    StringComparison.OrdinalIgnoreCase))
                throw InvalidPackage($"已有 {suffix} 关系目标与批注部件不一致。");
            return;
        }
        var id = NextRelationshipId(relationships);
        relationships.Root.Add(new XElement(PackageRelNs + "Relationship",
            new XAttribute("Id", id),
            new XAttribute("Type", OfficeRelationshipBase + suffix),
            new XAttribute("Target", RelativeTarget(sourcePart, targetPart))));
    }

    /// <summary>
    /// 创建新的关系 ID，避免与部件中任意现有关系冲突。
    /// </summary>
    /// <param name="relationships">关系 XML 文档。</param>
    /// <returns>未使用的 rId 名称。</returns>
    private static string NextRelationshipId(XDocument relationships)
    {
        var used = relationships.Root?.Elements(PackageRelNs + "Relationship")
            .Select(element => (string)element.Attribute("Id")).Where(id => !string.IsNullOrEmpty(id))
            .ToHashSet(StringComparer.Ordinal) ?? new HashSet<string>(StringComparer.Ordinal);
        for (var index = 1; ; index++)
        {
            var candidate = "rId" + index.ToString(CultureInfo.InvariantCulture);
            if (!used.Contains(candidate))
                return candidate;
        }
    }

    /// <summary>
    /// 返回已有关系部件或创建新的空关系部件。
    /// </summary>
    /// <param name="part">关系部件路径。</param>
    /// <param name="package">包状态。</param>
    /// <returns>关系 XML 文档。</returns>
    private static XDocument GetRelationships(PackageState package, string part)
    {
        if (package.Entries.TryGetValue(part, out var entry))
            return LoadXml(entry, package.Options, package.CancellationToken, "关系部件");
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(PackageRelNs + "Relationships"));
    }

    /// <summary>
    /// 追加工作表、批注和 VML 的内容类型声明。
    /// </summary>
    /// <param name="package">待更新的包状态。</param>
    /// <param name="commentsPart">批注部件路径。</param>
    private static void EnsureCommentContentTypes(PackageState package, string commentsPart)
    {
        var root = package.ContentTypes.Root;
        if (root == null || root.Name != ContentTypeNs + "Types")
            throw InvalidPackage("内容类型部件根元素无效。");
        EnsureOverride(root, commentsPart,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml");
        var vml = root.Elements(ContentTypeNs + "Default").FirstOrDefault(element =>
            string.Equals((string)element.Attribute("Extension"), "vml", StringComparison.OrdinalIgnoreCase));
        if (vml == null)
            root.Add(new XElement(ContentTypeNs + "Default",
                new XAttribute("Extension", "vml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.vmlDrawing")));
        package.ReplaceXml("[Content_Types].xml", package.ContentTypes);
    }

    /// <summary>
    /// 为新增错误摘要工作表添加内容类型覆盖项。
    /// </summary>
    /// <param name="package">待更新的包状态。</param>
    /// <param name="sheetPart">摘要工作表部件路径。</param>
    private static void EnsureWorksheetContentType(PackageState package, string sheetPart)
    {
        var root = package.ContentTypes.Root;
        if (root == null || root.Name != ContentTypeNs + "Types")
            throw InvalidPackage("内容类型部件根元素无效。");
        EnsureOverride(root, sheetPart,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
        package.ReplaceXml("[Content_Types].xml", package.ContentTypes);
    }

    /// <summary>
    /// 确保内容类型部件包含唯一且正确的部件覆盖项。
    /// </summary>
    /// <param name="root">内容类型根元素。</param>
    /// <param name="part">部件路径。</param>
    /// <param name="contentType">目标内容类型。</param>
    private static void EnsureOverride(XElement root, string part, string contentType)
    {
        var partName = "/" + part;
        var matches = root.Elements(ContentTypeNs + "Override").Where(element =>
            string.Equals((string)element.Attribute("PartName"), partName,
                StringComparison.OrdinalIgnoreCase)).ToArray();
        if (matches.Length > 1)
            throw InvalidPackage($"内容类型部件包含重复声明: {partName}");
        if (matches.Length == 1)
        {
            if (!string.Equals((string)matches[0].Attribute("ContentType"), contentType,
                    StringComparison.OrdinalIgnoreCase))
                throw InvalidPackage($"部件内容类型声明不匹配: {partName}");
            return;
        }
        root.Add(new XElement(ContentTypeNs + "Override",
            new XAttribute("PartName", partName), new XAttribute("ContentType", contentType)));
    }

    /// <summary>
    /// 计算不与原工作簿工作表重名的失败摘要工作表名称。
    /// </summary>
    /// <param name="workbook">工作簿 XML。</param>
    /// <param name="baseName">优先使用的名称。</param>
    /// <returns>不区分大小写唯一的工作表名。</returns>
    private static string UniqueSheetName(XDocument workbook, string baseName)
    {
        var names = workbook.Root?.Element(SpreadsheetNs + "sheets")?
            .Elements(SpreadsheetNs + "sheet").Select(sheet => (string)sheet.Attribute("name"))
            .Where(name => !string.IsNullOrEmpty(name)).ToHashSet(StringComparer.OrdinalIgnoreCase)
            ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!names.Contains(baseName))
            return baseName;
        for (var index = 1; ; index++)
        {
            var suffix = "_" + index.ToString(CultureInfo.InvariantCulture);
            var candidate = baseName.Substring(0, Math.Min(baseName.Length, 31 - suffix.Length)) + suffix;
            if (!names.Contains(candidate))
                return candidate;
        }
    }

    /// <summary>
    /// 创建单元格的 A1 坐标。
    /// </summary>
    /// <param name="column">一基列号。</param>
    /// <param name="row">一基行号。</param>
    /// <returns>A1 格式坐标。</returns>
    private static string CellReference(int column, int row)
    {
        var current = column;
        var letters = string.Empty;
        while (current > 0)
        {
            current--;
            letters = (char)('A' + current % 26) + letters;
            current /= 26;
        }
        return letters + row.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 从 A1 单元格坐标读取一基行号。
    /// </summary>
    /// <param name="reference">A1 坐标。</param>
    /// <returns>一基行号。</returns>
    private static int CellRow(string reference)
    {
        var index = 0;
        while (index < reference.Length && char.IsLetter(reference[index]))
            index++;
        return ParsePositiveInt(reference.Substring(index), "批注单元格行号");
    }

    /// <summary>
    /// 从 A1 单元格坐标读取一基列号。
    /// </summary>
    /// <param name="reference">A1 坐标。</param>
    /// <returns>一基列号。</returns>
    private static int CellColumn(string reference)
    {
        var value = 0;
        foreach (var character in reference.TakeWhile(char.IsLetter))
            value = checked(value * 26 + char.ToUpperInvariant(character) - 'A' + 1);
        if (value <= 0)
            throw InvalidPackage("批注单元格列号无效。");
        return value;
    }

    /// <summary>
    /// 创建描述输入包或部件错误的结构化导入异常。
    /// </summary>
    /// <param name="message">异常消息。</param>
    /// <param name="innerException">底层异常。</param>
    /// <returns>结构化 XLSX 失败工作簿异常。</returns>
    private static BingOfficesImportException InvalidPackage(string message,
        Exception innerException = null) => new(message, innerException, "ExcelDataReader",
            BingOfficesStage.Preflight);

    /// <summary>
    /// 创建 AnnotatedOriginal 资源限制异常。
    /// </summary>
    /// <param name="message">资源限制说明。</param>
    /// <param name="innerException">底层异常。</param>
    /// <returns>结构化资源限制异常。</returns>
    private static BingOfficesResourceLimitException ResourceLimit(string message,
        Exception innerException = null) => new(message, innerException, "ExcelDataReader",
            BingOfficesOperation.Import, BingOfficesStage.Preflight);

    /// <summary>
    /// 将 XML 文档编码为无 BOM UTF-8 字节。
    /// </summary>
    /// <param name="document">待编码文档。</param>
    /// <returns>UTF-8 XML 字节。</returns>
    private static byte[] SaveXml(XDocument document)
    {
        using var output = new MemoryStream();
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = false,
            Indent = false,
            CloseOutput = false
        };
        using (var writer = XmlWriter.Create(output, settings))
            document.Save(writer);
        return output.ToArray();
    }

    /// <summary>
    /// 比较工作表名称和一基行号时采用工作表名称不区分大小写的语义。
    /// </summary>
    private sealed class SheetRowComparer : IEqualityComparer<(string SheetName, int RowIndex)>
    {
        /// <summary>
        /// 比较两个工作表行键。
        /// </summary>
        /// <param name="x">第一个键。</param>
        /// <param name="y">第二个键。</param>
        /// <returns>键相等时返回 <see langword="true" />。</returns>
        public bool Equals((string SheetName, int RowIndex) x, (string SheetName, int RowIndex) y) =>
            x.RowIndex == y.RowIndex && string.Equals(x.SheetName, y.SheetName,
                StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// 计算工作表行键的哈希值。
        /// </summary>
        /// <param name="obj">待计算的键。</param>
        /// <returns>哈希值。</returns>
        public int GetHashCode((string SheetName, int RowIndex) obj) =>
            HashCode.Combine(StringComparer.OrdinalIgnoreCase.GetHashCode(obj.SheetName), obj.RowIndex);
    }

    /// <summary>
    /// 按工作表名和单元格坐标合并错误。
    /// </summary>
    private sealed class ErrorCellComparer : IEqualityComparer<(string SheetName, int RowIndex, int ColumnIndex)>
    {
        /// <summary>
        /// 比较两个错误单元格键。
        /// </summary>
        /// <param name="x">第一个键。</param>
        /// <param name="y">第二个键。</param>
        /// <returns>键相等时返回 <see langword="true" />。</returns>
        public bool Equals((string SheetName, int RowIndex, int ColumnIndex) x,
            (string SheetName, int RowIndex, int ColumnIndex) y) => x.RowIndex == y.RowIndex
            && x.ColumnIndex == y.ColumnIndex
            && string.Equals(x.SheetName, y.SheetName, StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// 计算错误单元格键的哈希值。
        /// </summary>
        /// <param name="obj">待计算的键。</param>
        /// <returns>哈希值。</returns>
        public int GetHashCode((string SheetName, int RowIndex, int ColumnIndex) obj) =>
            HashCode.Combine(StringComparer.OrdinalIgnoreCase.GetHashCode(obj.SheetName), obj.RowIndex,
                obj.ColumnIndex);
    }

    /// <summary>
    /// 保存待更新 XLSX 包的解析状态和部件变更。
    /// </summary>
    private sealed class PackageState
    {
        /// <summary>
        /// 初始化一个 <see cref="PackageState" /> 类型的实例。
        /// </summary>
        /// <param name="entries">源包部件。</param>
        /// <param name="contentTypes">内容类型文档。</param>
        /// <param name="workbook">工作簿文档。</param>
        /// <param name="workbookRelationships">工作簿关系文档。</param>
        /// <param name="workbookPart">工作簿部件路径。</param>
        /// <param name="workbookRelationshipsPart">工作簿关系部件路径。</param>
        /// <param name="sheets">源工作表。</param>
        /// <param name="errorSheetName">错误摘要表名。</param>
        /// <param name="cellCount">源单元格数量。</param>
        /// <param name="rowCount">源行数量。</param>
        /// <param name="pictureCount">源图片数量。</param>
        /// <param name="capture">读取阶段捕获。</param>
        /// <param name="options">输出选项。</param>
        /// <param name="cancellationToken">取消令牌。</param>
        internal PackageState(Dictionary<string, ZipArchiveEntry> entries, XDocument contentTypes,
            XDocument workbook, XDocument workbookRelationships, string workbookPart,
            string workbookRelationshipsPart, List<SheetState> sheets, string errorSheetName,
            long cellCount, long rowCount, ExcelDataReaderFailureWorkbookCapture capture,
            int pictureCount, ExcelImportFailureOptions options,
            CancellationToken cancellationToken)
        {
            Entries = entries;
            ContentTypes = contentTypes;
            Workbook = workbook;
            WorkbookRelationships = workbookRelationships;
            WorkbookPart = workbookPart;
            WorkbookRelationshipsPart = workbookRelationshipsPart;
            Sheets = sheets;
            ErrorSheetName = errorSheetName;
            CellCount = cellCount;
            RowCount = rowCount;
            PictureCount = pictureCount;
            Capture = capture;
            Options = options;
            CancellationToken = cancellationToken;
        }

        /// <summary>
        /// 获取源包部件索引。
        /// </summary>
        internal Dictionary<string, ZipArchiveEntry> Entries { get; }

        /// <summary>
        /// 获取需替换的现有部件 XML 字节。
        /// </summary>
        internal Dictionary<string, byte[]> Replacements { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 获取需要新增的部件 XML 字节。
        /// </summary>
        internal Dictionary<string, byte[]> Additions { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 获取内容类型文档。
        /// </summary>
        internal XDocument ContentTypes { get; }

        /// <summary>
        /// 获取工作簿文档。
        /// </summary>
        internal XDocument Workbook { get; }

        /// <summary>
        /// 获取工作簿关系文档。
        /// </summary>
        internal XDocument WorkbookRelationships { get; }

        /// <summary>
        /// 获取工作簿部件路径。
        /// </summary>
        internal string WorkbookPart { get; }

        /// <summary>
        /// 获取工作簿关系部件路径。
        /// </summary>
        internal string WorkbookRelationshipsPart { get; }

        /// <summary>
        /// 获取源工作表状态。
        /// </summary>
        internal List<SheetState> Sheets { get; }

        /// <summary>
        /// 获取错误摘要工作表名称。
        /// </summary>
        internal string ErrorSheetName { get; }

        /// <summary>
        /// 获取源包中所有已用单元格数量。
        /// </summary>
        internal long CellCount { get; }

        /// <summary>
        /// 获取源包中所有工作表行数量。
        /// </summary>
        internal long RowCount { get; }

        /// <summary>
        /// 获取源包中的图片数量。
        /// </summary>
        internal int PictureCount { get; }

        /// <summary>
        /// 获取前向导入捕获数据。
        /// </summary>
        internal ExcelDataReaderFailureWorkbookCapture Capture { get; }

        /// <summary>
        /// 获取当前输出资源预算。
        /// </summary>
        internal ExcelImportFailureOptions Options { get; }

        /// <summary>
        /// 获取当前取消令牌。
        /// </summary>
        internal CancellationToken CancellationToken { get; }

        /// <summary>
        /// 按需读取已有关系文档或创建新的关系文档。
        /// </summary>
        /// <param name="part">关系部件路径。</param>
        /// <returns>关系文档。</returns>
        internal XDocument GetOrCreateRelationships(string part) => GetRelationships(this, part);

        /// <summary>
        /// 生成不与输入包部件冲突的部件路径。
        /// </summary>
        /// <param name="baseName">无扩展名基础路径。</param>
        /// <param name="extension">部件扩展名。</param>
        /// <returns>可使用的部件路径。</returns>
        internal string UniquePart(string baseName, string extension)
        {
            var candidate = baseName + extension;
            var index = 1;
            while (Entries.ContainsKey(candidate) || Additions.ContainsKey(candidate))
            {
                candidate = baseName + index.ToString(CultureInfo.InvariantCulture) + extension;
                index++;
            }
            return candidate;
        }

        /// <summary>
        /// 保存现有部件替换或新部件添加。
        /// </summary>
        /// <param name="part">目标部件路径。</param>
        /// <param name="document">待保存的 XML 文档。</param>
        internal void ReplaceXml(string part, XDocument document)
        {
            var bytes = SaveXml(document);
            if (Entries.ContainsKey(part))
                Replacements[part] = bytes;
            else
                Additions[part] = bytes;
        }
    }

    /// <summary>
    /// 保存一个原工作表及其批注输出状态。
    /// </summary>
    private sealed class SheetState
    {
        /// <summary>
        /// 初始化一个 <see cref="SheetState" /> 类型的实例。
        /// </summary>
        /// <param name="name">工作表名称。</param>
        /// <param name="part">工作表部件路径。</param>
        /// <param name="sheetId">工作簿中的工作表 ID。</param>
        internal SheetState(string name, string part, int sheetId)
        {
            Name = name;
            Part = part;
            SheetId = sheetId;
        }

        /// <summary>
        /// 获取工作表名称。
        /// </summary>
        internal string Name { get; }

        /// <summary>
        /// 获取工作表部件路径。
        /// </summary>
        internal string Part { get; }

        /// <summary>
        /// 获取工作表 ID。
        /// </summary>
        internal int SheetId { get; }

        /// <summary>
        /// 获取当前工作表批注数。
        /// </summary>
        internal int CommentCount { get; set; }

        /// <summary>
        /// 获取工作表 XML；首次访问时载入原部件。
        /// </summary>
        internal XDocument Document { get; set; }
    }

    /// <summary>
    /// 对 ZIP 目标写入执行累计序列化字节限制。
    /// </summary>
    private sealed class LimitedWriteStream : Stream
    {
        /// <summary>
        /// 接收输出数据的底层流。
        /// </summary>
        private readonly Stream _inner;

        /// <summary>
        /// 输出允许的最大字节数。
        /// </summary>
        private readonly long? _maximum;

        /// <summary>
        /// 已输出字节数。
        /// </summary>
        private long _written;

        /// <summary>
        /// 初始化一个 <see cref="LimitedWriteStream" /> 类型的实例。
        /// </summary>
        /// <param name="inner">内存输出流。</param>
        /// <param name="maximum">最大序列化字节数。</param>
        internal LimitedWriteStream(Stream inner, long? maximum)
        {
            _inner = inner;
            _maximum = maximum;
        }

        /// <summary>
        /// 获取流是否可读；此流不可读。
        /// </summary>
        public override bool CanRead => false;

        /// <summary>
        /// 获取流是否可查找；此流不可查找。
        /// </summary>
        public override bool CanSeek => false;

        /// <summary>
        /// 获取流是否可写。
        /// </summary>
        public override bool CanWrite => _inner.CanWrite;

        /// <summary>
        /// 获取累计写入长度。
        /// </summary>
        public override long Length => _written;

        /// <summary>
        /// 获取累计写入位置；不支持设置。
        /// </summary>
        public override long Position
        {
            get => _written;
            set => throw new NotSupportedException();
        }

        /// <summary>
        /// 刷新输出流。
        /// </summary>
        public override void Flush() => _inner.Flush();

        /// <summary>
        /// 读取操作不受支持。
        /// </summary>
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        /// <summary>
        /// 查找操作不受支持。
        /// </summary>
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        /// <summary>
        /// 设置长度操作不受支持。
        /// </summary>
        public override void SetLength(long value) => throw new NotSupportedException();

        /// <summary>
        /// 写入数据并在超限前拒绝该次写入。
        /// </summary>
        public override void Write(byte[] buffer, int offset, int count)
        {
            Check(count);
            _inner.Write(buffer, offset, count);
            _written += count;
        }

        /// <summary>
        /// 释放包装器但保留底层调用方流。
        /// </summary>
        /// <param name="disposing">是否由 Dispose 调用。</param>
        protected override void Dispose(bool disposing) => base.Dispose(disposing);

        /// <summary>
        /// 检查下一次写入是否超过配置的字节上限。
        /// </summary>
        /// <param name="count">下一次写入字节数。</param>
        private void Check(int count)
        {
            if (_maximum.HasValue && count > _maximum.Value - _written)
                throw ResourceLimit($"AnnotatedOriginal 输出超过序列化字节限制: {_maximum.Value}");
        }
    }
}
