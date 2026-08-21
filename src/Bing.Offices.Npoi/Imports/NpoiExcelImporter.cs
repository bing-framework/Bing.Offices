using System.Globalization;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Metadata;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Validations;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// 基于 NPOI 的单遍流式 Excel 导入器。
/// </summary>
internal sealed class NpoiExcelImporter : IExcelImporter
{
    /// <summary>
    /// 当前导入器使用的校验规则。
    /// </summary>
    private readonly IReadOnlyList<IExcelValidationRule> _validationRules;

    /// <summary>
    /// 当前导入器使用的值转换器。
    /// </summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;

    /// <summary>
    /// 当前导入器使用的旧版仅文本单元格转换器。
    /// </summary>
    private readonly IReadOnlyList<ICellValueConverter> _legacyValueConverters;

    /// <summary>
    /// 当前导入器使用的命名配置校验规则。
    /// </summary>
    private readonly IReadOnlyList<INamedExcelValidationRule> _namedValidationRules;

    /// <summary>
    /// 初始化一个<see cref="NpoiExcelImporter"/>类型的实例。
    /// </summary>
    /// <param name="validationRules">校验规则集合。</param>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="namedValidationRules">命名配置校验规则集合。</param>
    /// <param name="legacyValueConverters">旧版仅文本单元格转换器集合。</param>
    public NpoiExcelImporter(IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null,
        IEnumerable<ICellValueConverter> legacyValueConverters = null)
    {
        _validationRules = validationRules?.ToArray() ?? ExcelValidationRules.CreateDefault();
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _namedValidationRules = namedValidationRules?.ToArray() ?? Array.Empty<INamedExcelValidationRule>();
        _legacyValueConverters = legacyValueConverters?.ToArray() ?? Array.Empty<ICellValueConverter>();
    }

    /// <inheritdoc />
    public ExcelWorkbookImportResult<TWorkbook> Import<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (!source.CanRead)
            throw new ArgumentException("输入流不可读取。", nameof(source));
        cancellationToken.ThrowIfCancellationRequested();

        using var bufferedSource = new MemoryStream();
        CopyTo(source, bufferedSource, cancellationToken, request.ResourceLimits?.MaxInputBytes);
        bufferedSource.Position = 0;
        using var workbook = WorkbookFactory.Create(bufferedSource);
        var root = new TWorkbook();
        var sheetResults = new List<ExcelSheetImportResult>();
        var errors = new ExcelImportErrorCollector(request.ResourceLimits?.MaxErrors);
        var sourceLocations = new Dictionary<object, SourceLocation>(ReferenceObjectComparer.Instance);
        var runtime = new ExcelImportRuntime(request.ResourceLimits);
        foreach (var sheetRequest in request.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached || runtime.RowLimitReached)
                break;
            var sheetIndex = ResolveSheetIndex(workbook, sheetRequest.Selector, request.SheetNameComparison);
            if (sheetIndex < 0)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                    $"缺少请求的 Sheet: {GetSelectorDescription(sheetRequest.Selector)}",
                    GetSelectorDescription(sheetRequest.Selector), 0, 0, null));
                continue;
            }
            if (workbook.IsSheetHidden(sheetIndex) || workbook.IsSheetVeryHidden(sheetIndex))
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                    $"请求的 Sheet 被隐藏: {workbook.GetSheetName(sheetIndex)}", workbook.GetSheetName(sheetIndex),
                    0, 0, null));
                continue;
            }
            var sheet = workbook.GetSheetAt(sheetIndex);
            try
            {
                ImportTypedSheet(sheet, sheetRequest, root, sheetResults, errors, request.ValidationMode,
                    request.ResourceLimits, request.UnsupportedFeaturePolicy, sourceLocations, runtime,
                    cancellationToken);
            }
            catch (SheetStructureException exception)
            {
                var error = new ExcelImportError(ExcelImportErrorCode.InvalidHeader, exception.Message,
                    sheet.SheetName, sheetRequest.HeaderRowIndex + 1, 0, null);
                errors.Add(error);
                sheetResults.Add(new ExcelSheetImportResult(sheet.SheetName, sheetRequest.ItemType,
                    Array.Empty<int>(), new[] { error }));
            }
        }

        foreach (var relation in request.Relations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
            BindTypedRelation(root, relation, errors, sourceLocations, cancellationToken);
        }
        WriteFailureWorkbook(workbook, request.FailureOptions, errors.Errors, request.Sheets, cancellationToken);
        return new ExcelWorkbookImportResult<TWorkbook>(root, sheetResults, errors.Errors,
            errors.IsTruncated, errors.MaxErrors);
    }

    private static int ResolveSheetIndex(IWorkbook workbook, ExcelSheetSelector selector,
        ExcelNameComparison comparison)
    {
        if (selector.Kind == ExcelSheetSelectorKind.ByIndex)
            return selector.Index.Value < workbook.NumberOfSheets ? selector.Index.Value : -1;
        var comparer = comparison == ExcelNameComparison.Ordinal
            ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;
        for (var index = 0; index < workbook.NumberOfSheets; index++)
        {
            if (string.Equals(workbook.GetSheetName(index), selector.Name, comparer))
                return index;
        }
        return -1;
    }

    private static string GetSelectorDescription(ExcelSheetSelector selector) =>
        selector.Kind == ExcelSheetSelectorKind.ByIndex ? $"#{selector.Index.Value}" : selector.Name;

    /// <summary>
    /// 通过一次类型擦除调用执行单个 Sheet 导入计划。
    /// </summary>
    private void ImportTypedSheet<TWorkbook>(ISheet sheet, ExcelSheetImportRequest request, TWorkbook root,
        ICollection<ExcelSheetImportResult> sheetResults, ExcelImportErrorCollector errors,
        ExcelImportValidationMode validationMode, ExcelResourceLimits resourceLimits,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy,
        IDictionary<object, SourceLocation> sourceLocations,
        ExcelImportRuntime runtime,
        CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        var method = GetType().GetMethod(nameof(ImportTypedSheetCore), BindingFlags.Instance | BindingFlags.NonPublic);
        try
        {
            method.MakeGenericMethod(typeof(TWorkbook), request.ItemType).Invoke(this,
                new object[] { sheet, request, root, sheetResults, errors, validationMode, resourceLimits,
                    unsupportedFeaturePolicy, sourceLocations, runtime, cancellationToken });
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }

    /// <summary>
    /// 导入一个具体实体类型，并将成功项写入 Workbook 根集合。
    /// </summary>
    private void ImportTypedSheetCore<TWorkbook, TItem>(ISheet sheet, ExcelSheetImportRequest request,
        TWorkbook root, ICollection<ExcelSheetImportResult> sheetResults, ExcelImportErrorCollector errors,
        ExcelImportValidationMode validationMode, ExcelResourceLimits resourceLimits,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy,
        IDictionary<object, SourceLocation> sourceLocations,
        ExcelImportRuntime runtime,
        CancellationToken cancellationToken)
        where TWorkbook : class, new()
        where TItem : class, new()
    {
        var options = new ExcelImportExecutionOptions<TItem>
        {
            HeaderRowIndex = request.HeaderRowIndex,
            DataRowIndex = request.DataRowStartIndex,
            MaxColumnLength = request.MaxColumnLength,
            ReadColumnRange = request.ReadColumnRange,
            HeaderComparison = request.HeaderComparison,
            HeaderWhitespace = request.HeaderWhitespace,
            BodyWhitespace = request.BodyWhitespace,
            ValidationMode = validationMode,
            UnsupportedFeaturePolicy = unsupportedFeaturePolicy,
            DynamicTargetGetter = request.DynamicTargetGetter,
            HeaderMatch = request.HeaderMatch,
            ValidateMode = request.ValidateMode,
            Culture = request.Culture,
            DynamicColumns = request.DynamicColumns,
            FailOnUnknownDynamicColumns = request.FailOnUnknownDynamicColumns,
            EnabledEmptyLine = request.EnabledEmptyLine,
            IgnoreEmptyLineAfterData = request.IgnoreEmptyLineAfterData,
            MappingConfiguration = request.MappingConfiguration,
            MappingProfile = request.MappingProfile as Configurations.ExcelMappingProfile<TItem>
        };
        var items = new List<TItem>();
        var rows = new List<int>();
        var sheetErrors = errors.CreateChild();
        ImportSheet(sheet, options, items, sheetErrors, runtime, cancellationToken, rows);
        var target = (ICollection<TItem>)request.Target(root);
        if (target == null)
            throw new InvalidOperationException($"Workbook 导入目标集合不可写入: {request.Name}");
        foreach (var item in items)
            target.Add(item);
        for (var index = 0; index < items.Count; index++)
            sourceLocations[items[index]] = new SourceLocation(sheet.SheetName, rows[index] + 1);
        var result = new ExcelSheetImportResult(sheet.SheetName, typeof(TItem), rows,
            sheetErrors.Errors);
        sheetResults.Add(result);
    }

    /// <summary>
    /// 线性绑定显式父子关系，不根据属性名猜测关系。
    /// </summary>
    private static void BindTypedRelation<TWorkbook>(TWorkbook root, ExcelRelationRequest request,
        ExcelImportErrorCollector errors, IReadOnlyDictionary<object, SourceLocation> sourceLocations,
        CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        var method = typeof(NpoiExcelImporter).GetMethod(nameof(BindTypedRelationCore),
            BindingFlags.Static | BindingFlags.NonPublic);
        method.MakeGenericMethod(typeof(TWorkbook), request.ParentType, request.ChildType, request.ParentKey.Method.ReturnType)
            .Invoke(null,
            new object[] { root, request, errors, sourceLocations, cancellationToken });
    }

    /// <summary>
    /// 执行父子键索引和导航集合绑定。
    /// </summary>
    private static void BindTypedRelationCore<TWorkbook, TParent, TChild, TKey>(TWorkbook root,
        ExcelRelationRequest request, ExcelImportErrorCollector errors,
        IReadOnlyDictionary<object, SourceLocation> sourceLocations, CancellationToken cancellationToken)
        where TWorkbook : class, new()
        where TParent : class
        where TChild : class
    {
        var parents = (ICollection<TParent>)request.Parents(root);
        var children = (ICollection<TChild>)request.Children(root);
        if (parents == null || children == null)
            throw new InvalidOperationException("关系绑定的父集合或子集合不可写入。");
        var parentByKey = new Dictionary<TKey, TParent>((IEqualityComparer<TKey>)request.Comparer);
        foreach (var parent in parents)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
            var key = ((Func<TParent, TKey>)request.ParentKey)(parent);
            if (key == null)
            {
                errors.Add(CreateRelationshipError("父项键为空。", sourceLocations, parent, key));
                continue;
            }
            if (!parentByKey.TryAdd(key, parent))
                errors.Add(CreateRelationshipError($"父项键重复: {key}", sourceLocations, parent, key));
        }

        var childByParent = new Dictionary<TParent, List<TChild>>();
        foreach (var child in children)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
            var key = ((Func<TChild, TKey>)request.ChildKey)(child);
            if (key == null)
            {
                errors.Add(CreateRelationshipError("子项键为空。", sourceLocations, child, key));
                continue;
            }
            if (!parentByKey.TryGetValue(key, out var parent))
            {
                errors.Add(CreateRelationshipError($"子项找不到父项: {key}", sourceLocations, child, key));
                continue;
            }
            if (!childByParent.TryGetValue(parent, out var list))
                childByParent[parent] = list = new List<TChild>();
            list.Add(child);
        }

        foreach (var pair in childByParent)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
            var target = (ICollection<TChild>)request.Navigation(pair.Key);
            if (target == null)
            {
                errors.Add(CreateRelationshipError("导航集合为空且不可写入。", sourceLocations, pair.Key, null));
                continue;
            }
            foreach (var child in pair.Value)
                target.Add(child);
        }
    }

    private static ExcelImportError CreateRelationshipError(string message,
        IReadOnlyDictionary<object, SourceLocation> sourceLocations, object source, object key)
    {
        sourceLocations.TryGetValue(source, out var location);
        return new ExcelImportError(ExcelImportErrorCode.Relationship, message,
            location?.SheetName ?? string.Empty, location?.RowIndex ?? 0, 0, null, "Key", null, key);
    }

    /// <inheritdoc />
    /// <summary>
    /// 将输入流复制到内部缓冲，并在每个数据块之间检查取消状态。
    /// </summary>
    /// <param name="source">调用方拥有的输入流。</param>
    /// <param name="destination">实现拥有的缓冲流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void CopyTo(Stream source, Stream destination, CancellationToken cancellationToken,
        long? maxBytes = null)
    {
        var buffer = new byte[81920];
        long total = 0;
        int count;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            count = source.Read(buffer, 0, buffer.Length);
            if (count == 0)
                break;
            total += count;
            if (maxBytes.HasValue && total > maxBytes.Value)
                throw new InvalidOperationException($"输入工作簿超过最大字节数: {maxBytes.Value}");
            cancellationToken.ThrowIfCancellationRequested();
            destination.Write(buffer, 0, count);
        }
    }

    private static void WriteFailureWorkbook(IWorkbook workbook, ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors, IReadOnlyList<ExcelSheetImportRequest> requests,
        CancellationToken cancellationToken)
    {
        if (options == null || options.Mode == ExcelImportFailureWorkbookMode.None || errors.Count == 0)
            return;
        cancellationToken.ThrowIfCancellationRequested();
        IWorkbook outputWorkbook = workbook;
        IWorkbook independentWorkbook = null;
        try
        {
            if (options.Mode == ExcelImportFailureWorkbookMode.ErrorRowsOnly)
                outputWorkbook = independentWorkbook = CreateErrorRowsWorkbook(workbook, errors, requests,
                    cancellationToken);
            else
                AnnotateErrors(outputWorkbook, errors);

            WriteFailureSummary(outputWorkbook, errors, cancellationToken);
            using var output = new MemoryStream();
            outputWorkbook.Write(output, false);
            var bytes = output.ToArray();
            if (options.MaxBytes.HasValue && bytes.LongLength > options.MaxBytes.Value)
                throw new InvalidOperationException($"失败工作簿超过最大字节数: {options.MaxBytes.Value}");
            WriteBytes(options.Destination, bytes, cancellationToken);
        }
        finally
        {
            independentWorkbook?.Close();
        }
    }

    private static void WriteFailureSummary(IWorkbook workbook, IReadOnlyCollection<ExcelImportError> errors,
        CancellationToken cancellationToken)
    {
        var summaryName = "_ImportErrors";
        var suffix = 1;
        while (workbook.GetSheet(summaryName) != null)
            summaryName = $"_ImportErrors{suffix++}";
        var summary = workbook.CreateSheet(summaryName);
        var header = summary.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Message");
        header.CreateCell(2).SetCellValue("Sheet");
        header.CreateCell(3).SetCellValue("Row");
        header.CreateCell(4).SetCellValue("Column");
        header.CreateCell(5).SetCellValue("Property");
        header.CreateCell(6).SetCellValue("Header");
        header.CreateCell(7).SetCellValue("RawValue");
        var rowIndex = 1;
        foreach (var error in errors)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = summary.CreateRow(rowIndex++);
            row.CreateCell(0).SetCellValue(error.Code.ToString());
            row.CreateCell(1).SetCellValue(error.Message ?? string.Empty);
            row.CreateCell(2).SetCellValue(error.SheetName ?? string.Empty);
            row.CreateCell(3).SetCellValue(error.RowIndex);
            row.CreateCell(4).SetCellValue(error.ColumnIndex);
            row.CreateCell(5).SetCellValue(error.PropertyName ?? string.Empty);
            row.CreateCell(6).SetCellValue(error.Header ?? string.Empty);
            row.CreateCell(7).SetCellValue(FormatRawValue(error.RawValue));
        }
    }

    private static IWorkbook CreateErrorRowsWorkbook(IWorkbook source, IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyList<ExcelSheetImportRequest> requests, CancellationToken cancellationToken)
    {
        var destination = source is NPOI.HSSF.UserModel.HSSFWorkbook
            ? (IWorkbook)new NPOI.HSSF.UserModel.HSSFWorkbook()
            : new NPOI.XSSF.UserModel.XSSFWorkbook();
        var groups = errors.Where(error => error.RowIndex > 1)
            .GroupBy(error => error.SheetName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
        foreach (var group in groups)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sourceSheet = source.GetSheet(group.Key);
            if (sourceSheet == null)
                continue;
            var destinationSheet = destination.CreateSheet(sourceSheet.SheetName);
            var headerRowIndex = requests.FirstOrDefault(request => string.Equals(request.Name,
                sourceSheet.SheetName, StringComparison.OrdinalIgnoreCase))?.HeaderRowIndex
                ?? sourceSheet.FirstRowNum;
            var sourceRows = new HashSet<int>(group.Value.Select(error => error.RowIndex - 1))
            {
                headerRowIndex
            };
            var rowMap = sourceRows.OrderBy(index => index)
                .Select((sourceRow, targetRow) => (sourceRow, targetRow))
                .ToDictionary(item => item.sourceRow, item => item.targetRow);
            var styleCache = new Dictionary<short, ICellStyle>();
            foreach (var pair in rowMap)
            {
                cancellationToken.ThrowIfCancellationRequested();
                CopyRow(sourceSheet.GetRow(pair.Key), destinationSheet.CreateRow(pair.Value), destination,
                    styleCache);
            }
            AddFailureColumns(destinationSheet, group.Value, rowMap,
                sourceSheet.GetRow(headerRowIndex)?.LastCellNum ?? 0);
            CopyMergedRegions(sourceSheet, destinationSheet, rowMap);
            CopyDataValidations(sourceSheet, destinationSheet, rowMap, cancellationToken);
            CopyPictures(sourceSheet, destinationSheet, rowMap, cancellationToken);
        }
        return destination;
    }

    private static void CopyRow(IRow sourceRow, IRow destinationRow, IWorkbook destination,
        IDictionary<short, ICellStyle> styleCache)
    {
        if (sourceRow == null)
            return;
        destinationRow.Height = sourceRow.Height;
        foreach (var sourceCell in sourceRow.Cells)
        {
            var destinationCell = destinationRow.CreateCell(sourceCell.ColumnIndex);
            switch (sourceCell.CellType)
            {
                case CellType.Formula:
                    destinationCell.SetCellFormula(sourceCell.CellFormula);
                    break;
                case CellType.Numeric:
                    destinationCell.SetCellValue(sourceCell.NumericCellValue);
                    break;
                case CellType.Boolean:
                    destinationCell.SetCellValue(sourceCell.BooleanCellValue);
                    break;
                case CellType.Error:
                    destinationCell.SetCellErrorValue(sourceCell.ErrorCellValue);
                    break;
                case CellType.String:
                    destinationCell.SetCellValue(sourceCell.StringCellValue);
                    break;
            }
            if (sourceCell.CellStyle != null)
            {
                if (!styleCache.TryGetValue(sourceCell.CellStyle.Index, out var style))
                {
                    style = destination.CreateCellStyle();
                    style.CloneStyleFrom(sourceCell.CellStyle);
                    styleCache[sourceCell.CellStyle.Index] = style;
                }
                destinationCell.CellStyle = style;
            }
        }
    }

    private static void AddFailureColumns(ISheet sheet, IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<int, int> rowMap, short sourceColumn)
    {
        var startColumn = Math.Max(0, (int)sourceColumn);
        var header = sheet.GetRow(0) ?? sheet.CreateRow(0);
        header.CreateCell(startColumn).SetCellValue("__SourceSheet");
        header.CreateCell(startColumn + 1).SetCellValue("__SourceRow");
        header.CreateCell(startColumn + 2).SetCellValue("__ErrorCode");
        header.CreateCell(startColumn + 3).SetCellValue("__Errors");
        foreach (var group in errors.Where(error => error.RowIndex > 1).GroupBy(error => error.RowIndex - 1))
        {
            if (!rowMap.TryGetValue(group.Key, out var targetRow))
                continue;
            var row = sheet.GetRow(targetRow) ?? sheet.CreateRow(targetRow);
            row.CreateCell(startColumn).SetCellValue(group.First().SheetName ?? string.Empty);
            row.CreateCell(startColumn + 1).SetCellValue(group.Key + 1);
            row.CreateCell(startColumn + 2).SetCellValue(string.Join(" | ", group.Select(error => error.Code)));
            row.CreateCell(startColumn + 3).SetCellValue(string.Join(" | ", group.Select(error => error.Message)
                .Where(message => !string.IsNullOrWhiteSpace(message))));
        }
    }

    private static void CopyMergedRegions(ISheet source, ISheet destination,
        IReadOnlyDictionary<int, int> rowMap)
    {
        foreach (var region in source.MergedRegions)
        {
            if (!rowMap.TryGetValue(region.FirstRow, out var firstRow)
                || !rowMap.TryGetValue(region.LastRow, out var lastRow))
                continue;
            var contiguous = true;
            for (var row = region.FirstRow; row <= region.LastRow; row++)
                contiguous &= rowMap.ContainsKey(row);
            if (contiguous)
                destination.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(firstRow, lastRow,
                    region.FirstColumn, region.LastColumn));
        }
    }

    private static void CopyDataValidations(ISheet source, ISheet destination,
        IReadOnlyDictionary<int, int> rowMap, CancellationToken cancellationToken)
    {
        var helper = destination.GetDataValidationHelper();
        foreach (var validation in source.GetDataValidations())
        {
            foreach (var region in validation.Regions.CellRangeAddresses)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!rowMap.TryGetValue(region.FirstRow, out var firstRow)
                    || !rowMap.TryGetValue(region.LastRow, out var lastRow))
                    continue;
                var contiguous = true;
                for (var row = region.FirstRow; row <= region.LastRow; row++)
                    contiguous &= rowMap.ContainsKey(row);
                if (!contiguous)
                    continue;
                var targetRegion = new NPOI.SS.Util.CellRangeAddressList(firstRow, lastRow,
                    region.FirstColumn, region.LastColumn);
                destination.AddValidationData(helper.CreateValidation(validation.ValidationConstraint,
                    targetRegion));
            }
        }
    }

    private static void CopyPictures(ISheet source, ISheet destination,
        IReadOnlyDictionary<int, int> rowMap, CancellationToken cancellationToken)
    {
        foreach (var picture in source.GetAllPictureInfos())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sourceLastRow = Math.Max(picture.MinRow, picture.MaxRow - 1);
            if (!rowMap.TryGetValue(picture.MinRow, out var minRow)
                || !rowMap.TryGetValue(sourceLastRow, out var mappedLastRow))
                continue;
            var contiguous = true;
            for (var row = picture.MinRow; row <= sourceLastRow; row++)
                contiguous &= rowMap.ContainsKey(row);
            if (!contiguous)
                continue;
            destination.AddPicture(new PictureInfo(minRow, mappedLastRow + 1, picture.MinCol, picture.MaxCol,
                picture.PictureData, picture.PictureStyle ?? new PictureStyle()));
        }
    }

    private static string FormatRawValue(object value)
    {
        if (value is byte[] bytes)
            return $"<binary:{bytes.Length}>";
        return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
    }

    private static void AnnotateErrors(IWorkbook workbook, IReadOnlyCollection<ExcelImportError> errors)
    {
        foreach (var error in errors.Where(error => error.RowIndex > 0 && error.ColumnIndex > 0))
        {
            var sheet = workbook.GetSheet(error.SheetName);
            var cell = sheet?.GetRow(error.RowIndex - 1)?.GetCell(error.ColumnIndex - 1);
            if (cell == null)
                continue;
            var existing = cell.CellComment;
            var text = error.Message ?? string.Empty;
            if (existing != null)
                text = existing.String.String + Environment.NewLine + text;
            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = cell.ColumnIndex;
            anchor.Col2 = cell.ColumnIndex + 2;
            anchor.Row1 = cell.RowIndex;
            anchor.Row2 = cell.RowIndex + 3;
            var comment = sheet.CreateDrawingPatriarch().CreateCellComment(anchor);
            comment.String = workbook.GetCreationHelper().CreateRichTextString(text);
            comment.Author = "Bing.Offices";
            cell.CellComment = comment;
        }
    }

    private static void WriteBytes(Stream destination, byte[] bytes, CancellationToken cancellationToken)
    {
        if (destination == null || !destination.CanWrite)
            throw new ArgumentException("失败工作簿目标流不可写入。", nameof(destination));
        var offset = 0;
        while (offset < bytes.Length)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var count = Math.Min(81920, bytes.Length - offset);
            destination.Write(bytes, offset, count);
            offset += count;
        }
    }

    /// <summary>
    /// 单遍处理工作表中的数据行。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="sheet">NPOI 工作表。</param>
    /// <param name="options">导入选项。</param>
    /// <param name="items">成功项集合。</param>
    /// <param name="errors">错误集合。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="sourceRows">成功实体对应的 zero-based 原始行集合。</param>
    private void ImportSheet<T>(ISheet sheet, ExcelImportExecutionOptions<T> options, ICollection<T> items,
        ExcelImportErrorCollector errors, ExcelImportRuntime runtime, CancellationToken cancellationToken,
        ICollection<int> sourceRows = null)
        where T : class, new()
    {
        var header = sheet.GetRow(options.HeaderRowIndex)
            ?? throw new SheetStructureException("导入的模板不正确，未匹配表头。");
        if (header.LastCellNum > options.MaxColumnLength)
            throw new SheetStructureException($"导入表头超过最大列长度: {options.MaxColumnLength}");

        var columns = CreateColumns<T>(header, options);
        if (runtime.RowLimitReached)
            return;
        IReadOnlyDictionary<(int Row, int Column), IReadOnlyList<PictureInfo>> imageIndex = null;
        if (columns.Values.Any(IsImageColumn))
        {
            try
            {
                imageIndex = BuildImageIndex(sheet, runtime.ImageResources, cancellationToken);
            }
            catch (ImageResourceLimitException exception)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit, exception.Message,
                    sheet.SheetName, options.HeaderRowIndex + 1, 0, null));
                return;
            }
        }
        var workbookValidations = options.ValidationMode == ExcelImportValidationMode.WorkbookRules
            || options.ValidationMode == ExcelImportValidationMode.ConfiguredAndWorkbook
            ? sheet.GetDataValidations().ToArray()
            : Array.Empty<IDataValidation>();
        var validationIndex = ValidationRangeIndex.Create(workbookValidations, options.DataRowIndex, sheet.LastRowNum,
            0, Math.Max(0, header.LastCellNum - 1));
        var duplicateValues = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        for (var rowIndex = options.DataRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
            if (!runtime.TryConsumeRow())
            {
                if (runtime.TryMarkRowLimitReported())
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                        $"Workbook 数据行数超过限制: {runtime.MaxRows}", sheet.SheetName, rowIndex + 1, 0, null));
                break;
            }
            var row = sheet.GetRow(rowIndex);
            if (IsEmpty(row, options.BodyWhitespace, imageIndex, rowIndex))
            {
                if (options.IgnoreEmptyLineAfterData)
                    break;
                if (options.EnabledEmptyLine)
                {
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidInput, "导入数据存在空行", sheet.SheetName,
                        rowIndex + 1, 0, null));
                    if (errors.IsLimitReached)
                    {
                        errors.MarkTruncated();
                        break;
                    }
                }
                continue;
            }
            var duplicateChanges = CaptureDuplicateChanges(row, columns, duplicateValues, options.BodyWhitespace);
            if (!ValidateWorkbookValues(row, columns, validationIndex, sheet, sheet.SheetName, rowIndex,
                options.BodyWhitespace, options.ValidateMode, options.UnsupportedFeaturePolicy, errors))
            {
                RollbackDuplicateChanges(duplicateValues, duplicateChanges);
                if (errors.IsLimitReached)
                {
                    errors.MarkTruncated();
                    break;
                }
                continue;
            }
            if (!ValidateRawValues(row, columns, duplicateValues, sheet.SheetName, rowIndex, options.ValidateMode,
                    options.Culture, options.BodyWhitespace, errors))
            {
                RollbackDuplicateChanges(duplicateValues, duplicateChanges);
                if (errors.IsLimitReached)
                {
                    errors.MarkTruncated();
                    break;
                }
                continue;
            }
            if (TryCreateItem(row, columns, duplicateValues, sheet.SheetName, rowIndex, options.ValidateMode, errors,
                    options.Culture, options.BodyWhitespace, options.DynamicTargetGetter, imageIndex, out T item))
            {
                items.Add(item);
                sourceRows?.Add(rowIndex);
            }
            else
            {
                RollbackDuplicateChanges(duplicateValues, duplicateChanges);
            }
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
        }
    }

    private static bool ValidateWorkbookValues(IRow row, IReadOnlyDictionary<int, ExcelColumnPlan> columns,
        ValidationRangeIndex validationIndex, ISheet sheet, string sheetName, int rowIndex,
        ExcelWhitespacePolicy bodyWhitespace, ValidateMode validateMode,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy, ExcelImportErrorCollector errors)
    {
        var valid = true;
        foreach (var column in columns)
        {
            var cell = row.GetCell(column.Key);
            var cellValue = ReadCellValue(cell);
            var value = NormalizeText(cellValue.Text, bodyWhitespace);
            foreach (var validation in validationIndex.Get(rowIndex, column.Key))
            {
                if (ValidateWorkbookValue(validation.ValidationConstraint, cellValue, value, sheet,
                    out var message))
                    continue;
                errors.Add(new ExcelImportError(ExcelImportErrorCode.WorkbookValidation, message, sheetName,
                    rowIndex + 1, column.Key + 1, column.Value.Property.Name, column.Value.DynamicDefinition?.Key
                    ?? column.Value.Property.Name, column.Value.HeaderName, value));
                if (unsupportedFeaturePolicy == ExcelUnsupportedFeaturePolicy.Report
                    && message != null && message.StartsWith("Workbook Data Validation 规则类型或公式暂不支持",
                        StringComparison.Ordinal))
                    continue;
                valid = false;
                if (validateMode == ValidateMode.StopOnFirstFailure)
                    return false;
            }
        }
        return valid;
    }

    private static bool ValidateWorkbookValue(IDataValidationConstraint constraint, ExcelCellValue cellValue,
        string value, ISheet sheet, out string message)
    {
        value ??= string.Empty;
        var type = constraint.GetValidationType();
        if (type == NPOI.SS.UserModel.ValidationType.ANY)
        {
            message = null;
            return true;
        }
        if (type == NPOI.SS.UserModel.ValidationType.LIST)
        {
            if (!TryGetExplicitListValues(constraint, sheet, out var values))
                return UnsupportedWorkbookValidation(out message);
            if (values.Contains(value, StringComparer.Ordinal))
            {
                message = null;
                return true;
            }
            message = "不符合 Workbook 显式列表校验。";
            return false;
        }
        if (type == NPOI.SS.UserModel.ValidationType.TEXT_LENGTH)
        {
            if (!decimal.TryParse(constraint.Formula1, NumberStyles.Float, CultureInfo.InvariantCulture,
                    out var first))
                return UnsupportedWorkbookValidation(out message);
            var length = value.Length;
            var secondParsed = decimal.TryParse(constraint.Formula2, NumberStyles.Float,
                CultureInfo.InvariantCulture, out var second);
            var valid = ValidateComparison(length, first, second, secondParsed, constraint.Operator);
            message = valid ? null : "不符合 Workbook 文本长度校验。";
            return valid;
        }
        if (type == NPOI.SS.UserModel.ValidationType.INTEGER || type == NPOI.SS.UserModel.ValidationType.DECIMAL)
        {
            var parsed = decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture,
                out var number);
            var firstParsed = decimal.TryParse(constraint.Formula1, NumberStyles.Number,
                CultureInfo.InvariantCulture, out var first);
            var secondParsed = decimal.TryParse(constraint.Formula2, NumberStyles.Number,
                CultureInfo.InvariantCulture, out var second);
            var valid = parsed && firstParsed && ValidateComparison(number, first, second, secondParsed,
                constraint.Operator);
            message = valid ? null : "不符合 Workbook 数值校验。";
            return valid;
        }
        if (type == NPOI.SS.UserModel.ValidationType.DATE || type == NPOI.SS.UserModel.ValidationType.TIME)
        {
            var parsed = TryGetExcelDate(cellValue, value, type == NPOI.SS.UserModel.ValidationType.TIME, out var date);
            var first = TryGetExcelDate(constraint.Formula1, type == NPOI.SS.UserModel.ValidationType.TIME,
                out var minimum);
            var second = TryGetExcelDate(constraint.Formula2, type == NPOI.SS.UserModel.ValidationType.TIME,
                out var maximum);
            var valid = parsed && first && ValidateComparison(date, minimum, maximum, second,
                constraint.Operator);
            message = valid ? null : "不符合 Workbook 日期/时间校验。";
            return valid;
        }
        return UnsupportedWorkbookValidation(out message);
    }

    private static bool TryGetExplicitListValues(IDataValidationConstraint constraint, ISheet currentSheet,
        out string[] values)
    {
        var formula = constraint.Formula1?.Replace("&#34;", "\"").Replace("&quot;", "\"")
            .Replace("\\\"", "\"").Trim();
        if (!string.IsNullOrWhiteSpace(formula)
            && TryGetCellRangeValues(formula, currentSheet, out values))
            return true;
        if (string.IsNullOrWhiteSpace(formula))
        {
            values = null;
            return false;
        }
        values = constraint.ExplicitListValues;
        if (values != null && values.Length > 0)
        {
            if (values.Length == 1 && (values[0].Contains(":") || values[0].Contains("!")))
            {
                values = null;
                return false;
            }
            if (values.Length == 1 && values[0].Contains(","))
                values = SplitExplicitList(values[0]);
            return values.Length > 0;
        }
        if (formula.StartsWith("=", StringComparison.Ordinal))
            return false;
        if (!(formula.StartsWith("\"", StringComparison.Ordinal) && formula.EndsWith("\"",
                StringComparison.Ordinal)))
            return false;
        values = SplitExplicitList(formula.Trim('\"'));
        return values.Length > 0;
    }

    private static bool TryGetCellRangeValues(string formula, ISheet currentSheet, out string[] values)
    {
        values = null;
        var expression = formula.Trim();
        if (expression.StartsWith("=", StringComparison.Ordinal))
            expression = expression.Substring(1).Trim();
        var separator = expression.LastIndexOf('!');
        var sheetName = separator < 0 ? null : expression.Substring(0, separator).Trim('\'');
        var range = separator < 0 ? expression : expression.Substring(separator + 1);
        var parts = range.Split(':');
        if (parts.Length != 2)
            return false;
        ISheet sheet = currentSheet;
        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            var workbook = currentSheet.Workbook;
            sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
                return false;
        }
        try
        {
            var first = new NPOI.SS.Util.CellReference(parts[0].Trim());
            var last = new NPOI.SS.Util.CellReference(parts[1].Trim());
            var result = new List<string>();
            for (var rowIndex = first.Row; rowIndex <= last.Row; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                for (var columnIndex = first.Col; columnIndex <= last.Col; columnIndex++)
                    result.Add(GetRawStringValue(row?.GetCell(columnIndex)));
            }
            values = result.ToArray();
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    private static string[] SplitExplicitList(string value) => value.Trim().Trim('\"')
        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
        .Select(item => item.Trim()).ToArray();

    private static bool TryGetExcelDate(ExcelCellValue cellValue, string text, bool timeOnly, out long ticks)
    {
        if (cellValue?.Value is DateTime date)
        {
            ticks = timeOnly ? date.TimeOfDay.Ticks : date.Ticks;
            return true;
        }
        if (cellValue?.Value is double number)
        {
            ticks = ExcelSerialToDateTime(number, timeOnly).Ticks;
            return true;
        }
        return TryGetExcelDate(text, timeOnly, out ticks);
    }

    private static bool TryGetExcelDate(string text, bool timeOnly, out long ticks)
    {
        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces,
                out var date))
        {
            ticks = timeOnly ? date.TimeOfDay.Ticks : date.Ticks;
            return true;
        }
        if (decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
        {
            ticks = ExcelSerialToDateTime((double)serial, timeOnly).Ticks;
            return true;
        }
        ticks = 0;
        return false;
    }

    private static DateTime ExcelSerialToDateTime(double serial, bool timeOnly)
    {
        var date = new DateTime(1899, 12, 30).AddDays(serial);
        return timeOnly ? DateTime.Today.Add(date.TimeOfDay) : date;
    }

    private static bool ValidateComparison(long value, long first, long second, bool secondParsed, int operation) => operation switch
    {
        0 => secondParsed && value >= first && value <= second,
        1 => secondParsed && (value < first || value > second),
        2 => value == first,
        3 => value != first,
        4 => value > first,
        5 => value < first,
        6 => value >= first,
        7 => secondParsed && value <= second,
        _ => false,
    };

    private static bool ValidateComparison(decimal value, decimal first, decimal second, bool secondParsed,
        int operation) => operation switch
    {
        0 => secondParsed && value >= first && value <= second,
        1 => secondParsed && (value < first || value > second),
        2 => value == first,
        3 => value != first,
        4 => value > first,
        5 => value < first,
        6 => value >= first,
        7 => secondParsed && value <= second,
        _ => false
    };

    private static bool UnsupportedWorkbookValidation(out string message)
    {
        message = "Workbook Data Validation 规则类型或公式暂不支持。";
        return false;
    }

    /// <summary>
    /// 建立当前工作表的列绑定。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="header">表头行。</param>
    /// <param name="options">导入选项。</param>
    /// <returns>按列索引访问的属性绑定。</returns>
    private IReadOnlyDictionary<int, ExcelColumnPlan> CreateColumns<T>(IRow header, ExcelImportExecutionOptions<T> options)
        where T : class, new()
    {
        var map = ExcelTypeMapFactory.Get(options.MappingProfile, options.MappingConfiguration);
        var dynamicProperties = map.Properties.Where(property => property.IsDynamicColumn).ToList();
        if (dynamicProperties.Count > 1)
            throw new InvalidOperationException($"导入模板 {typeof(T).FullName} 只能声明一个动态列属性。");

        var fixedProperties = map.Properties.Where(property => !property.Ignored && !property.IsDynamicColumn).ToList();
        var headerNames = new HashSet<string>(options.HeaderComparison == ExcelNameComparison.Ordinal
            ? StringComparer.Ordinal
            : StringComparer.OrdinalIgnoreCase);
        var columns = new Dictionary<int, ExcelColumnPlan>();
        foreach (var headerCell in header.Cells)
        {
            if (options.ReadColumnRange != null && !options.ReadColumnRange.Contains(headerCell.ColumnIndex))
                continue;
            var headerName = NormalizeText(GetRawStringValue(headerCell), options.HeaderWhitespace);
            if (string.IsNullOrWhiteSpace(headerName))
                continue;
            if (!headerNames.Add(headerName))
                throw new SheetStructureException($"导入的表格存在重复列:{headerName}");
            var property = FindProperty(fixedProperties, headerName, null, options.HeaderComparison);
            ExcelDynamicColumnDefinition dynamicDefinition = null;
            if (property == null && dynamicProperties.Count == 1)
            {
                dynamicDefinition = FindDynamicDefinition(headerName, options.DynamicColumns, options.HeaderComparison);
                if (options.DynamicColumns != null && options.DynamicColumns.Count > 0 && dynamicDefinition == null)
                {
                    if (options.FailOnUnknownDynamicColumns)
                        throw new SheetStructureException($"导入包含未知动态列: {headerName}");
                    continue;
                }
                property = dynamicProperties[0];
            }
            if (property != null)
            {
                if (!property.Property.CanWrite)
                    throw new InvalidOperationException($"导入模板属性不可写入: {property.Name}");
                var isUnspecifiedDynamicColumn = property.IsDynamicColumn && dynamicDefinition == null;
                var valueType = dynamicDefinition?.DataType ?? property.Property.PropertyType;
                var valueConverters = isUnspecifiedDynamicColumn
                    ? Array.Empty<IExcelValueConverter>()
                    : BindValueConverters(dynamicDefinition?.ConverterName ?? property.ConverterName, valueType);
                var attributeValidations = BindAttributeValidations(property);
                var namedValidationRules = BindNamedValidationRules(property, dynamicDefinition);
                columns[headerCell.ColumnIndex] = new ExcelColumnPlan(headerName, property, property.IsDynamicColumn,
                    headerCell.ColumnIndex, dynamicDefinition, null, valueConverters, attributeValidations,
                    namedValidationRules);
            }
        }

        if (options.HeaderMatch)
        {
            var missing = fixedProperties.Where(property => !columns.Values.Any(column => column.Property == property)
                && (options.ReadColumnRange == null || !header.Cells.Any(cell =>
                    !options.ReadColumnRange.Contains(cell.ColumnIndex)
                    && string.Equals(NormalizeText(GetRawStringValue(cell), options.HeaderWhitespace), property.Title,
                        ToStringComparison(options.HeaderComparison))))).ToList();
            if (missing.Any())
                throw new SheetStructureException($"导入的表格不存在列：{string.Join(",", missing.Select(property => property.Title))}");
        }
        return columns;
    }

    /// <summary>
    /// 按展示标题或历史别名查找 typed 动态列定义。
    /// </summary>
    private static ExcelDynamicColumnDefinition FindDynamicDefinition(string headerName,
        IReadOnlyList<ExcelDynamicColumnDefinition> definitions, ExcelNameComparison comparison)
    {
        return (definitions ?? Array.Empty<ExcelDynamicColumnDefinition>()).FirstOrDefault(definition =>
            string.Equals(definition.Title, headerName, ToStringComparison(comparison))
            || (definition.Aliases ?? Array.Empty<string>()).Any(alias =>
                string.Equals(alias, headerName, ToStringComparison(comparison))));
    }

    /// <summary>
    /// 验证请求级表头映射不会覆盖动态列绑定规则。
    /// </summary>
    /// <param name="headerMappings">属性名称到表头名称的请求级映射。</param>
    /// <param name="fixedProperties">可映射的固定属性集合。</param>
    /// <param name="dynamicProperties">动态列属性集合。</param>
    private static void ValidateHeaderMappings(IReadOnlyDictionary<string, string> headerMappings,
        IReadOnlyCollection<ExcelPropertyMap> fixedProperties, IReadOnlyCollection<ExcelPropertyMap> dynamicProperties)
    {
        if (headerMappings == null)
            return;
        var propertyNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var headerNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var mapping in headerMappings)
        {
            if (string.IsNullOrWhiteSpace(mapping.Key) || string.IsNullOrWhiteSpace(mapping.Value))
                throw new InvalidOperationException("导入模板表头映射的属性名称和表头名称不能为空。");
            if (dynamicProperties.Any(property => string.Equals(property.Name, mapping.Key,
                    StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException($"导入模板不允许为动态列属性配置表头映射: {mapping.Key}");
            if (!fixedProperties.Any(property => string.Equals(property.Name, mapping.Key,
                    StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException($"导入模板表头映射引用了不存在或已忽略的属性: {mapping.Key}");
            if (!propertyNames.Add(mapping.Key))
                throw new InvalidOperationException($"导入模板表头映射包含重复的属性: {mapping.Key}");
            if (!headerNames.Add(mapping.Value))
                throw new InvalidOperationException($"导入模板表头映射包含重复的表头: {mapping.Value}");
        }
    }

    /// <summary>
    /// 根据请求映射和默认标题查找固定属性。
    /// </summary>
    /// <param name="properties">固定属性集合。</param>
    /// <param name="headerName">表头名称。</param>
    /// <param name="headerMappings">请求级映射。</param>
    private static ExcelPropertyMap FindProperty(IEnumerable<ExcelPropertyMap> properties, string headerName,
        IReadOnlyDictionary<string, string> headerMappings, ExcelNameComparison comparison)
    {
        var stringComparison = ToStringComparison(comparison);
        foreach (var property in properties)
        {
            var mappedHeader = headerMappings?.FirstOrDefault(mapping =>
                string.Equals(mapping.Key, property.Name, StringComparison.OrdinalIgnoreCase)).Value;
            if (!string.IsNullOrWhiteSpace(mappedHeader)
                && string.Equals(mappedHeader, headerName, stringComparison))
                return property;
            if (string.Equals(property.Title, headerName, stringComparison)
                || string.Equals(property.Name, headerName, stringComparison))
                return property;
        }
        return null;
    }

    private static StringComparison ToStringComparison(ExcelNameComparison comparison) =>
        comparison == ExcelNameComparison.Ordinal ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

    private IReadOnlyList<ExcelAttributeValidationBinding> BindAttributeValidations(ExcelPropertyMap property)
    {
        return property.Property.GetCustomAttributes<FilterAttributeBase>()
            .Select(attribute => new ExcelAttributeValidationBinding(attribute, ResolveValidationRule(attribute),
                IsRawValidationAttribute(attribute)))
            .ToArray();
    }

    private IReadOnlyList<INamedExcelValidationRule> BindNamedValidationRules(ExcelPropertyMap property,
        ExcelDynamicColumnDefinition dynamicDefinition)
    {
        var names = new List<string>(property.ValidationRuleNames ?? Array.Empty<string>());
        if (!string.IsNullOrWhiteSpace(dynamicDefinition?.ValidatorName))
            names.Add(dynamicDefinition.ValidatorName);
        return names.Select(ResolveNamedValidationRule).ToArray();
    }

    private static string GetRawStringValue(ICell cell)
    {
        if (cell == null)
            return string.Empty;
        var cellType = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
        return cellType == CellType.String ? cell.StringCellValue ?? string.Empty : cell.GetStringValue();
    }

    private static ExcelCellValue NormalizeCellValue(ExcelCellValue cellValue, ExcelWhitespacePolicy policy)
    {
        if (cellValue == null || cellValue.Kind != ExcelCellKind.Text && cellValue.Kind != ExcelCellKind.Formula)
            return cellValue;
        var text = NormalizeText(cellValue.Text, policy);
        return text == cellValue.Text ? cellValue : new ExcelCellValue(cellValue.Value, text, cellValue.Kind,
            cellValue.CachedKind, cellValue.Formula, cellValue.ErrorCode, cellValue.FormatIndex);
    }

    private static string NormalizeText(string value, ExcelWhitespacePolicy policy)
    {
        value ??= string.Empty;
        return policy switch
        {
            ExcelWhitespacePolicy.Preserve => value,
            ExcelWhitespacePolicy.Trim => value.Trim(),
            ExcelWhitespacePolicy.RemoveAll => RemoveWhitespace(value),
            _ => throw new ArgumentOutOfRangeException(nameof(policy))
        };
    }

    private static string RemoveWhitespace(string value)
    {
        var hasWhitespace = value.Any(char.IsWhiteSpace);
        if (!hasWhitespace)
            return value;
        var buffer = new char[value.Length];
        var length = 0;
        foreach (var character in value)
        {
            if (!char.IsWhiteSpace(character))
                buffer[length++] = character;
        }
        return new string(buffer, 0, length);
    }

    /// <summary>
    /// 将单行数据绑定为实体。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="row">NPOI 数据行。</param>
    /// <param name="columns">列绑定。</param>
    /// <param name="duplicateValues">当前工作表的重复值状态。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">从零开始的行索引。</param>
    /// <param name="validateMode">错误收集模式。</param>
    /// <param name="errors">错误集合。</param>
    /// <param name="culture">值转换使用的区域性。</param>
    /// <param name="item">成功绑定的实体。</param>
    /// <returns>绑定成功时返回 <see langword="true"/>。</returns>
    private bool TryCreateItem<T>(IRow row, IReadOnlyDictionary<int, ExcelColumnPlan> columns,
        IDictionary<string, HashSet<string>> duplicateValues, string sheetName, int rowIndex, ValidateMode validateMode,
        ExcelImportErrorCollector errors, CultureInfo culture, ExcelWhitespacePolicy bodyWhitespace,
        Func<object, object> dynamicTargetGetter,
        IReadOnlyDictionary<(int Row, int Column), IReadOnlyList<PictureInfo>> imageIndex, out T item)
        where T : class, new()
    {
        item = new T();
        Dictionary<string, object> dynamicValues = null;
        foreach (var column in columns)
        {
            var cellValue = default(ExcelCellValue);
            try
            {
                var cell = row.GetCell(column.Key);
                var images = imageIndex == null ? null : FindImages(imageIndex, rowIndex, column.Key);
                    var imageMultiplicity = column.Value.ImageMultiplicity;
                if (images != null && images.Count > 1
                    && imageMultiplicity == ExcelImageMultiplicityPolicy.Fail)
                {
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidInput, "同一单元格存在多个图片。",
                        sheetName, rowIndex + 1, column.Key + 1, column.Value.Property.Name));
                    item = null;
                    return false;
                }
                var image = images?.FirstOrDefault();
                cellValue = image == null
                    ? NormalizeCellValue(ApplyLegacyTextConverters(cell, ReadCellValue(cell)), bodyWhitespace)
                    : new ExcelCellValue(image, string.Empty, ExcelCellKind.Empty);
                var value = cellValue.Text;
                if (column.Value.IsDynamic)
                {
                    dynamicValues ??= new Dictionary<string, object>(StringComparer.Ordinal);
                    object dynamicConvertedValue;
                    if (column.Value.DynamicDefinition == null)
                        dynamicConvertedValue = value;
                    else
                    {
                        var propertyType = column.Value.ValueType;
                        if (image != null && IsImageType(propertyType))
                            dynamicConvertedValue = ConvertImages(images, propertyType,
                                column.Value.ImageMultiplicity);
                        else
                            dynamicConvertedValue = column.Value.ConvertFrom(value, cellValue, sheetName, rowIndex + 1,
                                column.Key + 1, culture);
                    }
                    if (!ValidateColumnValue(value, cellValue, dynamicConvertedValue, column.Value,
                        duplicateValues, sheetName, rowIndex, validateMode, culture, errors))
                    {
                        item = null;
                        return false;
                    }
                    dynamicValues[column.Value.Key] = dynamicConvertedValue;
                    continue;
                }
                var converted = image != null && IsImageType(column.Value.ValueType)
                    ? ConvertImages(images, column.Value.ValueType,
                        imageMultiplicity)
                    : column.Value.ConvertFrom(value, cellValue, sheetName, rowIndex + 1, column.Key + 1, culture);
                if (!ValidateColumnValue(value, cellValue, converted, column.Value, duplicateValues, sheetName,
                    rowIndex, validateMode, culture, errors))
                {
                    item = null;
                    return false;
                }
                column.Value.Setter(item, converted);
            }
            catch (Exception exception)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.ValueConversion, exception.Message, sheetName,
                    rowIndex + 1, column.Key + 1, column.Value.Property.Name, column.Value.Property.Name,
                    column.Value.HeaderName, cellValue?.Value ?? cellValue?.Text));
                item = null;
                return false;
            }
        }
        if (dynamicValues != null)
        {
            var target = dynamicTargetGetter?.Invoke(item) as IDictionary<string, object>;
            if (target != null)
            {
                foreach (var pair in dynamicValues)
                    target[pair.Key] = pair.Value;
            }
            else
                columns.Values.First(column => column.IsDynamic).Setter(item, dynamicValues);
        }
        return true;
    }

    /// <summary>
    /// 读取不依赖 NPOI 的单元格值描述。
    /// </summary>
    /// <param name="cell">待读取的单元格。</param>
    /// <returns>用于转换器和默认转换的单元格值描述。</returns>
    private static ExcelCellValue ReadCellValue(ICell cell)
    {
        if (cell == null)
            return new ExcelCellValue(null, string.Empty, ExcelCellKind.Empty);

        var isFormula = cell.CellType == CellType.Formula;
        var effectiveType = isFormula ? cell.CachedFormulaResultType : cell.CellType;
        var effectiveKind = ResolveCellKind(effectiveType, cell);
        object value = effectiveType switch
        {
            CellType.Boolean => cell.BooleanCellValue,
            CellType.Numeric when DateUtil.IsCellDateFormatted(cell) => cell.DateCellValue,
            CellType.Numeric => cell.NumericCellValue,
            CellType.Error => null,
            _ => GetRawStringValue(cell)
        };
        return new ExcelCellValue(value, GetRawStringValue(cell), isFormula ? ExcelCellKind.Formula : effectiveKind,
            isFormula ? effectiveKind : null, isFormula ? cell.CellFormula : null,
            effectiveType == CellType.Error ? cell.ErrorCellValue : null, cell.CellStyle?.DataFormat);
    }

    private static bool IsImageColumn(ExcelColumnPlan column) => IsImageType(column.ValueType);

    private static bool IsImageType(Type type)
    {
        if (type == null || type == typeof(byte[]) || type == typeof(ExcelImageData))
            return type != null;
        if (type.IsArray)
            return type.GetElementType() == typeof(byte[]) || type.GetElementType() == typeof(ExcelImageData);
        if (type.IsGenericType && type.GetGenericArguments().Length == 1)
        {
            var elementType = type.GetGenericArguments()[0];
            if (elementType == typeof(byte[]) || elementType == typeof(ExcelImageData))
                return true;
        }
        return GetImageElementType(type) != null;
    }

    private static IReadOnlyDictionary<(int Row, int Column), IReadOnlyList<PictureInfo>> BuildImageIndex(ISheet sheet,
        ExcelImageResourceTracker resources, CancellationToken cancellationToken)
    {
        var result = new Dictionary<(int Row, int Column), IReadOnlyList<PictureInfo>>();
        foreach (var picture in sheet.GetAllPictureInfos())
        {
            cancellationToken.ThrowIfCancellationRequested();
            resources.Consume(picture.PictureData.LongLength);
            var key = (picture.MinRow, picture.MinCol);
            if (!result.TryGetValue(key, out var pictures))
                result[key] = pictures = new List<PictureInfo>();
            ((List<PictureInfo>)pictures).Add(picture);
        }
        return result;
    }

    private static IReadOnlyList<PictureInfo> FindImages(
        IReadOnlyDictionary<(int Row, int Column), IReadOnlyList<PictureInfo>> imageIndex, int row, int column)
    {
        return imageIndex.TryGetValue((row, column), out var pictures) ? pictures : null;
    }

    private static object ConvertImages(IReadOnlyList<PictureInfo> pictures, Type targetType,
        ExcelImageMultiplicityPolicy policy)
    {
        if (policy == ExcelImageMultiplicityPolicy.All)
        {
            var elementType = GetImageElementType(targetType);
            if (elementType == typeof(byte[]))
            {
                var values = pictures.Select(picture => picture.PictureData).ToArray();
                return targetType.IsArray ? values : values.ToList();
            }
            if (elementType == typeof(ExcelImageData))
            {
                var values = pictures.Select(picture => (ExcelImageData)ConvertImage(picture,
                    typeof(ExcelImageData))).ToList();
                return targetType.IsArray ? values.ToArray() : values;
            }
        }
        return ConvertImage(pictures[0], targetType);
    }

    private static Type GetImageElementType(Type type)
    {
        if (type == null)
            return null;
        if (type.IsArray)
            return type.GetElementType();
        if (type.IsGenericType && type.GetGenericArguments().Length == 1)
        {
            var directElementType = type.GetGenericArguments()[0];
            if (directElementType == typeof(byte[]) || directElementType == typeof(ExcelImageData))
                return directElementType;
        }
        return type.GetInterfaces().Concat(new[] { type })
            .Where(candidate => candidate.IsGenericType
                && candidate.GetGenericTypeDefinition() == typeof(IEnumerable<>))
            .Select(candidate => candidate.GetGenericArguments()[0])
            .FirstOrDefault(candidate => candidate == typeof(byte[]) || candidate == typeof(ExcelImageData));
    }

    private static object ConvertImage(PictureInfo picture, Type targetType)
    {
        if (targetType == typeof(byte[]))
            return picture.PictureData;
        return new ExcelImageData(picture.PictureData, ResolveImageContentType(picture.PictureData), picture.MinRow + 1,
            picture.MinCol + 1, picture.MaxRow + 1, picture.MaxCol + 1);
    }

    private static string ResolveImageContentType(byte[] bytes)
    {
        if (bytes?.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
            return "image/jpeg";
        if (bytes?.Length >= 6 && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46)
            return "image/gif";
        return "image/png";
    }

    /// <summary>
    /// 映射 NPOI 单元格类型到提供程序无关的逻辑类型。
    /// </summary>
    private static ExcelCellKind ResolveCellKind(CellType cellType, ICell cell) => cellType switch
    {
        CellType.Blank => ExcelCellKind.Empty,
        CellType.String => ExcelCellKind.Text,
        CellType.Boolean => ExcelCellKind.Boolean,
        CellType.Error => ExcelCellKind.Error,
        CellType.Numeric when DateUtil.IsCellDateFormatted(cell) => ExcelCellKind.DateTime,
        CellType.Numeric => ExcelCellKind.Number,
        _ => ExcelCellKind.Text
    };

    /// <summary>
    /// 将旧版文本转换器限制在文本覆盖范围内，保留新的 typed cell 语义。
    /// </summary>
    private ExcelCellValue ApplyLegacyTextConverters(ICell cell, ExcelCellValue cellValue)
    {
        if (cell == null || _legacyValueConverters.Count == 0)
            return cellValue;
        var text = cellValue.Text;
        foreach (var converter in _legacyValueConverters)
        {
            var converted = converter.GetStringValue(cell);
            if (converted != null)
                text = converted;
        }
        return text == cellValue.Text
            ? cellValue
            : new ExcelCellValue(cellValue.Value, text, cellValue.Kind, cellValue.CachedKind, cellValue.Formula,
                cellValue.ErrorCode, cellValue.FormatIndex);
    }

    /// <summary>
    /// 按属性特性校验单行数据。
    /// </summary>
    /// <param name="row">NPOI 数据行。</param>
    /// <param name="columns">列绑定。</param>
    /// <param name="duplicateValues">当前工作表的已出现重复校验值。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">从零开始的行索引。</param>
    /// <param name="validateMode">错误收集模式。</param>
    /// <param name="culture">当前请求的转换与校验区域性。</param>
    /// <param name="errors">错误集合。</param>
    /// <returns>行通过校验时返回 <see langword="true"/>。</returns>
    private bool ValidateRawValues(IRow row, IReadOnlyDictionary<int, ExcelColumnPlan> columns,
        IDictionary<string, HashSet<string>> duplicateValues, string sheetName, int rowIndex, ValidateMode validateMode,
        CultureInfo culture, ExcelWhitespacePolicy bodyWhitespace, ExcelImportErrorCollector errors)
    {
        var valid = true;
        foreach (var column in columns)
        {
            var cell = row.GetCell(column.Key);
            var cellValue = NormalizeCellValue(ApplyLegacyTextConverters(cell, ReadCellValue(cell)), bodyWhitespace);
            var value = cellValue.Text;
            foreach (var binding in column.Value.AttributeValidations.Where(binding => binding.IsRaw))
            {
                var context = new ExcelValidationContext(value, sheetName, rowIndex + 1, column.Key + 1,
                    column.Value.Property.Name, duplicateValues, null, column.Value.ValueType, cellValue,
                    culture, column.Value.Property.Property);
                bool isValid;
                try
                {
                    isValid = binding.Rule.Validate(binding.Attribute, context);
                }
                catch (Exception exception)
                {
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, exception.Message, sheetName,
                        rowIndex + 1, column.Key + 1, column.Value.Property.Name, column.Value.Property.Name,
                        column.Value.HeaderName, cellValue.Value ?? cellValue.Text));
                    valid = false;
                    if (validateMode == ValidateMode.StopOnFirstFailure)
                        return false;
                    continue;
                }
                if (isValid)
                    continue;
                errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, binding.Attribute.ErrorMsg, sheetName,
                    rowIndex + 1, column.Key + 1, column.Value.Property.Name, column.Value.Property.Name,
                    column.Value.HeaderName, cellValue.Value ?? cellValue.Text));
                valid = false;
                if (validateMode == ValidateMode.StopOnFirstFailure)
                    return false;
            }
        }
        return valid;
    }

    /// <summary>
    /// 校验转换后的属性值及配置命名规则。
    /// </summary>
    private bool ValidateColumnValue(string value, ExcelCellValue cellValue, object convertedValue, ExcelColumnPlan column,
        IDictionary<string, HashSet<string>> duplicateValues, string sheetName, int rowIndex, ValidateMode validateMode,
        CultureInfo culture, ExcelImportErrorCollector errors)
    {
        var valid = true;
        var property = column.Property;
        var context = new ExcelValidationContext(value, sheetName, rowIndex + 1, column.ColumnIndex + 1,
            column.IsDynamic ? column.Key : property.Name,
            duplicateValues, convertedValue, column.ValueType, cellValue, culture, property.Property);
        foreach (var binding in column.AttributeValidations.Where(binding => !binding.IsRaw))
        {
            bool isValid;
            try
            {
                isValid = binding.Rule.Validate(binding.Attribute, context);
            }
            catch (Exception exception)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, exception.Message, sheetName,
                    rowIndex + 1, column.ColumnIndex + 1, property.Name, property.Name, column.HeaderName,
                    cellValue.Value ?? cellValue.Text));
                valid = false;
                if (validateMode == ValidateMode.StopOnFirstFailure)
                    return false;
                continue;
            }
            if (isValid)
                continue;
            errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, binding.Attribute.ErrorMsg, sheetName,
                rowIndex + 1, column.ColumnIndex + 1, property.Name, property.Name, column.HeaderName,
                cellValue.Value ?? cellValue.Text));
            valid = false;
            if (validateMode == ValidateMode.StopOnFirstFailure)
                return false;
        }
        foreach (var rule in column.NamedValidationRules)
        {
            bool isValid;
            try
            {
                isValid = rule.Validate(context);
            }
            catch (Exception exception)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, exception.Message, sheetName,
                    rowIndex + 1, column.ColumnIndex + 1, property.Name, property.Name, column.HeaderName,
                    cellValue.Value ?? cellValue.Text));
                valid = false;
                if (validateMode == ValidateMode.StopOnFirstFailure)
                    return false;
                continue;
            }
            if (isValid)
                continue;
            errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, rule.ErrorMessage, sheetName,
                rowIndex + 1, column.ColumnIndex + 1, property.Name, property.Name, column.HeaderName,
                cellValue.Value ?? cellValue.Text));
            valid = false;
            if (validateMode == ValidateMode.StopOnFirstFailure)
                return false;
        }
        return valid;
    }

    private static bool IsRawValidationAttribute(FilterAttributeBase attribute) =>
        attribute is RequiredAttribute || attribute is RegexAttribute;

    /// <summary>
    /// 捕获当前行可能新增的重复校验值，行失败时只回滚本行新增项。
    /// </summary>
    private static IReadOnlyList<DuplicateChange> CaptureDuplicateChanges(IRow row,
        IReadOnlyDictionary<int, ExcelColumnPlan> columns, IReadOnlyDictionary<string, HashSet<string>> duplicateValues,
        ExcelWhitespacePolicy bodyWhitespace)
    {
        var changes = new List<DuplicateChange>();
        foreach (var column in columns.Values)
        {
            if (!column.IsUnique)
                continue;
            var cellValue = NormalizeCellValue(ReadCellValue(row.GetCell(column.ColumnIndex)), bodyWhitespace);
            if (string.IsNullOrWhiteSpace(cellValue.Text))
                continue;
            var key = column.Key;
            var wasPresent = duplicateValues.TryGetValue(key, out var values)
                             && values.Contains(cellValue.Text);
            changes.Add(new DuplicateChange(key, cellValue.Text, wasPresent));
        }
        return changes;
    }

    /// <summary>
    /// 回滚当前失败行新增的重复校验值。
    /// </summary>
    private static void RollbackDuplicateChanges(IDictionary<string, HashSet<string>> duplicateValues,
        IReadOnlyList<DuplicateChange> changes)
    {
        foreach (var change in changes)
        {
            if (change.WasPresent || !duplicateValues.TryGetValue(change.PropertyName, out var values))
                continue;
            values.Remove(change.Value);
            if (values.Count == 0)
                duplicateValues.Remove(change.PropertyName);
        }
    }

    /// <summary>
    /// 解析指定特性对应的校验规则。
    /// </summary>
    /// <param name="attribute">筛选特性。</param>
    private IExcelValidationRule ResolveValidationRule(FilterAttributeBase attribute)
    {
        var binding = attribute.GetType().GetCustomAttribute<BindFilterAttribute>();
        if (binding != null)
        {
            var boundRule = _validationRules.FirstOrDefault(rule => binding.RuleType.IsInstanceOfType(rule));
            if (boundRule == null)
                throw new InvalidOperationException($"未注册校验规则: {binding.RuleType.FullName}");
            if (!boundRule.CanValidate(attribute))
                throw new InvalidOperationException($"校验规则 {binding.RuleType.FullName} 不支持特性: {attribute.GetType().FullName}");
            return boundRule;
        }

        var rule = _validationRules.FirstOrDefault(candidate => candidate.CanValidate(attribute));
        return rule ?? throw new InvalidOperationException($"未找到特性对应的校验规则: {attribute.GetType().FullName}");
    }

    /// <summary>
    /// 在读取数据前验证静态特性校验规则均可解析。
    /// </summary>
    /// <summary>
    /// 解析配置指定的命名校验规则。
    /// </summary>
    /// <param name="ruleName">规则名称。</param>
    private INamedExcelValidationRule ResolveNamedValidationRule(string ruleName)
    {
        var rules = _namedValidationRules.Where(rule => string.Equals(rule.Name, ruleName,
            StringComparison.OrdinalIgnoreCase)).ToList();
        if (rules.Count != 1)
            throw new InvalidOperationException($"未找到唯一命名校验规则: {ruleName}");
        return rules[0];
    }

    private IReadOnlyList<IExcelValueConverter> BindValueConverters(string converterName, Type propertyType)
    {
        if (string.IsNullOrWhiteSpace(converterName))
            return _valueConverters.Where(converter => converter.CanConvert(propertyType)).ToArray();
        var converters = _valueConverters.OfType<INamedExcelValueConverter>().Where(converter =>
            string.Equals(converter.Name, converterName, StringComparison.OrdinalIgnoreCase)).ToList();
        if (converters.Count != 1)
            throw new InvalidOperationException($"未找到唯一命名值转换器: {converterName}");
        if (!converters[0].CanConvert(propertyType))
            throw new InvalidOperationException($"值转换器 {converterName} 不支持属性类型: {propertyType.FullName}");
        return converters;
    }

    /// <summary>
    /// 判断行是否没有任何非空单元格。
    /// </summary>
    /// <param name="row">NPOI 行。</param>
    private static bool IsEmpty(IRow row, ExcelWhitespacePolicy bodyWhitespace,
        IReadOnlyDictionary<(int Row, int Column), IReadOnlyList<PictureInfo>> imageIndex, int rowIndex) => row == null
        || (imageIndex == null || !imageIndex.Keys.Any(key => key.Row == rowIndex))
        && row.Cells.All(cell => string.IsNullOrWhiteSpace(NormalizeText(GetRawStringValue(cell), bodyWhitespace)));

    /// <summary>
    /// 单行重复校验状态变更。
    /// </summary>
    private sealed class DuplicateChange
    {
        public DuplicateChange(string propertyName, string value, bool wasPresent)
        {
            PropertyName = propertyName;
            Value = value;
            WasPresent = wasPresent;
        }

        public string PropertyName { get; }

        public string Value { get; }

        public bool WasPresent { get; }
    }

    private sealed class SheetStructureException : InvalidOperationException
    {
        public SheetStructureException(string message) : base(message)
        {
        }
    }
}

internal sealed class ExcelImportRuntime
{
    private int _rowCount;
    private bool _rowLimitReported;

    internal ExcelImportRuntime(ExcelResourceLimits limits)
    {
        MaxRows = limits?.MaxRows;
        ImageResources = new ExcelImageResourceTracker(limits);
    }

    internal int? MaxRows { get; }

    internal bool RowLimitReached => MaxRows.HasValue && _rowCount >= MaxRows.Value;

    internal ExcelImageResourceTracker ImageResources { get; }

    internal bool TryConsumeRow()
    {
        if (MaxRows.HasValue && _rowCount >= MaxRows.Value)
            return false;
        _rowCount++;
        return true;
    }

    internal bool TryMarkRowLimitReported()
    {
        if (_rowLimitReported)
            return false;
        _rowLimitReported = true;
        return true;
    }
}

internal sealed class ExcelImageResourceTracker
{
    private readonly int? _maxPictures;
    private readonly long? _maxPictureBytes;
    private readonly long? _maxTotalPictureBytes;
    private int _count;
    private long _totalBytes;

    internal ExcelImageResourceTracker(ExcelResourceLimits limits)
    {
        _maxPictures = limits?.MaxPictures;
        _maxPictureBytes = limits?.MaxPictureBytes;
        _maxTotalPictureBytes = limits?.MaxTotalPictureBytes;
    }

    internal void Consume(long bytes)
    {
        if (_maxPictureBytes.HasValue && bytes > _maxPictureBytes.Value)
            throw new ImageResourceLimitException($"单张图片超过最大字节数: {_maxPictureBytes.Value}");
        if (_maxPictures.HasValue && _count >= _maxPictures.Value)
            throw new ImageResourceLimitException($"图片数量超过限制: {_maxPictures.Value}");
        if (_maxTotalPictureBytes.HasValue && bytes > _maxTotalPictureBytes.Value - _totalBytes)
            throw new ImageResourceLimitException($"图片总字节数超过限制: {_maxTotalPictureBytes.Value}");
        _count++;
        _totalBytes += bytes;
    }
}

internal sealed class ImageResourceLimitException : InvalidOperationException
{
    internal ImageResourceLimitException(string message) : base(message)
    {
    }
}

internal sealed class SourceLocation
{
    internal SourceLocation(string sheetName, int rowIndex)
    {
        SheetName = sheetName;
        RowIndex = rowIndex;
    }

    internal string SheetName { get; }
    internal int RowIndex { get; }
}

internal sealed class ReferenceObjectComparer : IEqualityComparer<object>
{
    internal static readonly ReferenceObjectComparer Instance = new();

    public new bool Equals(object x, object y) => ReferenceEquals(x, y);

    public int GetHashCode(object obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
}

#pragma warning restore CS0618
