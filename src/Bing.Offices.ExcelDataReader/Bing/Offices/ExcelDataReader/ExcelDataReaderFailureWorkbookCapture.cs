using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.ExcelDataReader;

/// <summary>
/// 保存前向读取过程中发现的失败行快照。
/// </summary>
internal sealed class ExcelDataReaderFailureWorkbookCapture
{
    /// <summary>
    /// 失败工作簿候选行数量上限。
    /// </summary>
    private readonly int? _maxRows;

    /// <summary>
    /// 是否启用失败行捕获。
    /// </summary>
    private readonly bool _enabled;

    /// <summary>
    /// 是否保存失败行的原始单元格快照。
    /// </summary>
    private readonly bool _captureRows;

    /// <summary>
    /// 失败工作簿候选单元格数量上限。
    /// </summary>
    private readonly long? _maxCells;

    /// <summary>
    /// 失败工作簿目标对象数量上限。
    /// </summary>
    private readonly long? _maxEstimatedTargetObjects;

    /// <summary>
    /// 生成资源诊断时使用的 Provider 名称。
    /// </summary>
    private readonly string _providerName;

    /// <summary>
    /// 是否额外保留全局错误列表。
    /// </summary>
    private readonly bool _captureErrors;

    /// <summary>
    /// 可安全重映射的 XLSX 工作簿展示元数据。
    /// </summary>
    private ExcelXlsxFailureWorkbookMetadata _metadata;

    /// <summary>
    /// 已捕获的失败行数量。
    /// </summary>
    private int _rowCount;

    /// <summary>
    /// 已捕获的失败行单元格数量。
    /// </summary>
    private long _cellCount;

    /// <summary>
    /// 已为失败快照预留的目标对象数量。
    /// </summary>
    private long _estimatedTargetObjects;

    /// <summary>
    /// 按源工作表顺序保存失败行。
    /// </summary>
    private readonly List<ExcelDataReaderFailureSheet> _sheets = new List<ExcelDataReaderFailureSheet>();

    /// <summary>
    /// 保存批量导入阶段收集的结构化错误。
    /// </summary>
    private readonly List<ExcelImportError> _errors = new List<ExcelImportError>();

    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderFailureWorkbookCapture" /> 类型的实例。
    /// </summary>
    /// <param name="options">失败工作簿资源限制；未设置时不限制候选行。</param>
    /// <param name="providerName">资源诊断使用的 Provider 名称；为空时使用 ExcelDataReader。</param>
    /// <param name="captureErrors">是否额外保留全局错误列表。</param>
    internal ExcelDataReaderFailureWorkbookCapture(ExcelImportFailureOptions options,
        string providerName = null, bool captureErrors = true)
    {
        _enabled = options?.Mode != null && options.Mode != ExcelImportFailureWorkbookMode.None;
        _captureRows = options?.Mode == ExcelImportFailureWorkbookMode.ErrorRowsOnly;
        _maxRows = options?.MaxCandidateErrorRows;
        _maxCells = options?.MaxCopiedCells;
        _maxEstimatedTargetObjects = options?.MaxEstimatedTargetObjects;
        _providerName = string.IsNullOrWhiteSpace(providerName) ? "ExcelDataReader" : providerName;
        _captureErrors = captureErrors;
    }

    /// <summary>
    /// 获取按源工作表顺序排列的失败行快照。
    /// </summary>
    internal IReadOnlyList<ExcelDataReaderFailureSheet> Sheets => _sheets;

    /// <summary>
    /// 获取批量导入阶段收集的结构化错误快照。
    /// </summary>
    internal IReadOnlyList<ExcelImportError> Errors => _errors;

    /// <summary>
    /// 获取资源诊断使用的 Provider 名称。
    /// </summary>
    internal string ProviderName => _providerName;

    /// <summary>
    /// 获取可安全重映射的 XLSX 工作簿展示元数据。
    /// </summary>
    internal ExcelXlsxFailureWorkbookMetadata Metadata
        => _metadata ?? ExcelXlsxFailureWorkbookMetadata.Empty;

    /// <summary>
    /// 保存可安全重映射的 XLSX 工作簿展示元数据。
    /// </summary>
    /// <param name="metadata">工作簿展示元数据。</param>
    internal void CaptureWorkbookMetadata(ExcelXlsxFailureWorkbookMetadata metadata)
    {
        if (!_captureRows || metadata == null || _metadata != null)
            return;
        ReserveObjects(CountMetadataObjects(metadata));
        _metadata = metadata;
    }

    /// <summary>
    /// 注册一个源工作表及其表头快照。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <param name="header">表头单元格值。</param>
    /// <param name="headerRowIndex">表头的 zero-based 行索引。</param>
    /// <param name="numberFormats">表头单元格的数字格式。</param>
    /// <param name="boldColumns">表头单元格的加粗列。</param>
    /// <param name="fontStyles">表头单元格的字体状态。</param>
    internal void CaptureSheet(string name, IReadOnlyList<object> header, int headerRowIndex,
        IReadOnlyDictionary<int, string> numberFormats = null,
        IReadOnlyDictionary<int, bool> boldColumns = null,
        IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> fontStyles = null)
    {
        if (!_captureRows)
            return;
        if (string.IsNullOrWhiteSpace(name))
            return;
        var existing = FindSheet(name);
        if (existing != null)
        {
            if (existing.Header.Count == 0 && header != null && header.Count > 0)
                ReserveObjects(header.Count);
            existing.SetHeader(header, headerRowIndex, numberFormats, boldColumns, fontStyles);
            return;
        }
        ReserveObjects(2L + (header?.Count ?? 0));
        _sheets.Add(new ExcelDataReaderFailureSheet(name, header, headerRowIndex,
            numberFormats, boldColumns, fontStyles));
    }

    /// <summary>
    /// 保存不属于失败行、但被安全命名区域引用的源数据行。
    /// </summary>
    /// <param name="sheetName">源工作表名称。</param>
    /// <param name="rowIndex">从零开始的源行号。</param>
    /// <param name="values">源行单元格值；数组索引对应源列号。</param>
    /// <param name="numberFormats">源单元格数字格式。</param>
    /// <param name="boldColumns">源单元格加粗列。</param>
    /// <param name="fontStyles">源单元格字体状态。</param>
    internal void CaptureContextRow(string sheetName, int rowIndex,
        IReadOnlyList<object> values,
        IReadOnlyDictionary<int, string> numberFormats = null,
        IReadOnlyDictionary<int, bool> boldColumns = null,
        IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> fontStyles = null)
    {
        if (!_enabled || !_captureRows || string.IsNullOrWhiteSpace(sheetName)
            || rowIndex < 0)
            return;
        var sheet = GetOrCreateSheet(sheetName);
        var sourceRowIndex = rowIndex + 1;
        if (sheet.ContextRows.Any(row => row.RowIndex == sourceRowIndex))
            return;
        var count = values?.Count ?? 0;
        if (_maxCells.HasValue && count > _maxCells.Value - _cellCount)
            throw Limit($"失败工作簿复制单元格数量超过资源限制: {_maxCells.Value}");
        ReserveObjects(1L + count);
        _cellCount += count;
        sheet.ContextRows.Add(new ExcelDataReaderFailureRow(sourceRowIndex, values,
            Array.Empty<ExcelImportError>(), numberFormats, boldColumns, fontStyles));
    }

    /// <summary>
    /// 保存源工作表的可见性状态。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <param name="visibility">空值、hidden 或 veryHidden。</param>
    internal void CaptureSheetVisibility(string name, string visibility)
    {
        if (!_captureRows || string.IsNullOrWhiteSpace(name))
            return;
        var sheet = GetOrCreateSheet(name);
        sheet.SetVisibility(visibility);
    }

    /// <summary>
    /// 保存源工作表的显式列宽。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <param name="columnMetadata">零基列号到列宽和隐藏状态的映射。</param>
    internal void CaptureSheetColumnMetadata(string name,
        IReadOnlyDictionary<int, (double Width, bool HasWidth, bool Hidden)> columnMetadata)
    {
        if (!_captureRows || string.IsNullOrWhiteSpace(name))
            return;
        var sheet = GetOrCreateSheet(name);
        if (sheet.ColumnMetadataCaptured)
            return;
        sheet.SetColumnMetadata(columnMetadata);
    }

    /// <summary>
    /// 保存源工作表的显式行高和隐藏状态。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <param name="rowMetadata">零基行号到行高和隐藏状态的映射。</param>
    internal void CaptureSheetRowMetadata(string name,
        IReadOnlyDictionary<int, (short Height, bool HasHeight, bool Hidden)> rowMetadata)
    {
        if (!_captureRows || string.IsNullOrWhiteSpace(name))
            return;
        var sheet = GetOrCreateSheet(name);
        if (sheet.RowMetadataCaptured)
            return;
        sheet.SetRowMetadata(rowMetadata);
    }

    /// <summary>
    /// 保存源工作表的合并区域快照。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <param name="mergeRegions">零基行列坐标的合并区域。</param>
    internal void CaptureSheetMergeRegions(string name,
        IReadOnlyList<(int FirstRow, int LastRow, int FirstColumn, int LastColumn)> mergeRegions)
    {
        if (!_captureRows || string.IsNullOrWhiteSpace(name))
            return;
        var sheet = GetOrCreateSheet(name);
        if (sheet.MergeRegionsCaptured)
            return;
        sheet.SetMergeRegions(mergeRegions);
    }

    /// <summary>
    /// 保存源工作表的冻结窗格坐标。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <param name="columnSplit">冻结前的列数。</param>
    /// <param name="rowSplit">冻结前的行数。</param>
    /// <param name="leftColumn">冻结后左侧可见列。</param>
    /// <param name="topRow">冻结后顶部可见行。</param>
    internal void CaptureSheetFreezePane(string name, int columnSplit, int rowSplit,
        int leftColumn, int topRow)
    {
        if (!_captureRows || string.IsNullOrWhiteSpace(name))
            return;
        var sheet = GetOrCreateSheet(name);
        sheet.SetFreezePane(columnSplit, rowSplit, leftColumn, topRow);
    }

    /// <summary>
    /// 标记源工作表包含原生数据校验规则。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    internal void CaptureSheetNativeDataValidation(string name)
    {
        if (!_captureRows || string.IsNullOrWhiteSpace(name))
            return;
        var sheet = GetOrCreateSheet(name);
        sheet.HasNativeDataValidation = true;
    }

    /// <summary>
    /// 保存源工作表能够参与低保留重建的原生数据校验描述。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <param name="validations">原生数据校验描述。</param>
    internal void CaptureSheetNativeDataValidations(string name,
        IReadOnlyList<ExcelDataReaderFailureNativeDataValidation> validations)
    {
        if (!_captureRows || string.IsNullOrWhiteSpace(name))
            return;
        var sheet = GetOrCreateSheet(name);
        sheet.HasNativeDataValidation = true;
        sheet.SetNativeDataValidations(validations);
    }

    /// <summary>
    /// 保存源工作表能够参与低保留重建的条件格式描述。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <param name="unsupported">是否包含当前不支持的条件格式。</param>
    /// <param name="conditionalFormattings">条件格式描述。</param>
    internal void CaptureSheetConditionalFormattings(string name, bool unsupported,
        IReadOnlyList<ExcelDataReaderFailureConditionalFormatting> conditionalFormattings)
    {
        if (!_captureRows || string.IsNullOrWhiteSpace(name))
            return;
        var sheet = GetOrCreateSheet(name);
        sheet.HasConditionalFormatting = true;
        sheet.SetConditionalFormattings(unsupported ? null : conditionalFormattings);
    }

    /// <summary>
    /// 保存一行失败数据及其结构化错误。
    /// </summary>
    /// <param name="sheetName">源工作表名称。</param>
    /// <param name="rowIndex">从一开始的源行号。</param>
    /// <param name="values">当前行的原始单元格值。</param>
    /// <param name="errors">当前行的结构化错误。</param>
    /// <param name="numberFormats">失败行单元格的数字格式。</param>
    /// <param name="boldColumns">失败行单元格的加粗列。</param>
    /// <param name="fontStyles">失败行单元格的字体状态。</param>
    internal void CaptureErrorRow(string sheetName, int rowIndex, IReadOnlyList<object> values,
        IReadOnlyList<ExcelImportError> errors,
        IReadOnlyDictionary<int, string> numberFormats = null,
        IReadOnlyDictionary<int, bool> boldColumns = null,
        IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> fontStyles = null)
    {
        if (!_enabled)
            return;
        if (!_captureRows)
        {
            CaptureErrors(errors);
            return;
        }
        if (errors == null || errors.Count == 0)
            return;
        var sheet = FindSheet(sheetName);
        if (sheet == null)
        {
            CaptureSheet(sheetName, Array.Empty<object>(), Math.Max(0, rowIndex - 1),
                numberFormats, boldColumns, fontStyles);
            sheet = FindSheet(sheetName);
        }
        if (_maxRows.HasValue && _rowCount >= _maxRows.Value)
            throw Limit($"失败工作簿错误行数量超过资源限制: {_maxRows.Value}");
        var count = values?.Count ?? 0;
        if (_maxCells.HasValue && count > _maxCells.Value - _cellCount)
            throw Limit($"失败工作簿复制单元格数量超过资源限制: {_maxCells.Value}");
        ReserveObjects(1L + count);
        _rowCount++;
        _cellCount += count;
        if (_captureErrors)
            CaptureErrors(errors);
        // 低保留路径由写出阶段按源行从最终错误集合回查，避免在每个失败行重复保存错误引用。
        var rowErrors = _captureErrors ? errors : Array.Empty<ExcelImportError>();
        sheet.Rows.Add(new ExcelDataReaderFailureRow(rowIndex, values, rowErrors,
            numberFormats, boldColumns, fontStyles));
    }

    /// <summary>
    /// 保存没有对应原始数据行的结构化错误。
    /// </summary>
    /// <param name="errors">待保存的错误集合。</param>
    internal void CaptureErrors(IEnumerable<ExcelImportError> errors)
    {
        if (!_enabled || !_captureErrors || errors == null)
            return;
        foreach (var error in errors)
        {
            if (error == null)
                continue;
            ReserveObjects(1);
            _errors.Add(error);
        }
    }

    /// <summary>
    /// 查找已捕获的源工作表。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <returns>匹配的工作表；不存在时返回 <see langword="null" />。</returns>
    private ExcelDataReaderFailureSheet FindSheet(string name)
        => _sheets.FirstOrDefault(sheet =>
            string.Equals(sheet.Name, name, StringComparison.Ordinal));

    /// <summary>
    /// 获取或创建一个空表头的失败工作表快照。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <returns>失败工作表快照。</returns>
    private ExcelDataReaderFailureSheet GetOrCreateSheet(string name)
    {
        var existing = FindSheet(name);
        if (existing != null)
            return existing;
        ReserveObjects(2);
        var sheet = new ExcelDataReaderFailureSheet(name, Array.Empty<object>(), 0);
        _sheets.Add(sheet);
        return sheet;
    }

    /// <summary>
    /// 估算一次元数据快照需要的目标对象数量。
    /// </summary>
    /// <param name="metadata">工作簿元数据。</param>
    /// <returns>目标对象估算数量。</returns>
    private static long CountMetadataObjects(ExcelXlsxFailureWorkbookMetadata metadata)
    {
        long count = metadata.DefinedNames.Count;
        foreach (var sheet in metadata.Sheets)
        {
            count = checked(count + sheet.Validations.Count);
            foreach (var validation in sheet.Validations)
                count = checked(count + validation.Ranges.Count);
            foreach (var formatting in sheet.ConditionalFormattings)
                count = checked(count + 1L + formatting.Ranges.Count);
            count = checked(count + sheet.ColumnRanges.Count
                + sheet.UnsupportedColumnMetadataCount);
        }
        return count;
    }

    /// <summary>
    /// 为快照对象预留资源预算。
    /// </summary>
    /// <param name="count">待预留的对象数量。</param>
    private void ReserveObjects(long count)
    {
        if (count <= 0 || !_maxEstimatedTargetObjects.HasValue)
            return;
        long total;
        try
        {
            total = checked(_estimatedTargetObjects + count);
        }
        catch (OverflowException)
        {
            throw Limit($"失败工作簿目标对象估算数量超过限制: {_maxEstimatedTargetObjects.Value}");
        }
        if (total > _maxEstimatedTargetObjects.Value)
            throw Limit($"失败工作簿目标对象估算数量超过限制: {_maxEstimatedTargetObjects.Value}");
        _estimatedTargetObjects = total;
    }

    /// <summary>
    /// 创建失败工作簿资源限制异常。
    /// </summary>
    /// <param name="message">限制说明。</param>
    /// <returns>资源限制异常。</returns>
    private BingOfficesResourceLimitException Limit(string message)
        => new BingOfficesResourceLimitException(message, provider: _providerName,
            operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);
}

/// <summary>
/// 一个源工作表的失败行快照。
/// </summary>
internal sealed class ExcelDataReaderFailureSheet
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderFailureSheet" /> 类型的实例。
    /// </summary>
    /// <param name="name">源工作表名称。</param>
    /// <param name="header">表头单元格值。</param>
    /// <param name="headerRowIndex">表头的 zero-based 行索引。</param>
    /// <param name="numberFormats">表头单元格的数字格式。</param>
    /// <param name="boldColumns">表头单元格的加粗列。</param>
    /// <param name="fontStyles">表头单元格的字体状态。</param>
    internal ExcelDataReaderFailureSheet(string name, IReadOnlyList<object> header, int headerRowIndex,
        IReadOnlyDictionary<int, string> numberFormats = null,
        IReadOnlyDictionary<int, bool> boldColumns = null,
        IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> fontStyles = null)
    {
        Name = name;
        Header = (header ?? Array.Empty<object>()).ToArray();
        HeaderRowIndex = headerRowIndex;
        HeaderFormats = ExcelDataReaderFailureMetadata.CopyNumberFormats(numberFormats);
        HeaderBoldColumns = ExcelDataReaderFailureMetadata.CopyBoldColumns(boldColumns);
        HeaderFontStyles = ExcelDataReaderFailureMetadata.CopyFontStyles(fontStyles);
        ColumnMetadata = new Dictionary<int, ExcelDataReaderFailureColumnMetadata>();
        RowMetadata = new Dictionary<int, ExcelDataReaderFailureRowMetadata>();
        MergeRegions = Array.Empty<ExcelDataReaderFailureMergeRegion>();
    }

    /// <summary>
    /// 获取源工作表名称。
    /// </summary>
    internal string Name { get; }

    /// <summary>
    /// 获取表头单元格值。
    /// </summary>
    internal IReadOnlyList<object> Header { get; private set; }

    /// <summary>
    /// 获取表头的 zero-based 行索引。
    /// </summary>
    internal int HeaderRowIndex { get; private set; }

    /// <summary>
    /// 获取表头单元格的数字格式。
    /// </summary>
    internal IReadOnlyDictionary<int, string> HeaderFormats { get; private set; }

    /// <summary>
    /// 获取表头单元格的加粗列。
    /// </summary>
    internal IReadOnlyDictionary<int, bool> HeaderBoldColumns { get; private set; }

    /// <summary>
    /// 获取表头单元格的字体状态。
    /// </summary>
    internal IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> HeaderFontStyles
    {
        get;
        private set;
    }

    /// <summary>
    /// 获取源工作表的显式列元数据。
    /// </summary>
    internal IReadOnlyDictionary<int, ExcelDataReaderFailureColumnMetadata> ColumnMetadata
    {
        get;
        private set;
    }

    /// <summary>
    /// 获取是否已读取过源工作表列宽元数据。
    /// </summary>
    internal bool ColumnMetadataCaptured { get; private set; }

    /// <summary>
    /// 获取源工作表的显式行元数据。
    /// </summary>
    internal IReadOnlyDictionary<int, ExcelDataReaderFailureRowMetadata> RowMetadata
    {
        get;
        private set;
    }

    /// <summary>
    /// 获取是否已读取过源工作表行元数据。
    /// </summary>
    internal bool RowMetadataCaptured { get; private set; }

    /// <summary>
    /// 获取源工作表的合并区域。
    /// </summary>
    internal IReadOnlyList<ExcelDataReaderFailureMergeRegion> MergeRegions { get; private set; }

    /// <summary>
    /// 获取是否已读取过源工作表合并区域。
    /// </summary>
    internal bool MergeRegionsCaptured { get; private set; }

    /// <summary>
    /// 获取源工作表的冻结窗格坐标。
    /// </summary>
    internal ExcelDataReaderFailureFreezePane FreezePane { get; private set; }

    /// <summary>
    /// 获取是否包含源工作簿条件格式规则。
    /// </summary>
    internal bool HasConditionalFormatting { get; set; }

    /// <summary>
    /// 获取能够参与低保留重建的条件格式描述。
    /// </summary>
    internal IReadOnlyList<ExcelDataReaderFailureConditionalFormatting> ConditionalFormattings
        { get; private set; } = Array.Empty<ExcelDataReaderFailureConditionalFormatting>();

    /// <summary>
    /// 获取是否包含源工作簿原生数据校验规则。
    /// </summary>
    internal bool HasNativeDataValidation { get; set; }

    /// <summary>
    /// 获取能够参与低保留重建的原生数据校验描述。
    /// </summary>
    internal IReadOnlyList<ExcelDataReaderFailureNativeDataValidation> NativeDataValidations
        { get; private set; } = Array.Empty<ExcelDataReaderFailureNativeDataValidation>();

    /// <summary>
    /// 获取源工作表的可见性状态。
    /// </summary>
    internal string Visibility { get; private set; }

    /// <summary>
    /// 获取失败行快照。
    /// </summary>
    internal List<ExcelDataReaderFailureRow> Rows { get; } = new List<ExcelDataReaderFailureRow>();

    /// <summary>
    /// 获取命名区域引用的上下文行快照。
    /// </summary>
    internal List<ExcelDataReaderFailureRow> ContextRows { get; }
        = new List<ExcelDataReaderFailureRow>();

    /// <summary>
    /// 更新表头快照；仅允许补充预先创建的空快照。
    /// </summary>
    /// <param name="header">表头单元格值。</param>
    /// <param name="headerRowIndex">表头的 zero-based 行索引。</param>
    /// <param name="numberFormats">表头单元格的数字格式。</param>
    /// <param name="boldColumns">表头单元格的加粗列。</param>
    /// <param name="fontStyles">表头单元格的字体状态。</param>
    internal void SetHeader(IReadOnlyList<object> header, int headerRowIndex,
        IReadOnlyDictionary<int, string> numberFormats = null,
        IReadOnlyDictionary<int, bool> boldColumns = null,
        IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> fontStyles = null)
    {
        if (Header.Count != 0 || header == null || header.Count == 0)
            return;
        Header = header.ToArray();
        HeaderRowIndex = headerRowIndex;
        HeaderFormats = ExcelDataReaderFailureMetadata.CopyNumberFormats(numberFormats);
        HeaderBoldColumns = ExcelDataReaderFailureMetadata.CopyBoldColumns(boldColumns);
        HeaderFontStyles = ExcelDataReaderFailureMetadata.CopyFontStyles(fontStyles);
    }

    /// <summary>
    /// 设置源工作表的显式列元数据快照。
    /// </summary>
    /// <param name="columnMetadata">零基列号到列元数据的映射。</param>
    internal void SetColumnMetadata(
        IReadOnlyDictionary<int, (double Width, bool HasWidth, bool Hidden)> columnMetadata)
    {
        var metadata = new Dictionary<int, ExcelDataReaderFailureColumnMetadata>();
        if (columnMetadata != null)
        {
            foreach (var pair in columnMetadata)
                metadata[pair.Key] = new ExcelDataReaderFailureColumnMetadata(pair.Value.Width,
                    pair.Value.HasWidth, pair.Value.Hidden);
        }
        ColumnMetadata = metadata;
        ColumnMetadataCaptured = true;
    }

    /// <summary>
    /// 设置源工作表的显式行元数据快照。
    /// </summary>
    /// <param name="rowMetadata">零基行号到行元数据的映射。</param>
    internal void SetRowMetadata(
        IReadOnlyDictionary<int, (short Height, bool HasHeight, bool Hidden)> rowMetadata)
    {
        var metadata = new Dictionary<int, ExcelDataReaderFailureRowMetadata>();
        if (rowMetadata != null)
        {
            foreach (var pair in rowMetadata)
                metadata[pair.Key] = new ExcelDataReaderFailureRowMetadata(pair.Value.Height,
                    pair.Value.HasHeight, pair.Value.Hidden);
        }
        RowMetadata = metadata;
        RowMetadataCaptured = true;
    }

    /// <summary>
    /// 设置源工作表的合并区域快照。
    /// </summary>
    /// <param name="mergeRegions">零基行列坐标的合并区域。</param>
    internal void SetMergeRegions(
        IReadOnlyList<(int FirstRow, int LastRow, int FirstColumn, int LastColumn)> mergeRegions)
    {
        var regions = new List<ExcelDataReaderFailureMergeRegion>();
        var seen = new HashSet<(int FirstRow, int LastRow, int FirstColumn, int LastColumn)>();
        if (mergeRegions != null)
        {
            foreach (var region in mergeRegions)
            {
                if (region.FirstRow < 0 || region.LastRow < region.FirstRow
                    || region.FirstColumn < 0 || region.LastColumn < region.FirstColumn
                    || !seen.Add(region))
                    continue;
                regions.Add(new ExcelDataReaderFailureMergeRegion(region.FirstRow,
                    region.LastRow, region.FirstColumn, region.LastColumn));
            }
        }
        MergeRegions = regions;
        MergeRegionsCaptured = true;
    }

    /// <summary>
    /// 设置源工作表的冻结窗格坐标。
    /// </summary>
    /// <param name="columnSplit">冻结前的列数。</param>
    /// <param name="rowSplit">冻结前的行数。</param>
    /// <param name="leftColumn">冻结后左侧可见列。</param>
    /// <param name="topRow">冻结后顶部可见行。</param>
    internal void SetFreezePane(int columnSplit, int rowSplit, int leftColumn, int topRow)
    {
        FreezePane = new ExcelDataReaderFailureFreezePane(columnSplit, rowSplit,
            leftColumn, topRow);
    }

    /// <summary>
    /// 保存条件格式描述快照。
    /// </summary>
    /// <param name="conditionalFormattings">条件格式描述。</param>
    internal void SetConditionalFormattings(
        IReadOnlyList<ExcelDataReaderFailureConditionalFormatting> conditionalFormattings)
    {
        ConditionalFormattings = conditionalFormattings == null
            ? Array.Empty<ExcelDataReaderFailureConditionalFormatting>()
            : conditionalFormattings.Where(item => item != null).ToArray();
    }

    /// <summary>
    /// 保存原生数据校验描述快照。
    /// </summary>
    /// <param name="validations">原生数据校验描述。</param>
    internal void SetNativeDataValidations(
        IReadOnlyList<ExcelDataReaderFailureNativeDataValidation> validations)
    {
        NativeDataValidations = validations == null
            ? Array.Empty<ExcelDataReaderFailureNativeDataValidation>()
            : validations.Where(item => item != null).ToArray();
    }

    /// <summary>
    /// 设置源工作表的可见性状态。
    /// </summary>
    /// <param name="visibility">空值、hidden 或 veryHidden。</param>
    internal void SetVisibility(string visibility)
    {
        Visibility = string.Equals(visibility, "hidden", StringComparison.OrdinalIgnoreCase)
            ? "hidden"
            : string.Equals(visibility, "veryHidden", StringComparison.OrdinalIgnoreCase)
                ? "veryHidden" : null;
    }
}

/// <summary>
/// 保存失败工作簿的冻结窗格坐标。
/// </summary>
internal sealed class ExcelDataReaderFailureFreezePane
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderFailureFreezePane" /> 类型的实例。
    /// </summary>
    /// <param name="columnSplit">冻结前的列数。</param>
    /// <param name="rowSplit">冻结前的行数。</param>
    /// <param name="leftColumn">冻结后左侧可见列。</param>
    /// <param name="topRow">冻结后顶部可见行。</param>
    internal ExcelDataReaderFailureFreezePane(int columnSplit, int rowSplit,
        int leftColumn, int topRow)
    {
        ColumnSplit = columnSplit;
        RowSplit = rowSplit;
        LeftColumn = leftColumn;
        TopRow = topRow;
    }

    /// <summary>
    /// 获取冻结前的列数。
    /// </summary>
    internal int ColumnSplit { get; }

    /// <summary>
    /// 获取冻结前的行数。
    /// </summary>
    internal int RowSplit { get; }

    /// <summary>
    /// 获取冻结后左侧可见列。
    /// </summary>
    internal int LeftColumn { get; }

    /// <summary>
    /// 获取冻结后顶部可见行。
    /// </summary>
    internal int TopRow { get; }
}

/// <summary>
/// 保存一个条件格式区域及其规则。
/// </summary>
internal sealed class ExcelDataReaderFailureConditionalFormatting
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderFailureConditionalFormatting" /> 类型的实例。
    /// </summary>
    /// <param name="ranges">条件格式适用区域。</param>
    /// <param name="rules">条件格式规则。</param>
    internal ExcelDataReaderFailureConditionalFormatting(
        IReadOnlyList<ExcelDataReaderFailureConditionalFormattingRange> ranges,
        IReadOnlyList<ExcelDataReaderFailureConditionalFormattingRule> rules)
    {
        Ranges = ranges == null
            ? Array.Empty<ExcelDataReaderFailureConditionalFormattingRange>()
            : ranges.Where(item => item != null).ToArray();
        Rules = rules == null
            ? Array.Empty<ExcelDataReaderFailureConditionalFormattingRule>()
            : rules.Where(item => item != null).ToArray();
    }

    /// <summary>
    /// 获取条件格式区域。
    /// </summary>
    internal IReadOnlyList<ExcelDataReaderFailureConditionalFormattingRange> Ranges { get; }

    /// <summary>
    /// 获取条件格式规则。
    /// </summary>
    internal IReadOnlyList<ExcelDataReaderFailureConditionalFormattingRule> Rules { get; }
}

/// <summary>
/// 保存条件格式的零基单元格区域。
/// </summary>
internal sealed class ExcelDataReaderFailureConditionalFormattingRange
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderFailureConditionalFormattingRange" /> 类型的实例。
    /// </summary>
    /// <param name="firstRow">首行。</param>
    /// <param name="lastRow">末行。</param>
    /// <param name="firstColumn">首列。</param>
    /// <param name="lastColumn">末列。</param>
    internal ExcelDataReaderFailureConditionalFormattingRange(int firstRow, int lastRow,
        int firstColumn, int lastColumn)
    {
        FirstRow = firstRow;
        LastRow = lastRow;
        FirstColumn = firstColumn;
        LastColumn = lastColumn;
    }

    /// <summary>
    /// 获取首行。
    /// </summary>
    internal int FirstRow { get; }

    /// <summary>
    /// 获取末行。
    /// </summary>
    internal int LastRow { get; }

    /// <summary>
    /// 获取首列。
    /// </summary>
    internal int FirstColumn { get; }

    /// <summary>
    /// 获取末列。
    /// </summary>
    internal int LastColumn { get; }
}

/// <summary>
/// 保存条件格式的一条受限规则。
/// </summary>
internal sealed class ExcelDataReaderFailureConditionalFormattingRule
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderFailureConditionalFormattingRule" /> 类型的实例。
    /// </summary>
    /// <param name="conditionType">条件类型。</param>
    /// <param name="comparisonOperator">比较运算符。</param>
    /// <param name="formula1">第一个公式。</param>
    /// <param name="formula2">第二个公式。</param>
    /// <param name="formattingStyle">条件格式样式。</param>
    internal ExcelDataReaderFailureConditionalFormattingRule(int conditionType,
        int comparisonOperator, string formula1, string formula2,
        ExcelDataReaderFailureFontStyle formattingStyle)
    {
        ConditionType = conditionType;
        ComparisonOperator = comparisonOperator;
        Formula1 = formula1;
        Formula2 = formula2;
        FormattingStyle = formattingStyle;
    }

    /// <summary>
    /// 获取条件类型。
    /// </summary>
    internal int ConditionType { get; }

    /// <summary>
    /// 获取比较运算符。
    /// </summary>
    internal int ComparisonOperator { get; }

    /// <summary>
    /// 获取第一个公式。
    /// </summary>
    internal string Formula1 { get; }

    /// <summary>
    /// 获取第二个公式。
    /// </summary>
    internal string Formula2 { get; }

    /// <summary>
    /// 获取条件格式样式。
    /// </summary>
    internal ExcelDataReaderFailureFontStyle FormattingStyle { get; }
}

/// <summary>
/// 保存一个 HSSF 原生数据校验区域的低保留重建信息。
/// </summary>
internal sealed class ExcelDataReaderFailureNativeDataValidation
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderFailureNativeDataValidation" /> 类型的实例。
    /// </summary>
    /// <param name="firstRow">零基首行。</param>
    /// <param name="lastRow">零基末行。</param>
    /// <param name="firstColumn">零基首列。</param>
    /// <param name="lastColumn">零基末列。</param>
    /// <param name="validationType">原生校验类型。</param>
    /// <param name="comparisonOperator">比较运算符。</param>
    /// <param name="formula1">第一个常量公式。</param>
    /// <param name="formula2">第二个常量公式。</param>
    /// <param name="emptyCellAllowed">是否允许空单元格。</param>
    /// <param name="suppressDropdownArrow">是否隐藏下拉箭头。</param>
    /// <param name="showPromptBox">是否显示输入提示。</param>
    /// <param name="promptTitle">输入提示标题。</param>
    /// <param name="promptText">输入提示内容。</param>
    /// <param name="showErrorBox">是否显示错误提示。</param>
    /// <param name="errorTitle">错误提示标题。</param>
    /// <param name="errorText">错误提示内容。</param>
    /// <param name="errorStyle">错误样式。</param>
    /// <param name="listSourceName">命名列表来源；非列表校验为 null。</param>
    /// <param name="listSourceLocalSheetName">命名列表的局部作用域工作表；全局名称为 null。</param>
    /// <param name="listValues">显式列表值；名称列表或非列表校验为空。</param>
    internal ExcelDataReaderFailureNativeDataValidation(int firstRow, int lastRow,
        int firstColumn, int lastColumn, int validationType, int comparisonOperator,
        string formula1, string formula2, bool emptyCellAllowed,
        bool suppressDropdownArrow, bool showPromptBox, string promptTitle,
        string promptText, bool showErrorBox, string errorTitle, string errorText,
        int errorStyle, string listSourceName = null,
        string listSourceLocalSheetName = null, IReadOnlyList<string> listValues = null)
    {
        FirstRow = firstRow;
        LastRow = lastRow;
        FirstColumn = firstColumn;
        LastColumn = lastColumn;
        ValidationType = validationType;
        ComparisonOperator = comparisonOperator;
        Formula1 = formula1;
        Formula2 = formula2;
        EmptyCellAllowed = emptyCellAllowed;
        SuppressDropdownArrow = suppressDropdownArrow;
        ShowPromptBox = showPromptBox;
        PromptTitle = promptTitle;
        PromptText = promptText;
        ShowErrorBox = showErrorBox;
        ErrorTitle = errorTitle;
        ErrorText = errorText;
        ErrorStyle = errorStyle;
        ListSourceName = listSourceName;
        ListSourceLocalSheetName = listSourceLocalSheetName;
        ListValues = listValues == null
            ? Array.Empty<string>() : listValues.Where(value => value != null).ToArray();
    }

    /// <summary>获取零基首行。</summary>
    internal int FirstRow { get; }

    /// <summary>获取零基末行。</summary>
    internal int LastRow { get; }

    /// <summary>获取零基首列。</summary>
    internal int FirstColumn { get; }

    /// <summary>获取零基末列。</summary>
    internal int LastColumn { get; }

    /// <summary>获取原生校验类型。</summary>
    internal int ValidationType { get; }

    /// <summary>获取比较运算符。</summary>
    internal int ComparisonOperator { get; }

    /// <summary>获取第一个常量公式。</summary>
    internal string Formula1 { get; }

    /// <summary>获取第二个常量公式。</summary>
    internal string Formula2 { get; }

    /// <summary>获取是否允许空单元格。</summary>
    internal bool EmptyCellAllowed { get; }

    /// <summary>获取是否隐藏下拉箭头。</summary>
    internal bool SuppressDropdownArrow { get; }

    /// <summary>获取是否显示输入提示。</summary>
    internal bool ShowPromptBox { get; }

    /// <summary>获取输入提示标题。</summary>
    internal string PromptTitle { get; }

    /// <summary>获取输入提示内容。</summary>
    internal string PromptText { get; }

    /// <summary>获取是否显示错误提示。</summary>
    internal bool ShowErrorBox { get; }

    /// <summary>获取错误提示标题。</summary>
    internal string ErrorTitle { get; }

    /// <summary>获取错误提示内容。</summary>
    internal string ErrorText { get; }

    /// <summary>获取错误样式。</summary>
    internal int ErrorStyle { get; }

    /// <summary>获取命名列表来源。</summary>
    internal string ListSourceName { get; }

    /// <summary>获取命名列表的局部作用域工作表。</summary>
    internal string ListSourceLocalSheetName { get; }

    /// <summary>获取显式列表值。</summary>
    internal IReadOnlyList<string> ListValues { get; }
}

/// <summary>
/// 保存失败工作簿行的展示元数据。
/// </summary>
internal sealed class ExcelDataReaderFailureRowMetadata
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderFailureRowMetadata" /> 类型的实例。
    /// </summary>
    /// <param name="height">行高的 HSSF 原始单位。</param>
    /// <param name="hasHeight">是否包含显式行高。</param>
    /// <param name="hidden">是否隐藏该行。</param>
    internal ExcelDataReaderFailureRowMetadata(short height, bool hasHeight, bool hidden)
    {
        Height = height;
        HasHeight = hasHeight;
        Hidden = hidden;
    }

    /// <summary>
    /// 获取行高的 HSSF 原始单位。
    /// </summary>
    internal short Height { get; }

    /// <summary>
    /// 获取是否包含显式行高。
    /// </summary>
    internal bool HasHeight { get; }

    /// <summary>
    /// 获取行是否隐藏。
    /// </summary>
    internal bool Hidden { get; }
}

/// <summary>
/// 保存失败工作簿的合并区域元数据。
/// </summary>
internal sealed class ExcelDataReaderFailureMergeRegion
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderFailureMergeRegion" /> 类型的实例。
    /// </summary>
    /// <param name="firstRow">首行的零基索引。</param>
    /// <param name="lastRow">末行的零基索引。</param>
    /// <param name="firstColumn">首列的零基索引。</param>
    /// <param name="lastColumn">末列的零基索引。</param>
    internal ExcelDataReaderFailureMergeRegion(int firstRow, int lastRow,
        int firstColumn, int lastColumn)
    {
        FirstRow = firstRow;
        LastRow = lastRow;
        FirstColumn = firstColumn;
        LastColumn = lastColumn;
    }

    /// <summary>
    /// 获取首行的零基索引。
    /// </summary>
    internal int FirstRow { get; }

    /// <summary>
    /// 获取末行的零基索引。
    /// </summary>
    internal int LastRow { get; }

    /// <summary>
    /// 获取首列的零基索引。
    /// </summary>
    internal int FirstColumn { get; }

    /// <summary>
    /// 获取末列的零基索引。
    /// </summary>
    internal int LastColumn { get; }
}

/// <summary>
/// 保存失败工作簿列的展示元数据。
/// </summary>
internal sealed class ExcelDataReaderFailureColumnMetadata
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderFailureColumnMetadata" /> 类型的实例。
    /// </summary>
    /// <param name="width">列宽。</param>
    /// <param name="hasWidth">是否包含显式列宽。</param>
    /// <param name="hidden">是否隐藏该列。</param>
    internal ExcelDataReaderFailureColumnMetadata(double width, bool hasWidth, bool hidden)
    {
        Width = width;
        HasWidth = hasWidth;
        Hidden = hidden;
    }

    /// <summary>
    /// 获取列宽。
    /// </summary>
    internal double Width { get; }

    /// <summary>
    /// 获取是否包含显式列宽。
    /// </summary>
    internal bool HasWidth { get; }

    /// <summary>
    /// 获取列是否隐藏。
    /// </summary>
    internal bool Hidden { get; }
}

/// <summary>
/// 一个失败数据行的原始值和错误快照。
/// </summary>
internal sealed class ExcelDataReaderFailureRow
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderFailureRow" /> 类型的实例。
    /// </summary>
    /// <param name="rowIndex">从一开始的源行号。</param>
    /// <param name="values">当前行的原始单元格值。</param>
    /// <param name="errors">当前行的结构化错误。</param>
    /// <param name="numberFormats">失败行单元格的数字格式。</param>
    /// <param name="boldColumns">失败行单元格的加粗列。</param>
    /// <param name="fontStyles">失败行单元格的字体状态。</param>
    internal ExcelDataReaderFailureRow(int rowIndex, IReadOnlyList<object> values,
        IReadOnlyList<ExcelImportError> errors,
        IReadOnlyDictionary<int, string> numberFormats = null,
        IReadOnlyDictionary<int, bool> boldColumns = null,
        IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> fontStyles = null)
    {
        RowIndex = rowIndex;
        Values = (values ?? Array.Empty<object>()).ToArray();
        Errors = (errors ?? Array.Empty<ExcelImportError>()).ToArray();
        NumberFormats = ExcelDataReaderFailureMetadata.CopyNumberFormats(numberFormats);
        BoldColumns = ExcelDataReaderFailureMetadata.CopyBoldColumns(boldColumns);
        FontStyles = ExcelDataReaderFailureMetadata.CopyFontStyles(fontStyles);
    }

    /// <summary>
    /// 获取源行号。
    /// </summary>
    internal int RowIndex { get; }

    /// <summary>
    /// 获取原始单元格值。
    /// </summary>
    internal IReadOnlyList<object> Values { get; }

    /// <summary>
    /// 获取当前行的结构化错误。
    /// </summary>
    internal IReadOnlyList<ExcelImportError> Errors { get; }

    /// <summary>
    /// 获取单元格的数字格式。
    /// </summary>
    internal IReadOnlyDictionary<int, string> NumberFormats { get; }

    /// <summary>
    /// 获取单元格的加粗列。
    /// </summary>
    internal IReadOnlyDictionary<int, bool> BoldColumns { get; }

    /// <summary>
    /// 获取单元格的字体状态。
    /// </summary>
    internal IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> FontStyles { get; }

}

/// <summary>
/// 复制失败工作簿元数据的内部辅助方法。
/// </summary>
internal static class ExcelDataReaderFailureMetadata
{
    /// <summary>
    /// 复制并过滤数字格式映射。
    /// </summary>
    /// <param name="numberFormats">源单元格数字格式。</param>
    /// <returns>隔离后的数字格式映射。</returns>
    internal static IReadOnlyDictionary<int, string> CopyNumberFormats(
        IReadOnlyDictionary<int, string> numberFormats)
    {
        if (numberFormats == null || numberFormats.Count == 0)
            return new Dictionary<int, string>();
        return numberFormats.Where(pair => pair.Key >= 0 && !string.IsNullOrWhiteSpace(pair.Value))
            .ToDictionary(pair => pair.Key, pair => pair.Value, EqualityComparer<int>.Default);
    }

    /// <summary>
    /// 复制并过滤加粗列映射。
    /// </summary>
    /// <param name="boldColumns">源单元格加粗列。</param>
    /// <returns>隔离后的加粗列映射。</returns>
    internal static IReadOnlyDictionary<int, bool> CopyBoldColumns(
        IReadOnlyDictionary<int, bool> boldColumns)
    {
        if (boldColumns == null || boldColumns.Count == 0)
            return new Dictionary<int, bool>();
        return boldColumns.Where(pair => pair.Key >= 0 && pair.Value)
            .ToDictionary(pair => pair.Key, pair => true, EqualityComparer<int>.Default);
    }

    /// <summary>
    /// 复制并过滤字体状态映射。
    /// </summary>
    /// <param name="fontStyles">源单元格字体状态。</param>
    /// <returns>隔离后的字体状态映射。</returns>
    internal static IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> CopyFontStyles(
        IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> fontStyles)
    {
        if (fontStyles == null || fontStyles.Count == 0)
            return new Dictionary<int, ExcelDataReaderFailureFontStyle>();
        return fontStyles
            .Where(pair => pair.Key >= 0 && pair.Value != null)
            .ToDictionary(pair => pair.Key,
                pair => pair.Value,
                EqualityComparer<int>.Default);
    }
}

/// <summary>
/// 保存失败工作簿单元格的最小字体、对齐、填充、边框和保护状态。
/// </summary>
internal sealed class ExcelDataReaderFailureFontStyle
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderFailureFontStyle" /> 类型的实例。
    /// </summary>
    /// <param name="isBold">是否粗体。</param>
    /// <param name="isItalic">是否斜体。</param>
    /// <param name="isStrikeout">是否删除线。</param>
    /// <param name="isUnderline">是否单下划线。</param>
    /// <param name="fontName">字体名称。</param>
    /// <param name="fontHeightInPoints">字号；无效时使用默认字号。</param>
    /// <param name="fontColorIndex">字体索引色；未设置时使用默认颜色。</param>
    /// <param name="horizontalAlignment">水平对齐名称。</param>
    /// <param name="verticalAlignment">垂直对齐名称。</param>
    /// <param name="wrapText">是否自动换行。</param>
    /// <param name="shrinkToFit">是否缩小字体适应单元格。</param>
    /// <param name="textRotation">文本旋转角度。</param>
    /// <param name="indentation">文本缩进量。</param>
    /// <param name="fillPattern">填充模式；当前仅保留纯色前景填充。</param>
    /// <param name="fillForegroundColorIndex">填充前景索引色。</param>
    /// <param name="fillBackgroundColorIndex">填充背景索引色。</param>
    /// <param name="borderLeft">左边框状态。</param>
    /// <param name="borderRight">右边框状态。</param>
    /// <param name="borderTop">上边框状态。</param>
    /// <param name="borderBottom">下边框状态。</param>
    /// <param name="isLocked">是否锁定单元格；未设置时使用工作簿默认值。</param>
    /// <param name="isHidden">是否隐藏单元格公式；未设置时使用工作簿默认值。</param>
    internal ExcelDataReaderFailureFontStyle(bool isBold, bool isItalic,
        bool isStrikeout, bool isUnderline, string fontName = null,
        double fontHeightInPoints = 0, short? fontColorIndex = null,
        string horizontalAlignment = null, string verticalAlignment = null,
        bool wrapText = false, bool shrinkToFit = false,
        short? textRotation = null, short indentation = 0,
        string fillPattern = null, short? fillForegroundColorIndex = null,
        short? fillBackgroundColorIndex = null,
        ExcelDataReaderFailureBorderSide borderLeft = null,
        ExcelDataReaderFailureBorderSide borderRight = null,
        ExcelDataReaderFailureBorderSide borderTop = null,
        ExcelDataReaderFailureBorderSide borderBottom = null,
        bool? isLocked = null, bool? isHidden = null)
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
        HorizontalAlignment = string.IsNullOrWhiteSpace(horizontalAlignment)
            ? null : horizontalAlignment.Trim();
        VerticalAlignment = string.IsNullOrWhiteSpace(verticalAlignment)
            ? null : verticalAlignment.Trim();
        WrapText = wrapText;
        ShrinkToFit = shrinkToFit;
        TextRotation = textRotation >= 0 && textRotation <= 180 ? textRotation : null;
        Indentation = indentation >= 0 && indentation <= 250 ? indentation : (short)0;
        FillPattern = string.IsNullOrWhiteSpace(fillPattern) ? null : fillPattern.Trim();
        FillForegroundColorIndex = fillForegroundColorIndex >= 0 ? fillForegroundColorIndex : null;
        FillBackgroundColorIndex = fillBackgroundColorIndex >= 0 ? fillBackgroundColorIndex : null;
        BorderLeft = borderLeft;
        BorderRight = borderRight;
        BorderTop = borderTop;
        BorderBottom = borderBottom;
        IsLocked = isLocked;
        IsHidden = isHidden;
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
    /// 获取水平对齐名称。
    /// </summary>
    internal string HorizontalAlignment { get; }

    /// <summary>
    /// 获取垂直对齐名称。
    /// </summary>
    internal string VerticalAlignment { get; }

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
    internal short? TextRotation { get; }

    /// <summary>
    /// 获取文本缩进量。
    /// </summary>
    internal short Indentation { get; }

    /// <summary>
    /// 获取填充模式。
    /// </summary>
    internal string FillPattern { get; }

    /// <summary>
    /// 获取填充前景索引色。
    /// </summary>
    internal short? FillForegroundColorIndex { get; }

    /// <summary>
    /// 获取填充背景索引色。
    /// </summary>
    internal short? FillBackgroundColorIndex { get; }

    /// <summary>
    /// 获取左边框状态。
    /// </summary>
    internal ExcelDataReaderFailureBorderSide BorderLeft { get; }

    /// <summary>
    /// 获取右边框状态。
    /// </summary>
    internal ExcelDataReaderFailureBorderSide BorderRight { get; }

    /// <summary>
    /// 获取上边框状态。
    /// </summary>
    internal ExcelDataReaderFailureBorderSide BorderTop { get; }

    /// <summary>
    /// 获取下边框状态。
    /// </summary>
    internal ExcelDataReaderFailureBorderSide BorderBottom { get; }

    /// <summary>
    /// 获取是否锁定单元格。
    /// </summary>
    internal bool? IsLocked { get; }

    /// <summary>
    /// 获取是否隐藏单元格公式。
    /// </summary>
    internal bool? IsHidden { get; }
}

/// <summary>
/// 保存失败工作簿单元格的一侧边框状态。
/// </summary>
internal sealed class ExcelDataReaderFailureBorderSide
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderFailureBorderSide" /> 类型的实例。
    /// </summary>
    /// <param name="style">边框线型名称。</param>
    /// <param name="colorIndex">边框索引色。</param>
    internal ExcelDataReaderFailureBorderSide(string style, short? colorIndex = null)
    {
        Style = string.IsNullOrWhiteSpace(style) ? null : style.Trim();
        ColorIndex = colorIndex >= 0 ? colorIndex : null;
    }

    /// <summary>
    /// 获取边框线型名称。
    /// </summary>
    internal string Style { get; }

    /// <summary>
    /// 获取边框索引色。
    /// </summary>
    internal short? ColorIndex { get; }
}
