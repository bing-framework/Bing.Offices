using Bing.Offices.Exports;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Internals;
using Bing.Offices.Npoi.Resolvers;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using NPOI.SS.Util;
using NPOI.SS.UserModel.Charts;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Exports;

/// <summary>
/// 基于 NPOI 的单工作簿 Excel 导出器；NPOI 工作簿在内存中构建后写入目标流。
/// </summary>
internal sealed class NpoiExcelExporter : IExcelExporter
{
    /// <summary>
    /// 当前导出器使用的值转换器。
    /// </summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;
    private readonly NpoiExportPlanBuilder _planBuilder;
    private readonly NpoiExportSheetWriter _sheetWriter;

    /// <summary>
    /// 初始化一个<see cref="NpoiExcelExporter"/>类型的实例。
    /// </summary>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="mappingPlanFactory">方向化映射计划工厂。</param>
    public NpoiExcelExporter(IEnumerable<IExcelValueConverter> valueConverters = null,
        IExcelMappingPlanFactory mappingPlanFactory = null)
    {
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _planBuilder = new NpoiExportPlanBuilder(mappingPlanFactory ?? NpoiMappingPlanFactoryResolver.CreateDefault(
            _valueConverters));
        _sheetWriter = new NpoiExportSheetWriter();
    }

    /// <inheritdoc />
    public void Export(ExcelWorkbookExportRequest request, Stream destination,
        CancellationToken cancellationToken = default)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (destination == null)
            throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite)
            throw new ArgumentException("目标流不可写入。", nameof(destination));
        cancellationToken.ThrowIfCancellationRequested();

        try
        {
            using var workbook = CreateWorkbook(request);
            var planBySheet = _planBuilder.Create(request);
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sheetRequest in request.Sheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ValidateSheetName(sheetRequest.Name);
                if (!names.Add(sheetRequest.Name))
                    throw new ArgumentException($"Workbook 包含重复 Sheet 名称: {sheetRequest.Name}");
                WriteSheet(workbook, sheetRequest, request.Template != null, cancellationToken,
                    planBySheet[sheetRequest]);
            }

            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                workbook.Write(new NonDisposingStream(destination, cancellationToken), false);
            }
            catch (Exception exception) when (cancellationToken.IsCancellationRequested
                && exception.GetBaseException() is OperationCanceledException)
            {
                throw new OperationCanceledException(cancellationToken);
            }
            cancellationToken.ThrowIfCancellationRequested();
        }
        finally
        {
            if (request.Template != null && !request.LeaveTemplateOpen)
                request.Template.Dispose();
        }
    }

    /// <summary>
    /// 创建普通或模板工作簿。模板加载后沿用同一 Sheet Writer。
    /// </summary>
    private static NPOI.SS.UserModel.IWorkbook CreateWorkbook(ExcelWorkbookExportRequest request)
    {
        if (request.Template == null)
            return ExcelHelper.PrepareWorkbook(request.Format, request.Metadata);
        if (!request.Template.CanRead)
            throw new ArgumentException("模板流不可读取。", nameof(request));
        var workbook = NPOI.SS.UserModel.WorkbookFactory.Create(new NonDisposingStream(request.Template));
        if (request.MetadataSpecified)
            ExcelHelper.ApplyWorkbookMetadata(workbook, request.Metadata);
        return workbook;
    }

    /// <summary>
    /// 验证 Excel Sheet 名称边界。
    /// </summary>
    private static void ValidateSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("工作表名称不能为空。", nameof(name));
        if (name.Length > 31)
            throw new ArgumentException("工作表名称不能超过 31 个字符。", nameof(name));
        if (name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0)
            throw new ArgumentException($"工作表名称包含非法字符: {name}", nameof(name));
    }

    private void WriteSheet(NPOI.SS.UserModel.IWorkbook workbook, ExcelSheetExportRequest request,
        bool isTemplate, CancellationToken cancellationToken, IExcelMappingPlan mapping)
    {
        var method = GetType().GetMethod(nameof(WriteTypedSheet), BindingFlags.Instance | BindingFlags.NonPublic);
        try
        {
            method.MakeGenericMethod(request.ItemType).Invoke(this,
                new object[] { workbook, request, isTemplate, cancellationToken, mapping });
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }

    /// <summary>
    /// 执行一个泛型 Sheet 的统一列计划和 Cell Writer。
    /// </summary>
    private void WriteTypedSheet<T>(NPOI.SS.UserModel.IWorkbook workbook, ExcelSheetExportRequest request,
        bool isTemplate, CancellationToken cancellationToken, IExcelMappingPlan map) where T : class, new()
    {
        var sheet = workbook.GetSheet(request.Name);
        if (sheet == null && isTemplate)
            throw new InvalidOperationException($"模板缺少请求的 Sheet: {request.Name}");
        sheet ??= workbook.CreateSheet(request.Name);
        if (request.Hidden)
            workbook.SetSheetVisibility(workbook.GetSheetIndex(sheet), NPOI.SS.UserModel.SheetVisibility.Hidden);
        if (request.HeaderRowIndex < 0 || request.DataRowStartIndex <= request.HeaderRowIndex)
            throw new ArgumentOutOfRangeException(nameof(request.DataRowStartIndex));

        var templateOrigin = ResolveTemplateOrigin(workbook, sheet, request, isTemplate);
        var headerRowIndex = templateOrigin.Row + request.HeaderRowIndex;
        var firstColumnIndex = templateOrigin.Column;
        var dynamicDefinitions = map.DynamicColumns.Select(column => CreateDynamicDefinition(column,
            request.DynamicColumns.FirstOrDefault(item => string.Equals(item.Key, column.Key,
                StringComparison.OrdinalIgnoreCase)), map)).ToArray();
        ValidateDynamicDefinitions(dynamicDefinitions);
        var columns = CreateColumns<T>(map, dynamicDefinitions);
        ValidateDynamicColumns(dynamicDefinitions, columns);
        _sheetWriter.Write<T>(workbook, request, cancellationToken, map, columns, templateOrigin.Row,
            firstColumnIndex);
    }

    /// <summary>
    /// 解析模板命名区域的起点；普通导出从零坐标开始。
    /// </summary>
    private static (int Row, int Column) ResolveTemplateOrigin(NPOI.SS.UserModel.IWorkbook workbook,
        NPOI.SS.UserModel.ISheet sheet, ExcelSheetExportRequest request, bool isTemplate)
    {
        if (!isTemplate || string.IsNullOrWhiteSpace(request.TemplateRegion))
            return (0, 0);
        var name = workbook.GetName(request.TemplateRegion);
        if (name == null || string.IsNullOrWhiteSpace(name.RefersToFormula))
            throw new InvalidOperationException($"模板缺少命名区域: {request.TemplateRegion}");
        var formula = name.RefersToFormula.TrimStart('=');
        var separator = formula.LastIndexOf('!');
        if (separator < 0)
            throw new InvalidOperationException($"模板命名区域缺少 Sheet 引用: {request.TemplateRegion}");
        var sheetName = formula.Substring(0, separator).Trim('\'', ' ');
        if (!string.Equals(sheetName, sheet.SheetName, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException($"模板命名区域不属于请求 Sheet: {request.TemplateRegion}");
        var address = formula.Substring(separator + 1).Split(':')[0].Replace("$", string.Empty);
        var match = Regex.Match(address, "^([A-Za-z]+)([0-9]+)$");
        if (!match.Success)
            throw new InvalidOperationException($"模板命名区域地址无效: {request.TemplateRegion}");
        var column = 0;
        foreach (var character in match.Groups[1].Value.ToUpperInvariant())
            column = column * 26 + character - 'A' + 1;
        return (int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) - 1, column - 1);
    }

    /// <summary>
    /// 检查动态列定义和列键唯一性。
    /// </summary>
    private static void ValidateDynamicColumns(IReadOnlyList<ExcelDynamicColumnDefinition> definitions,
        IReadOnlyList<ExcelColumnPlan> columns)
    {
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var titles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in columns.Where(column => !column.IsDynamic))
            titles.Add(column.Title);
        foreach (var definition in definitions ?? Array.Empty<ExcelDynamicColumnDefinition>())
        {
            if (!keys.Add(definition.Key))
                throw new ArgumentException($"动态列包含重复 Key: {definition.Key}", nameof(definitions));
            if (!titles.Add(definition.Title))
                throw new ArgumentException($"动态列包含重复标题: {definition.Title}", nameof(definitions));
            foreach (var alias in definition.Aliases ?? Array.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(alias) || !titles.Add(alias))
                    throw new ArgumentException($"动态列包含重复或空标题别名: {alias}", nameof(definitions));
            }
            if (definition.PhysicalColumnIndex.HasValue && definition.Placement != null)
                throw new ArgumentException($"动态列 {definition.Key} 不能同时指定相对位置和物理索引。",
                    nameof(definitions));
            if (definition.Placement?.PhysicalColumnIndex != null && definition.PhysicalColumnIndex.HasValue)
                throw new ArgumentException($"动态列 {definition.Key} 不能重复指定物理索引。",
                    nameof(definitions));
        }
        ValidateColumns(columns);
    }

    private static void ValidateDynamicDefinitions(IReadOnlyList<ExcelDynamicColumnDefinition> definitions)
    {
        foreach (var definition in definitions ?? Array.Empty<ExcelDynamicColumnDefinition>())
        {
            if (definition == null || string.IsNullOrWhiteSpace(definition.Key)
                || string.IsNullOrWhiteSpace(definition.Title))
                throw new ArgumentException("动态列 Key 和 Title 不能为空。", nameof(definitions));
        }
    }

    /// <summary>
    /// 创建 Workbook 请求使用的固定列和 typed 动态列计划。
    /// </summary>
    private static IReadOnlyList<ExcelColumnPlan> CreateColumns<T>(IExcelMappingPlan typeMap,
        IReadOnlyList<ExcelDynamicColumnDefinition> dynamicColumns)
        where T : class, new()
    {
        var fixedColumns = new List<ExcelColumnPlan>();
        foreach (var property in typeMap.Columns)
        {
            if (property.Ignored)
                continue;
            if (!property.IsDynamicColumn)
            {
                var reflectionProperty = ResolveProperty<T>(property.Name);
                fixedColumns.Add(new ExcelColumnPlan(property.Title, property, false, -1, null, null,
                    property.ValueConverters, property.ValidationBindings,
                    reflectionProperty: reflectionProperty));
                continue;
            }
        }
        var columns = fixedColumns.ToList();
        var dynamicProperty = typeMap.Columns.FirstOrDefault(property => property.IsDynamicColumn);
        if (dynamicProperty == null)
            return columns;
        var dynamicReflectionProperty = ResolveProperty<T>(dynamicProperty.Name);
        var definitions = (dynamicColumns ?? Array.Empty<ExcelDynamicColumnDefinition>())
            .OrderBy(definition => definition.Order)
            .ThenBy(definition => definition.Key, StringComparer.Ordinal)
            .ToList();
        foreach (var definition in definitions)
        {
            var dynamicPlan = typeMap.DynamicColumns.Single(column => string.Equals(column.Key, definition.Key,
                StringComparison.OrdinalIgnoreCase));
            var column = new ExcelColumnPlan(definition.Title, dynamicProperty, true, -1, definition, definition.Key,
                dynamicPlan.ValueConverters, dynamicPlan.ValidationBindings,
                reflectionProperty: dynamicReflectionProperty,
                isUnique: dynamicPlan.IsUnique,
                uniqueIgnoreEmpty: dynamicPlan.UniqueIgnoreEmpty);
            var placement = definition.Placement;
            var physicalIndex = definition.PhysicalColumnIndex ?? placement?.PhysicalColumnIndex;
            if (physicalIndex.HasValue)
            {
                if (physicalIndex.Value > columns.Count)
                    throw new ArgumentOutOfRangeException(nameof(definition.PhysicalColumnIndex),
                        $"动态列 {definition.Key} 的物理索引超出当前列计划。");
                columns.Insert(physicalIndex.Value, column);
                continue;
            }
            if (placement != null && !string.IsNullOrWhiteSpace(placement.BeforeKey))
            {
                var index = columns.FindIndex(item => string.Equals(item.Key, placement.BeforeKey,
                    StringComparison.OrdinalIgnoreCase));
                if (index < 0)
                    throw new ArgumentException($"动态列 {definition.Key} 的 Before 目标不存在: {placement.BeforeKey}");
                columns.Insert(index, column);
                continue;
            }
            if (placement != null && !string.IsNullOrWhiteSpace(placement.AfterKey))
            {
                var index = columns.FindIndex(item => string.Equals(item.Key, placement.AfterKey,
                    StringComparison.OrdinalIgnoreCase));
                if (index < 0)
                    throw new ArgumentException($"动态列 {definition.Key} 的 After 目标不存在: {placement.AfterKey}");
                columns.Insert(index + 1, column);
                continue;
            }
            columns.Add(column);
        }
        return columns;
    }

    private static ExcelDynamicColumnDefinition CreateDynamicDefinition(IExcelDynamicMappingColumn column,
        ExcelDynamicColumnDefinition requestColumn, IExcelMappingPlan mapping)
    {
        var columnIndex = column.ColumnIndex;
        var placementKey = column.PlacementKey;
        if (!columnIndex.HasValue && string.IsNullOrWhiteSpace(placementKey)
            && mapping.DynamicColumns.Count == 1)
        {
            columnIndex = mapping.Layout?.ColumnIndex;
            placementKey = mapping.Layout?.PlacementKey;
        }
        return new ExcelDynamicColumnDefinition
        {
            Key = column.Key,
            Title = column.Title,
            Aliases = column.Aliases,
            DataType = ResolveDynamicType(column.DataTypeName),
            Order = column.Order,
            Placement = CreatePlacement(placementKey),
            PhysicalColumnIndex = columnIndex,
            NumberFormat = column.NumberFormat ?? mapping.Style?.NumberFormat,
            HeaderStyle = requestColumn?.HeaderStyle,
            BodyStyle = requestColumn?.BodyStyle,
            ConverterName = column.ConverterName,
            ValidatorName = column.ValidatorName,
            ValidationRuleNames = column.ValidationRuleNames,
            ImageMultiplicity = column.ImageMultiplicity
        };
    }

    private static ExcelColumnPlacement CreatePlacement(string placementKey)
    {
        if (string.IsNullOrWhiteSpace(placementKey))
            return null;
        var key = placementKey.Substring(placementKey.IndexOfAny(new[] { ':', '-' }) + 1);
        return placementKey.StartsWith("before:", StringComparison.OrdinalIgnoreCase)
            || placementKey.StartsWith("before-", StringComparison.OrdinalIgnoreCase)
            ? ExcelColumnPlacement.Before(key)
            : ExcelColumnPlacement.After(key);
    }

    private static Type ResolveDynamicType(string name)
    {
        switch ((name ?? "string").ToLowerInvariant())
        {
            case "object": return typeof(object);
            case "string": return typeof(string);
            case "boolean": case "bool": return typeof(bool);
            case "byte": return typeof(byte);
            case "int16": return typeof(short);
            case "int32": case "int": return typeof(int);
            case "int64": case "long": return typeof(long);
            case "single": case "float": return typeof(float);
            case "double": return typeof(double);
            case "decimal": return typeof(decimal);
            case "datetime": return typeof(DateTime);
            case "datetimeoffset": return typeof(DateTimeOffset);
            case "guid": return typeof(Guid);
            case "bytes": return typeof(byte[]);
            default: throw new InvalidOperationException($"动态列数据类型不在允许列表中: {name}");
        }
    }

    /// <summary>
    /// 验证解析后的列标题唯一，确保导出结果可被导入器无歧义读取。
    /// </summary>
    /// <param name="columns">当前请求的导出列。</param>
    private static void ValidateColumns(IReadOnlyList<ExcelColumnPlan> columns)
    {
        var titles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in columns)
        {
            if (!titles.Add(column.Title))
                throw new ArgumentException($"导出列标题重复: {column.Title}", nameof(columns));
        }
    }

    private static PropertyInfo ResolveProperty<T>(string name) where T : class, new()
    {
        var property = typeof(T).GetProperty(name, BindingFlags.Instance | BindingFlags.Public);
        if (property == null)
            throw new InvalidOperationException($"无法解析映射属性: {name}");
        return property;
    }

    /// <summary>
    /// 让 NPOI 可以释放包装器但不能关闭调用方拥有的内部缓冲流。
    /// </summary>
    private sealed class NonDisposingStream : Stream
    {
        private readonly Stream _inner;
        private readonly CancellationToken _cancellationToken;

        public NonDisposingStream(Stream inner, CancellationToken cancellationToken = default)
        {
            _inner = inner;
            _cancellationToken = cancellationToken;
        }

        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => _inner.CanSeek;
        public override bool CanWrite => _inner.CanWrite;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
        public override void Flush()
        {
            _cancellationToken.ThrowIfCancellationRequested();
            _inner.Flush();
        }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        public override void SetLength(long value) => _inner.SetLength(value);
        public override void Write(byte[] buffer, int offset, int count)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            _inner.Write(buffer, offset, count);
            _cancellationToken.ThrowIfCancellationRequested();
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Flush();
            base.Dispose(disposing);
        }
    }

}
