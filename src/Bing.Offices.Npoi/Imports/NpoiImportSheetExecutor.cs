using System.Globalization;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Metadata;
using Bing.Offices.Npoi;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// 按工作表执行已解析的 NPOI 导入计划。
/// </summary>
internal sealed class NpoiImportSheetExecutor
{
    private readonly NpoiImportRowMaterializer _rowMaterializer;

    /// <summary>创建工作表执行器。</summary>
    internal NpoiImportSheetExecutor(NpoiImportRowMaterializer rowMaterializer)
    {
        _rowMaterializer = rowMaterializer ?? throw new ArgumentNullException(nameof(rowMaterializer));
    }

    /// <summary>执行单个泛型工作表的行导入，并返回成功行索引。</summary>
    internal void Execute<T>(ISheet sheet, ExcelImportExecutionOptions<T> options, ICollection<T> items,
        ExcelImportErrorCollector errors, ExcelImportRuntime runtime, CancellationToken cancellationToken,
        ICollection<int> sourceRows = null) where T : class, new()
    {
        var header = sheet.GetRow(options.HeaderRowIndex)
            ?? throw new NpoiSheetStructureException("导入的模板不正确，未匹配表头。");
        if (header.LastCellNum > options.MaxReadColumns)
            throw new NpoiSheetStructureException($"导入表头超过最大列长度: {options.MaxReadColumns}");

        var columns = CreateColumns<T>(header, options);
        if (runtime.RowLimitReached)
            return;
        IReadOnlyDictionary<(int Row, int Column), IReadOnlyList<PictureInfo>> imageIndex = null;
        HashSet<int> imageRows = null;
        if (columns.Values.Any(NpoiImportRowMaterializer.IsImageColumn))
        {
            try
            {
                imageIndex = NpoiImportRowMaterializer.BuildImageIndex(sheet, runtime.ImageResources,
                    cancellationToken, out imageRows);
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
        var uniqueTracker = new UniqueTracker(duplicateValues, options.MaxTrackedUniqueValues,
            CreateStringComparer(options.UniqueComparison));
        var configuredValidationEnabled = IsConfiguredValidationEnabled(options.ValidationMode);
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
            if (NpoiImportRowMaterializer.IsEmpty(row, options.BodyWhitespace, imageRows, rowIndex))
            {
                if (options.StopAtFirstEmptyRow)
                    break;
                if (options.ReportEmptyRows)
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
            if (configuredValidationEnabled)
                uniqueTracker.BeginRow();
            var workbookValid = NpoiWorkbookValidationPipeline.Validate(row, columns, validationIndex, sheet,
                sheet.SheetName, rowIndex, options.BodyWhitespace, options.ValidateMode,
                options.UnsupportedFeaturePolicy, errors, options.IsDate1904);
            if (!workbookValid)
            {
                if (configuredValidationEnabled)
                    uniqueTracker.RollbackRow();
                if (errors.IsLimitReached)
                {
                    errors.MarkTruncated();
                    break;
                }
                continue;
            }
            if (configuredValidationEnabled && !_rowMaterializer.ValidateRawValues(row, columns, duplicateValues,
                    sheet.SheetName, rowIndex, options.ValidateMode, options.Culture, options.BodyWhitespace, errors,
                    options.IsDate1904))
            {
                uniqueTracker.RollbackRow();
                if (errors.IsLimitReached)
                {
                    errors.MarkTruncated();
                    break;
                }
                continue;
            }
            if (_rowMaterializer.TryCreateItem(row, columns, duplicateValues, uniqueTracker, sheet.SheetName,
                    rowIndex, options.ValidateMode, configuredValidationEnabled, errors, options.Culture,
                    options.BodyWhitespace, options.DynamicTargetGetter, imageIndex, options.IsDate1904,
                    out T item))
            {
                items.Add(item);
                sourceRows?.Add(rowIndex);
                if (configuredValidationEnabled)
                    uniqueTracker.CommitRow();
            }
            else if (configuredValidationEnabled)
                uniqueTracker.RollbackRow();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
        }
    }

    private static bool IsConfiguredValidationEnabled(ExcelImportValidationMode mode) =>
        mode == ExcelImportValidationMode.ConfiguredRules
        || mode == ExcelImportValidationMode.ConfiguredAndWorkbook;

    private static IReadOnlyDictionary<int, ExcelColumnPlan> CreateColumns<T>(IRow header,
        ExcelImportExecutionOptions<T> options) where T : class, new()
    {
        var map = options.MappingPlan;
        if (map == null)
            throw new BingOfficesConfigurationException("工作表导入计划不可用。",
                stage: BingOfficesStage.Plan);
        var dynamicProperties = map.Columns.Where(property => property.IsDynamicColumn).ToList();
        var dynamicPlans = map.DynamicColumns;
        if (dynamicProperties.Count > 1)
            throw new BingOfficesConfigurationException(
                $"导入模板 {typeof(T).FullName} 只能声明一个动态列属性。", stage: BingOfficesStage.Plan);
        var fixedProperties = map.Columns.Where(property => !property.Ignored && !property.IsDynamicColumn).ToList();
        var headerNames = new HashSet<string>(options.HeaderComparison == ExcelNameComparison.Ordinal
            ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase);
        var columns = new Dictionary<int, ExcelColumnPlan>();
        foreach (var headerCell in header.Cells)
        {
            if (options.ReadColumnRange != null && !options.ReadColumnRange.Contains(headerCell.ColumnIndex))
                continue;
            var headerName = NpoiExcelImporter.NormalizeText(NpoiExcelImporter.GetRawStringValue(headerCell),
                options.HeaderWhitespace);
            if (string.IsNullOrWhiteSpace(headerName))
                continue;
            if (!headerNames.Add(headerName))
                throw new NpoiSheetStructureException($"导入的表格存在重复列:{headerName}");
            var property = FindProperty(fixedProperties, headerName, options.HeaderComparison);
            ExcelDynamicColumnDefinition dynamicDefinition = null;
            IExcelDynamicMappingColumn dynamicPlan = null;
            if (property == null && dynamicProperties.Count == 1)
            {
                dynamicPlan = FindDynamicDefinition(headerName, dynamicPlans, options.HeaderComparison);
                dynamicDefinition = dynamicPlan == null ? null : CreateDynamicDefinition(dynamicPlan);
                if (dynamicPlans.Count > 0 && dynamicDefinition == null)
                {
                    if (options.FailOnUnknownDynamicColumns)
                        throw new NpoiSheetStructureException($"导入包含未知动态列: {headerName}");
                    continue;
                }
                property = dynamicProperties[0];
            }
            if (property == null)
                continue;
            var isUnspecifiedDynamicColumn = property.IsDynamicColumn && dynamicDefinition == null;
            var reflectionProperty = typeof(T).GetProperty(property.Name, BindingFlags.Instance | BindingFlags.Public);
            if (reflectionProperty == null)
                throw new BingOfficesConfigurationException($"无法解析映射属性: {property.Name}",
                    stage: BingOfficesStage.Plan);
            if (!property.IsDynamicColumn && !reflectionProperty.CanWrite)
                throw new BingOfficesConfigurationException($"属性不可写入: {property.Name}",
                    stage: BingOfficesStage.Plan);
            var valueConverters = isUnspecifiedDynamicColumn
                ? (IReadOnlyList<Bing.Offices.Conversions.IExcelValueConverter>)Array.Empty<Bing.Offices.Conversions.IExcelValueConverter>()
                : property.IsDynamicColumn ? dynamicPlan.ValueConverters : property.ValueConverters;
            var validationBindings = property.IsDynamicColumn && dynamicPlan != null
                ? dynamicPlan.ValidationBindings : property.ValidationBindings;
            columns[headerCell.ColumnIndex] = new ExcelColumnPlan(headerName, property, property.IsDynamicColumn,
                headerCell.ColumnIndex, dynamicDefinition, null, valueConverters, validationBindings,
                reflectionProperty: reflectionProperty, isUnique: dynamicPlan?.IsUnique,
                uniqueIgnoreEmpty: dynamicPlan?.UniqueIgnoreEmpty ?? true);
        }
        if (options.RequireExpectedHeaders)
        {
            var missing = fixedProperties.Where(property => !columns.Values.Any(column => column.Property == property)
                && (options.ReadColumnRange == null || !header.Cells.Any(cell =>
                    !options.ReadColumnRange.Contains(cell.ColumnIndex)
                    && (string.Equals(NpoiExcelImporter.NormalizeText(NpoiExcelImporter.GetRawStringValue(cell),
                            options.HeaderWhitespace), property.Title, ToStringComparison(options.HeaderComparison))
                        || property.Aliases.Any(alias => string.Equals(
                            NpoiExcelImporter.NormalizeText(NpoiExcelImporter.GetRawStringValue(cell),
                                options.HeaderWhitespace), alias, ToStringComparison(options.HeaderComparison)))))))
                .ToList();
            if (missing.Any())
                throw new NpoiSheetStructureException($"导入的表格不存在列：{string.Join(",", missing.Select(property => property.Title))}");
        }
        return columns;
    }

    private static IExcelDynamicMappingColumn FindDynamicDefinition(string headerName,
        IReadOnlyList<IExcelDynamicMappingColumn> definitions, ExcelNameComparison comparison) =>
        (definitions ?? Array.Empty<IExcelDynamicMappingColumn>()).FirstOrDefault(definition =>
            string.Equals(definition.Title, headerName, ToStringComparison(comparison))
            || (definition.Aliases ?? Array.Empty<string>()).Any(alias =>
                string.Equals(alias, headerName, ToStringComparison(comparison))));

    private static ExcelDynamicColumnDefinition CreateDynamicDefinition(IExcelDynamicMappingColumn column) => new()
    {
        Key = column.Key,
        Title = column.Title,
        Aliases = column.Aliases,
        DataType = ResolveDynamicType(column.DataTypeName),
        Order = column.Order,
        Placement = CreatePlacement(column.PlacementKey),
        PhysicalColumnIndex = column.ColumnIndex,
        NumberFormat = column.NumberFormat,
        ConverterName = column.ConverterName,
        ValidatorName = column.ValidatorName,
        ValidationRuleNames = column.ValidationRuleNames,
        ImageMultiplicity = column.ImageMultiplicity
    };

    private static ExcelColumnPlacement CreatePlacement(string placementKey)
    {
        if (string.IsNullOrWhiteSpace(placementKey))
            return null;
        var separator = placementKey.IndexOfAny(new[] { ':', '-' });
        var key = placementKey.Substring(separator + 1);
        return placementKey.StartsWith("before:", StringComparison.OrdinalIgnoreCase)
            || placementKey.StartsWith("before-", StringComparison.OrdinalIgnoreCase)
            ? ExcelColumnPlacement.Before(key) : ExcelColumnPlacement.After(key);
    }

    private static Type ResolveDynamicType(string name) => (name ?? "string").ToLowerInvariant() switch
    {
        "object" => typeof(object), "string" => typeof(string), "boolean" or "bool" => typeof(bool),
        "byte" => typeof(byte), "int16" => typeof(short), "int32" or "int" => typeof(int),
        "int64" or "long" => typeof(long), "single" or "float" => typeof(float), "double" => typeof(double),
        "decimal" => typeof(decimal), "datetime" => typeof(DateTime), "datetimeoffset" => typeof(DateTimeOffset),
        "guid" => typeof(Guid), "bytes" => typeof(byte[]),
        _ => throw new NpoiSheetStructureException($"动态列数据类型不在允许列表中: {name}")
    };

    private static IExcelMappingColumn FindProperty(IEnumerable<IExcelMappingColumn> properties, string headerName,
        ExcelNameComparison comparison)
    {
        var stringComparison = ToStringComparison(comparison);
        return properties.FirstOrDefault(property => string.Equals(property.Title, headerName, stringComparison)
            || property.Aliases.Any(alias => string.Equals(alias, headerName, stringComparison))
            || string.Equals(property.Name, headerName, stringComparison));
    }

    private static StringComparison ToStringComparison(ExcelNameComparison comparison) =>
        comparison == ExcelNameComparison.Ordinal ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

    private static StringComparer CreateStringComparer(StringComparison comparison) => comparison switch
    {
        StringComparison.Ordinal => StringComparer.Ordinal,
        StringComparison.OrdinalIgnoreCase => StringComparer.OrdinalIgnoreCase,
        StringComparison.InvariantCulture => StringComparer.InvariantCulture,
        StringComparison.InvariantCultureIgnoreCase => StringComparer.InvariantCultureIgnoreCase,
        StringComparison.CurrentCulture => StringComparer.CurrentCulture,
        StringComparison.CurrentCultureIgnoreCase => StringComparer.CurrentCultureIgnoreCase,
        _ => throw new ArgumentOutOfRangeException(nameof(comparison))
    };

}

internal sealed class NpoiSheetStructureException : InvalidOperationException
{
    internal NpoiSheetStructureException(string message) : base(message) { }
}
