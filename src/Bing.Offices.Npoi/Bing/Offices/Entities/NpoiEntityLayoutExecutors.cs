using System.Collections;
using System.Collections.Concurrent;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Text.RegularExpressions;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Providers;
using Bing.Offices.Styles;
using Bing.Offices.Validations;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Bing.Offices;

/// <summary>
/// NPOI 实体布局执行共用的 Sheet、合并区域和转换边界。
/// </summary>
internal static class NpoiEntityLayoutSupport
{
    /// <summary>
    /// 识别模板命名锚点引用的单元格或单格区域。
    /// </summary>
    private static readonly Regex NamedAnchorFormula = new Regex(
        @"^\s*=?\s*(?:'(?<quoted>(?:[^']|'')+)'|(?<plain>[^'!]+))?!\$?(?<startColumn>[A-Za-z]{1,3})\$?(?<startRow>[1-9][0-9]*)(?::\$?(?<endColumn>[A-Za-z]{1,3})\$?(?<endRow>[1-9][0-9]*))?\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// 获取指定工作表，并在非模板模式下按需创建工作表。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="name">工作表名称。</param>
    /// <param name="template">是否要求工作表已存在于模板中。</param>
    /// <returns>已解析或新建的工作表。</returns>
    internal static ISheet ResolveSheet(IWorkbook workbook, string name, bool template)
    {
        var sheet = workbook.GetSheet(name);
        if (sheet == null && template)
            throw new BingOfficesConfigurationException($"模板缺少请求的 Sheet: {name}", stage: BingOfficesStage.Plan);
        return sheet ?? workbook.CreateSheet(name);
    }

    /// <summary>
    /// 解析模板命名锚点并返回其所在工作表和左上角坐标。
    /// </summary>
    /// <param name="workbook">包含命名范围的工作簿。</param>
    /// <param name="sheetName">布局声明的工作表名称。</param>
    /// <param name="anchorName">命名范围名称。</param>
    /// <returns>命名锚点所在工作表和坐标。</returns>
    internal static (ISheet Sheet, ExcelEntityCellReference Reference) ResolveNamedAnchor(
        IWorkbook workbook, string sheetName, string anchorName)
    {
        var sheet = workbook.GetSheet(sheetName);
        if (sheet == null)
            throw new BingOfficesConfigurationException($"工作簿缺少命名锚点所属的 Sheet: {sheetName}",
                stage: BingOfficesStage.Plan);
        var sheetIndex = workbook.GetSheetIndex(sheetName);
        var candidates = workbook.GetAllNames()
            .Where(name => string.Equals(name.NameName, anchorName, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        var scoped = candidates.Where(name => name.SheetIndex == sheetIndex).ToArray();
        var global = candidates.Where(name => name.SheetIndex < 0).ToArray();
        var matches = scoped.Length > 0 ? scoped : global;
        if (matches.Length == 0)
            throw new BingOfficesConfigurationException(
                $"模板缺少命名锚点: {sheetName}!{anchorName}", stage: BingOfficesStage.Plan);
        if (matches.Length > 1)
            throw new BingOfficesConfigurationException(
                $"模板命名锚点不唯一: {sheetName}!{anchorName}", stage: BingOfficesStage.Plan);

        var match = NamedAnchorFormula.Match(matches[0].RefersToFormula ?? string.Empty);
        if (!match.Success)
            throw new BingOfficesConfigurationException(
                $"模板命名锚点必须引用单个有界单元格: {sheetName}!{anchorName}",
                stage: BingOfficesStage.Plan);
        var targetSheet = match.Groups["quoted"].Success
            ? match.Groups["quoted"].Value.Replace("''", "'")
            : match.Groups["plain"].Success ? match.Groups["plain"].Value : sheetName;
        if (!match.Groups["quoted"].Success && !match.Groups["plain"].Success
            && matches[0].SheetIndex != sheetIndex)
            throw new BingOfficesConfigurationException(
                $"模板命名锚点缺少 Sheet 引用: {sheetName}!{anchorName}", stage: BingOfficesStage.Plan);
        if (!string.Equals(targetSheet, sheetName, StringComparison.OrdinalIgnoreCase))
            throw new BingOfficesConfigurationException(
                $"模板命名锚点引用了其他 Sheet: {sheetName}!{anchorName}", stage: BingOfficesStage.Plan);
        if (match.Groups["endColumn"].Success
            && (!string.Equals(match.Groups["startColumn"].Value, match.Groups["endColumn"].Value,
                StringComparison.OrdinalIgnoreCase)
                || !string.Equals(match.Groups["startRow"].Value, match.Groups["endRow"].Value,
                    StringComparison.Ordinal)))
            throw new BingOfficesConfigurationException(
                $"模板命名锚点必须引用单个单元格: {sheetName}!{anchorName}",
                stage: BingOfficesStage.Plan);
        try
        {
            return (sheet, ExcelEntityCellReference.Parse(match.Groups["startColumn"].Value
                + match.Groups["startRow"].Value));
        }
        catch (ArgumentException exception)
        {
            throw new BingOfficesConfigurationException(
                $"模板命名锚点地址超出工作表边界: {sheetName}!{anchorName}", exception,
                BingOfficesStage.Plan);
        }
    }

    /// <summary>
    /// 检查模板解析后的列表区域是否覆盖固定单元格。
    /// </summary>
    internal static void ValidateCellRegionConflict<TEntity>(IWorkbook workbook,
        ExcelEntityLayout<TEntity> layout, ExcelEntityListRegion<TEntity> region)
        where TEntity : class, new()
    {
        foreach (var cell in layout.Cells.Where(item => string.Equals(item.SheetName,
                     region.SheetName, StringComparison.OrdinalIgnoreCase)))
        {
            var reference = cell.AnchorName == null ? cell.Reference :
                ResolveNamedAnchor(workbook, cell.SheetName, cell.AnchorName).Reference;
            if (reference.Row >= region.Start.Row && reference.Column >= region.Start.Column
                && (!region.End.HasValue || (reference.Row <= region.End.Value.Row
                    && reference.Column <= region.End.Value.Column)))
                throw new BingOfficesConfigurationException(
                    $"固定单元格与列表区域重叠: {cell.SheetName}!{reference.Address}",
                    stage: BingOfficesStage.Plan);
        }
    }

    /// <summary>
    /// 获取指定位置的单元格，并根据要求创建缺失的行和单元格。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="reference">零基单元格引用。</param>
    /// <param name="create">是否创建缺失的行和单元格。</param>
    /// <returns>找到或创建的单元格；未找到且不允许创建时返回 <see langword="null"/>。</returns>
    internal static ICell GetCell(ISheet sheet, ExcelEntityCellReference reference, bool create)
    {
        var row = sheet.GetRow(reference.Row);
        if (row == null && !create)
            return null;
        row ??= sheet.CreateRow(reference.Row);
        var cell = row.GetCell(reference.Column);
        return cell ?? (create ? row.CreateCell(reference.Column) : null);
    }

    /// <summary>
    /// 将合并区域内的单元格引用解析为该区域的左上角锚点。
    /// </summary>
    /// <param name="sheet">包含合并区域的工作表。</param>
    /// <param name="reference">待解析的单元格引用。</param>
    /// <returns>合并区域锚点；引用不在合并区域内时返回原引用。</returns>
    internal static ExcelEntityCellReference ResolveAnchor(ISheet sheet, ExcelEntityCellReference reference)
    {
        for (var index = 0; index < sheet.NumMergedRegions; index++)
        {
            var range = sheet.GetMergedRegion(index);
            if (reference.Row < range.FirstRow || reference.Row > range.LastRow
                || reference.Column < range.FirstColumn || reference.Column > range.LastColumn)
                continue;
            return ExcelEntityCellReference.Parse(ToAddress(range.FirstRow, range.FirstColumn));
        }
        return reference;
    }

    /// <summary>
    /// 校验布局声明的合并区域，并按执行模式补建或要求模板已存在。
    /// </summary>
    /// <param name="sheet">待校验的工作表。</param>
    /// <param name="requested">布局声明的合并区域。</param>
    /// <param name="addMissing">是否添加工作表中缺失的合并区域。</param>
    /// <param name="requireExisting">是否要求声明的区域已存在于工作表中。</param>
    internal static void PreflightMerges(ISheet sheet, IEnumerable<ExcelEntityMergeRegion> requested,
        bool addMissing, bool requireExisting)
    {
        var pending = requested?.Where(item => string.Equals(item.SheetName, sheet.SheetName,
            StringComparison.OrdinalIgnoreCase)).Select(item => item.Range).ToArray()
            ?? Array.Empty<ExcelEntityCellRange>();
        foreach (var range in pending)
        {
            var exact = FindExistingMerge(sheet, range);
            if (exact != null)
                continue;
            for (var index = 0; index < sheet.NumMergedRegions; index++)
            {
                var existing = sheet.GetMergedRegion(index);
                if (Overlaps(existing, range) && !IsExact(existing, range))
                    throw new BingOfficesConfigurationException(
                        $"工作表 {sheet.SheetName} 的合并区域与模板已有区域冲突: {range.Address}",
                        stage: BingOfficesStage.Plan);
            }
            if (requireExisting)
                throw new BingOfficesConfigurationException(
                    $"模板缺少声明的合并区域: {sheet.SheetName}!{range.Address}", stage: BingOfficesStage.Plan);
        }
        for (var left = 0; left < pending.Length; left++)
        {
            for (var right = left + 1; right < pending.Length; right++)
            {
                if (Overlaps(pending[left], pending[right]) && !pending[left].Equals(pending[right]))
                    throw new BingOfficesConfigurationException(
                        $"实体布局包含重叠合并区域: {sheet.SheetName}!{pending[left].Address}",
                        stage: BingOfficesStage.Plan);
            }
        }
        if (!addMissing)
            return;
        foreach (var range in pending)
        {
            if (FindExistingMerge(sheet, range) == null)
                sheet.AddMergedRegion(new CellRangeAddress(range.First.Row, range.Last.Row,
                    range.First.Column, range.Last.Column));
        }
    }

    /// <summary>
    /// 使用已注册的值转换器将实体属性值转换为单元格写入值。
    /// </summary>
    /// <param name="value">待转换的属性值。</param>
    /// <param name="propertyName">属性名称。</param>
    /// <param name="propertyType">属性类型。</param>
    /// <param name="converterName">可选的命名转换器名称。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="reference">目标单元格引用。</param>
    /// <param name="converters">候选值转换器。</param>
    /// <returns>转换后的单元格写入值；没有转换器处理时返回原值。</returns>
    internal static object ConvertTo(object value, string propertyName, Type propertyType,
        string converterName, string sheetName, ExcelEntityCellReference reference,
        IReadOnlyList<IExcelValueConverter> converters)
    {
        var context = new ExcelConversionContext(value, propertyName, propertyType, sheetName,
            reference.Row + 1, reference.Column + 1, CultureInfo.InvariantCulture);
        foreach (var converter in ResolveConverters(converterName, propertyType, converters))
        {
            try
            {
                if (converter.TryConvertTo(context, out var converted))
                    return converted;
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
                throw new BingOfficesExportException("Excel 实体值转换器执行失败。", exception, "NPOI",
                    BingOfficesStage.Validate, sheetName, reference.Row + 1, reference.Column + 1,
                    propertyName, BingOfficesErrorCode.UserExtensionFailed);
            }
        }
        return value;
    }

    /// <summary>
    /// 使用已注册的值转换器或内置规则将单元格值转换为实体属性值。
    /// </summary>
    /// <param name="cellValue">源单元格值。</param>
    /// <param name="propertyName">目标属性名称。</param>
    /// <param name="propertyType">目标属性类型。</param>
    /// <param name="converterName">可选的命名转换器名称。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="reference">源单元格引用。</param>
    /// <param name="converters">候选值转换器。</param>
    /// <returns>适用于目标属性的值。</returns>
    internal static object ConvertFrom(ExcelCellValue cellValue, string propertyName, Type propertyType,
        string converterName, string sheetName, ExcelEntityCellReference reference,
        IReadOnlyList<IExcelValueConverter> converters)
    {
        var text = NpoiExcelImporter.NormalizeText(cellValue?.Text, ExcelWhitespacePolicy.Trim);
        var context = new ExcelConversionContext(text, propertyName, propertyType, sheetName,
            reference.Row + 1, reference.Column + 1, CultureInfo.InvariantCulture, cellValue);
        foreach (var converter in ResolveConverters(converterName, propertyType, converters))
        {
            try
            {
                if (converter.TryConvertFrom(context, out var converted))
                    return converted;
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
                throw new BingOfficesImportException("Excel 实体值转换器执行失败。", exception, "NPOI",
                    BingOfficesStage.Validate, sheetName, reference.Row + 1, reference.Column + 1,
                    propertyName, BingOfficesErrorCode.UserExtensionFailed);
            }
        }
        if (cellValue != null && (Nullable.GetUnderlyingType(propertyType) ?? propertyType) is var target
            && (target == typeof(DateTime) || target == typeof(DateTimeOffset))
            && new DateTimeExcelValidationRule().TryParseValue(cellValue, text, propertyType,
                CultureInfo.InvariantCulture, null, out var dateValue))
            return dateValue;
        if (string.IsNullOrWhiteSpace(text))
        {
            if (!propertyType.IsValueType || Nullable.GetUnderlyingType(propertyType) != null)
                return null;
            throw new InvalidCastException($"值转换失败。输入值为空，目标类型为: {propertyType.FullName}");
        }
        var targetType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
        if (targetType == typeof(string))
            return text;
        if (targetType.IsEnum)
            return Enum.Parse(targetType, text, true);
        if (targetType == typeof(Guid))
            return Guid.Parse(text);
        if (targetType == typeof(Version))
            return new Version(text);
        return Convert.ChangeType(cellValue?.Value ?? text, targetType, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 按既有映射计划创建固定单元格列计划。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="mappingPlanFactory">映射计划工厂。</param>
    /// <param name="binding">固定单元格绑定。</param>
    /// <param name="direction">映射方向。</param>
    /// <returns>与列表列相同的固定单元格列计划。</returns>
    internal static ExcelColumnPlan CreateCellPlan<TEntity>(IExcelMappingPlanFactory mappingPlanFactory,
        ExcelEntityCellBinding<TEntity> binding, MappingDirection direction) where TEntity : class, new()
    {
        if (mappingPlanFactory == null)
            throw new ArgumentNullException(nameof(mappingPlanFactory));
        if (binding == null)
            throw new ArgumentNullException(nameof(binding));

        var document = binding.MappingDocument ?? new ExcelMappingDocument
        {
            UseConventionFallback = true
        };
        var requestConfiguration = binding.MappingConfiguration;
        if (!string.IsNullOrWhiteSpace(binding.ConverterName))
        {
            requestConfiguration ??= new ExcelMappingConfiguration();
            var column = requestConfiguration.Columns.FirstOrDefault(item =>
                string.Equals(item.PropertyName, binding.PropertyName, StringComparison.OrdinalIgnoreCase));
            if (column == null)
            {
                column = new ExcelColumnConfiguration { PropertyName = binding.PropertyName };
                requestConfiguration.Columns.Add(column);
            }
            column.ConverterName = binding.ConverterName;
            column.ClearConverterName = false;
        }

        var mapping = mappingPlanFactory.Create<TEntity>(document, requestConfiguration, direction);
        var property = mapping.Columns.FirstOrDefault(item =>
            string.Equals(item.Name, binding.PropertyName, StringComparison.OrdinalIgnoreCase));
        if (property == null || property.Ignored || property.IsDynamicColumn)
            throw new BingOfficesConfigurationException(
                $"固定单元格属性未生成可用映射列: {binding.PropertyName}", stage: BingOfficesStage.Plan);
        var reflectionProperty = typeof(TEntity).GetProperty(binding.PropertyName,
            BindingFlags.Instance | BindingFlags.Public);
        if (reflectionProperty == null)
            throw new BingOfficesConfigurationException(
                $"无法解析实体属性: {binding.PropertyName}", stage: BingOfficesStage.Plan);
        return new ExcelColumnPlan(property.Title, property, false, binding.Reference.Column,
            null, binding.PropertyName, property.ValueConverters, property.ValidationBindings,
            reflectionProperty: reflectionProperty);
    }

    /// <summary>
    /// 执行固定单元格导入的原始值校验。
    /// </summary>
    /// <param name="column">固定单元格列计划。</param>
    /// <param name="cellValue">源单元格值。</param>
    /// <param name="value">规范化后的文本值。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="reference">单元格坐标。</param>
    /// <param name="errors">导入错误收集器。</param>
    /// <returns>所有校验通过时返回 <see langword="true"/>。</returns>
    internal static bool ValidateCellImportRaw(ExcelColumnPlan column, ExcelCellValue cellValue, string value,
        string sheetName, ExcelEntityCellReference reference,
        ExcelImportErrorCollector errors)
    {
        var context = new ExcelValidationContext(value, sheetName, reference.Row + 1, reference.Column + 1,
            column.Property.Name, null, column.ValueType, cellValue, CultureInfo.InvariantCulture);
        return ValidateCellBindings(column.ValidationBindings.Where(item => item.IsRaw), context, column,
            sheetName, reference, cellValue, errors);
    }

    /// <summary>
    /// 执行固定单元格导入的转换后值校验。
    /// </summary>
    /// <param name="column">固定单元格列计划。</param>
    /// <param name="cellValue">源单元格值。</param>
    /// <param name="value">规范化后的文本值。</param>
    /// <param name="convertedValue">转换后的属性值。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="reference">单元格坐标。</param>
    /// <param name="errors">导入错误收集器。</param>
    /// <returns>所有校验通过时返回 <see langword="true"/>。</returns>
    internal static bool ValidateCellImportConverted(ExcelColumnPlan column, ExcelCellValue cellValue, string value,
        object convertedValue, string sheetName, ExcelEntityCellReference reference,
        ExcelImportErrorCollector errors)
    {
        var context = new ExcelValidationContext(value, sheetName, reference.Row + 1, reference.Column + 1,
            column.Property.Name, convertedValue, column.ValueType, cellValue, CultureInfo.InvariantCulture);
        return ValidateCellBindings(column.ValidationBindings.Where(item => !item.IsRaw
                && item.Kind != ExcelValidationBindingKind.Unique), context, column, sheetName, reference,
            cellValue, errors);
    }

    /// <summary>
    /// 按映射计划绑定执行一阶段校验，统一保留导入错误坐标和错误码语义。
    /// </summary>
    /// <param name="bindings">当前阶段的校验绑定。</param>
    /// <param name="context">当前阶段的校验上下文。</param>
    /// <param name="column">当前列计划。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="reference">单元格坐标。</param>
    /// <param name="cellValue">源单元格值。</param>
    /// <param name="errors">导入错误收集器。</param>
    /// <returns>所有绑定通过时返回 <see langword="true"/>。</returns>
    private static bool ValidateCellBindings(IEnumerable<IExcelValidationBinding> bindings,
        ExcelValidationContext context, ExcelColumnPlan column, string sheetName,
        ExcelEntityCellReference reference, ExcelCellValue cellValue, ExcelImportErrorCollector errors)
    {
        foreach (var binding in bindings)
        {
            if (!TryValidateCellBinding(binding, context, column, sheetName, reference, cellValue, errors))
                return false;
        }
        return true;
    }

    /// <summary>
    /// 执行固定单元格导出校验，并在失败时保留完整坐标上下文。
    /// </summary>
    /// <param name="column">固定单元格列计划。</param>
    /// <param name="rawValue">实体属性原始值。</param>
    /// <param name="convertedValue">转换后待写入的值。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="reference">单元格坐标。</param>
    internal static void ValidateCellExport(ExcelColumnPlan column, object rawValue, object convertedValue,
        string sheetName, ExcelEntityCellReference reference)
    {
        var text = convertedValue == null ? string.Empty :
            Convert.ToString(convertedValue, CultureInfo.InvariantCulture);
        var context = new ExcelValidationContext(text, sheetName, reference.Row + 1, reference.Column + 1,
            column.Property.Name, rawValue, column.ValueType, null, CultureInfo.InvariantCulture);
        foreach (var binding in column.ValidationBindings.Where(item => item.Kind != ExcelValidationBindingKind.Unique))
        {
            bool valid;
            try
            {
                valid = binding.Validate(context);
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
                throw new BingOfficesExportException("Excel 实体导出校验器执行失败。", exception, "NPOI",
                    BingOfficesStage.Validate, sheetName, reference.Row + 1, reference.Column + 1,
                    column.Property.Name, BingOfficesErrorCode.UserExtensionFailed);
            }
            if (!valid)
                throw new BingOfficesExportException(binding.ErrorMessage ?? "Excel 实体属性未通过校验。",
                    provider: "NPOI", stage: BingOfficesStage.Validate, sheetName: sheetName,
                    rowIndex: reference.Row + 1, columnIndex: reference.Column + 1,
                    propertyName: column.Property.Name);
        }
    }

    /// <summary>
    /// 执行一个导入校验绑定并将失败记录为完整单元格错误。
    /// </summary>
    /// <param name="binding">待执行的校验绑定。</param>
    /// <param name="context">校验上下文。</param>
    /// <param name="column">当前列计划。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="reference">单元格坐标。</param>
    /// <param name="cellValue">源单元格值。</param>
    /// <param name="errors">导入错误收集器。</param>
    /// <returns>校验通过时返回 <see langword="true"/>。</returns>
    private static bool TryValidateCellBinding(IExcelValidationBinding binding, ExcelValidationContext context,
        ExcelColumnPlan column, string sheetName, ExcelEntityCellReference reference, ExcelCellValue cellValue,
        ExcelImportErrorCollector errors)
    {
        bool valid;
        try
        {
            valid = binding.Validate(context);
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (binding.Kind == ExcelValidationBindingKind.Custom
            && exception is not OutOfMemoryException && exception is not StackOverflowException)
        {
            throw new BingOfficesImportException("Excel 实体自定义校验器执行失败。", exception, "NPOI",
                BingOfficesStage.Validate, sheetName, reference.Row + 1, reference.Column + 1,
                column.Property.Name, BingOfficesErrorCode.UserExtensionFailed);
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, exception.Message, sheetName,
                reference.Row + 1, reference.Column + 1, column.Property.Name, column.Property.Name,
                column.HeaderName, cellValue?.Value ?? cellValue?.Text));
            return false;
        }
        if (valid)
            return true;
        errors.Add(new ExcelImportError(GetValidationErrorCode(binding), binding.ErrorMessage, sheetName,
            reference.Row + 1, reference.Column + 1, column.Property.Name, column.Property.Name,
            column.HeaderName, cellValue?.Value ?? cellValue?.Text));
        return false;
    }

    /// <summary>
    /// 将校验绑定映射到公共导入错误码。
    /// </summary>
    /// <param name="binding">产生校验结果的规则绑定。</param>
    /// <returns>与绑定类型对应的导入错误码。</returns>
    private static ExcelImportErrorCode GetValidationErrorCode(IExcelValidationBinding binding) =>
        binding.Kind == ExcelValidationBindingKind.MaxLength
            ? ExcelImportErrorCode.MaxLength
            : binding.Kind == ExcelValidationBindingKind.MaxValue
                ? ExcelImportErrorCode.MaxValue
                : ExcelImportErrorCode.Validation;

    /// <summary>
    /// 按固定单元格声明的名称解析转换器；未指定名称时保留按类型探测顺序。
    /// </summary>
    /// <param name="converterName">可选的命名转换器名称。</param>
    /// <param name="propertyType">实体属性类型。</param>
    /// <param name="converters">当前 Provider 注册的转换器集合。</param>
    /// <returns>可用于当前属性的转换器集合。</returns>
    private static IReadOnlyList<IExcelValueConverter> ResolveConverters(string converterName,
        Type propertyType, IReadOnlyList<IExcelValueConverter> converters)
    {
        var items = converters ?? Array.Empty<IExcelValueConverter>();
        if (string.IsNullOrWhiteSpace(converterName))
            return items.Where(converter => converter.CanConvert(propertyType)).ToArray();
        var named = items.OfType<INamedExcelValueConverter>().Where(converter =>
            string.Equals(converter.Name, converterName, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (named.Length != 1)
            throw new InvalidOperationException($"未找到唯一命名值转换器: {converterName}");
        if (!named[0].CanConvert(propertyType))
            throw new InvalidOperationException($"值转换器 {converterName} 不支持属性类型: {propertyType.FullName}");
        return named;
    }

    /// <summary>
    /// 查找与布局范围边界完全一致的已有合并区域。
    /// </summary>
    /// <param name="sheet">待查询的工作表。</param>
    /// <param name="range">布局声明的单元格范围。</param>
    /// <returns>匹配的合并区域；不存在时返回 <see langword="null"/>。</returns>
    private static CellRangeAddress FindExistingMerge(ISheet sheet, ExcelEntityCellRange range)
    {
        for (var index = 0; index < sheet.NumMergedRegions; index++)
        {
            var existing = sheet.GetMergedRegion(index);
            if (IsExact(existing, range))
                return existing;
        }
        return null;
    }

    /// <summary>
    /// 判断 NPOI 合并区域与布局范围的边界是否完全相同。
    /// </summary>
    /// <param name="actual">NPOI 合并区域。</param>
    /// <param name="expected">布局期望范围。</param>
    /// <returns>边界完全相同时返回 <see langword="true"/>。</returns>
    private static bool IsExact(CellRangeAddress actual, ExcelEntityCellRange expected) =>
        actual.FirstRow == expected.First.Row && actual.LastRow == expected.Last.Row
        && actual.FirstColumn == expected.First.Column && actual.LastColumn == expected.Last.Column;

    /// <summary>
    /// 判断 NPOI 合并区域是否与布局范围相交。
    /// </summary>
    /// <param name="actual">NPOI 合并区域。</param>
    /// <param name="expected">布局范围。</param>
    /// <returns>两个范围相交时返回 <see langword="true"/>。</returns>
    private static bool Overlaps(CellRangeAddress actual, ExcelEntityCellRange expected) =>
        actual.FirstRow <= expected.Last.Row && expected.First.Row <= actual.LastRow
        && actual.FirstColumn <= expected.Last.Column && expected.First.Column <= actual.LastColumn;

    /// <summary>
    /// 判断两个布局单元格范围是否相交。
    /// </summary>
    /// <param name="left">左侧范围。</param>
    /// <param name="right">右侧范围。</param>
    /// <returns>两个范围相交时返回 <see langword="true"/>。</returns>
    private static bool Overlaps(ExcelEntityCellRange left, ExcelEntityCellRange right) =>
        left.First.Row <= right.Last.Row && right.First.Row <= left.Last.Row
        && left.First.Column <= right.Last.Column && right.First.Column <= left.Last.Column;

    /// <summary>
    /// 将零基行列索引转换为 A1 单元格地址。
    /// </summary>
    /// <param name="row">零基行索引。</param>
    /// <param name="column">零基列索引。</param>
    /// <returns>A1 格式的单元格地址。</returns>
    private static string ToAddress(int row, int column)
    {
        var value = column + 1;
        var letters = string.Empty;
        while (value > 0)
        {
            var remainder = (value - 1) % 26;
            letters = (char)('A' + remainder) + letters;
            value = (value - 1) / 26;
        }
        return letters + (row + 1).ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 判断单元格是否位于指定列表区域的矩形边界内。
    /// </summary>
    /// <param name="start">区域起始单元格。</param>
    /// <param name="end">可选的区域结束单元格。</param>
    /// <param name="cell">待检查的单元格。</param>
    /// <returns>单元格位于区域内时返回 <see langword="true"/>。</returns>
    internal static bool Overlaps(ExcelEntityCellReference start, ExcelEntityCellReference? end,
        ExcelEntityCellReference cell) => cell.Row >= start.Row && cell.Column >= start.Column
        && (!end.HasValue || (cell.Row <= end.Value.Row && cell.Column <= end.Value.Column));

    /// <summary>
    /// 校验列表区域实际写入或读取的矩形范围是否位于声明边界内。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="region">列表区域定义。</param>
    /// <param name="columnCount">映射计划的实际列数。</param>
    /// <param name="itemCount">区域中将处理的列表项数。</param>
    /// <param name="includeHeader">是否包含表头行。</param>
    internal static void ValidateListRegionBounds<TEntity>(ExcelEntityListRegion<TEntity> region,
        int columnCount, int itemCount, bool includeHeader) where TEntity : class, new()
    {
        if (!region.End.HasValue)
            return;
        var end = region.End.Value;
        var availableColumnCount = end.Column - region.Start.Column + 1;
        var lastRow = region.Start.Row + itemCount + (includeHeader ? 0 : -1);
        if (columnCount > availableColumnCount || (itemCount > 0 && lastRow > end.Row))
            throw new BingOfficesConfigurationException($"列表区域超出声明边界: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
    }

    /// <summary>
    /// 按列表明细行数写入水平分页符。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="itemCount">已写入的明细行数。</param>
    internal static void ApplyPageBreaks<TEntity>(ISheet sheet, ExcelEntityListRegion<TEntity> region,
        int itemCount) where TEntity : class, new()
    {
        if (sheet == null || region.PageBreakRows is not int rowsPerPage || itemCount <= rowsPerPage)
            return;
        var firstDataRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        for (var offset = rowsPerPage; offset < itemCount; offset += rowsPerPage)
        {
            var breakRow = firstDataRow + offset - 1;
            if (!sheet.IsRowBroken(breakRow))
                sheet.SetRowBreak(breakRow);
        }
    }

    /// <summary>
    /// 设置分页小计完成后的水平分页符。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="breakRow">分页符所在的零基行号。</param>
    internal static void ApplyPageBreak(ISheet sheet, int breakRow)
    {
        if (sheet == null || breakRow < 0)
            return;
        if (!sheet.IsRowBroken(breakRow))
            sheet.SetRowBreak(breakRow);
    }

    /// <summary>
    /// 校验分页小计和最终尾部的物理及声明边界。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="workbook">待写入的工作簿。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="columnCount">映射计划的实际列数。</param>
    /// <param name="itemCount">明细数量。</param>
    /// <returns>最终明细结束后的下一行零基索引。</returns>
    internal static int ValidatePagedFooterBounds<TEntity>(IWorkbook workbook,
        ExcelEntityListRegion<TEntity> region, int columnCount, int itemCount) where TEntity : class, new()
    {
        if (region.PageSubtotal == null || region.PageBreakRows is not int rowsPerPage)
            return region.Start.Row + (region.IncludeHeader ? 1 : 0) + itemCount;
        ValidateListRegionBounds(region, columnCount, itemCount, region.IncludeHeader);
        var row = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var written = 0;
        while (written < itemCount)
        {
            var count = Math.Min(rowsPerPage, itemCount - written);
            row += count;
            // 分页尾部和签字区也占用工作表行，末页明细必须按推移后的范围预检。
            ValidateListRegionBounds(region, columnCount,
                row - region.Start.Row - (region.IncludeHeader ? 1 : 0), region.IncludeHeader);
            if (row - 1 > workbook.SpreadsheetVersion.LastRowIndex || (long)region.Start.Column + columnCount - 1 > workbook.SpreadsheetVersion.LastColumnIndex)
                throw new BingOfficesConfigurationException(
                    $"分页列表超出工作表物理边界: {region.SheetName}!{region.Start.Address}",
                    stage: BingOfficesStage.Plan);
            written += count;
            if (written >= itemCount)
                break;
            ValidateFooterBounds(workbook, region, region.PageSubtotal, row);
            row += region.PageSubtotal.GapRows + GetFooterHeight(region.PageSubtotal);
        }
        ValidateFooterBounds(workbook, region, region.Footer, row);
        return row;
    }

    /// <summary>
    /// 校验列表尾部相对于明细区域的边界。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="workbook">待写入的工作簿。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="footer">列表尾部定义。</param>
    /// <param name="dataEndRow">明细结束后的零基行索引。</param>
    internal static void ValidateFooterBounds<TEntity>(IWorkbook workbook,
        ExcelEntityListRegion<TEntity> region, IExcelEntityListFooter footer, int dataEndRow)
        where TEntity : class, new()
    {
        if (footer == null)
            return;
        var footerRow = (long)dataEndRow + footer.GapRows;
        var lastRow = footerRow;
        var lastColumn = (long)region.Start.Column;
        foreach (var cell in footer.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
        {
            lastRow = Math.Max(lastRow, footerRow + cell.Reference.Row);
            lastColumn = Math.Max(lastColumn, (long)region.Start.Column + cell.Reference.Column);
        }
        foreach (var merge in footer.Merges ?? Array.Empty<ExcelEntityCellRange>())
        {
            lastRow = Math.Max(lastRow, footerRow + merge.Last.Row);
            lastColumn = Math.Max(lastColumn, (long)region.Start.Column + merge.Last.Column);
        }
        if (region.End.HasValue && (lastRow > region.End.Value.Row || lastColumn > region.End.Value.Column))
            throw new BingOfficesConfigurationException(
                $"列表尾部超出声明边界: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        var version = workbook?.SpreadsheetVersion ?? NPOI.SS.SpreadsheetVersion.EXCEL2007;
        if (lastRow > version.LastRowIndex || lastColumn > version.LastColumnIndex)
            throw new BingOfficesConfigurationException(
                $"列表尾部超出 {version.DefaultExtension.ToUpperInvariant()} 工作表物理边界: "
                + $"{region.SheetName}!{region.Start.Address}", stage: BingOfficesStage.Plan);
    }

    /// <summary>
    /// 预检并添加列表尾部的绝对合并区域。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="requested">待添加的合并区域。</param>
    internal static void PreflightFooterMerges(ISheet sheet, IEnumerable<CellRangeAddress> requested)
    {
        foreach (var range in requested ?? Array.Empty<CellRangeAddress>())
        {
            var exact = false;
            for (var index = 0; index < sheet.NumMergedRegions; index++)
            {
                var existing = sheet.GetMergedRegion(index);
                if (IsExact(existing, range))
                {
                    exact = true;
                    break;
                }
                if (existing.FirstRow <= range.LastRow && range.FirstRow <= existing.LastRow
                    && existing.FirstColumn <= range.LastColumn && range.FirstColumn <= existing.LastColumn)
                    throw new BingOfficesConfigurationException(
                        $"工作表 {sheet.SheetName} 的尾部合并区域与已有区域冲突: {range.FormatAsString()}",
                        stage: BingOfficesStage.Plan);
            }
            if (!exact)
                sheet.AddMergedRegion(range);
        }
    }

    /// <summary>
    /// 预检指定尾部在运行时行位置的相对合并区域。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="footer">尾部定义。</param>
    /// <param name="dataEndRow">尾部间隔起始前的零基行索引。</param>
    /// <param name="startColumn">列表区域起始零基列索引。</param>
    internal static void PreflightFooterMerges(ISheet sheet, IExcelEntityListFooter footer,
        int dataEndRow, int startColumn)
    {
        if (footer == null)
            return;
        var footerRow = dataEndRow + footer.GapRows;
        var ranges = (footer.Merges ?? Array.Empty<ExcelEntityCellRange>()).Select(merge =>
            new CellRangeAddress(footerRow + merge.First.Row, footerRow + merge.Last.Row,
                startColumn + merge.First.Column, startColumn + merge.Last.Column));
        PreflightFooterMerges(sheet, ranges);
    }

    /// <summary>
    /// 查找列表尾部的唯一结束标记行。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="sheet">源工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <returns>标记所在的零基行索引；未配置尾部时返回 <see langword="null" />。</returns>
    internal static int? FindFooterMarkerRow<TEntity>(ISheet sheet, ExcelEntityListRegion<TEntity> region)
        where TEntity : class, new()
    {
        var footer = region.Footer;
        if (footer == null)
            return null;
        var firstRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var lastRow = region.End?.Row ?? sheet.LastRowNum;
        var matches = new List<int>();
        for (var rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
        {
            var cell = sheet.GetRow(rowIndex)?.GetCell(region.Start.Column);
            var value = cell == null ? null : NpoiExcelImporter.GetRawStringValue(cell);
            if (string.Equals(NpoiExcelImporter.NormalizeText(value, ExcelWhitespacePolicy.Trim),
                    footer.MarkerText, StringComparison.Ordinal))
                matches.Add(rowIndex);
        }
        if (matches.Count == 0)
            throw new BingOfficesConfigurationException(
                $"列表区域缺少尾部标记: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        if (matches.Count > 1)
            throw new BingOfficesConfigurationException(
                $"列表区域尾部标记重复: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        if (matches[0] < firstRow + footer.GapRows)
            throw new BingOfficesConfigurationException(
                $"列表区域尾部标记位于明细起始位置之前: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        return matches[0];
    }

    /// <summary>
    /// 获取尾部从标记前间隔到相对内容末行的占用行数。
    /// </summary>
    /// <param name="footer">尾部定义。</param>
    /// <returns>尾部占用的行数。</returns>
    internal static int GetFooterHeight(IExcelEntityListFooter footer)
    {
        if (footer == null)
            return 0;
        var lastRelativeRow = 0;
        foreach (var cell in footer.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
            lastRelativeRow = Math.Max(lastRelativeRow, cell.Reference.Row);
        foreach (var merge in footer.Merges ?? Array.Empty<ExcelEntityCellRange>())
            lastRelativeRow = Math.Max(lastRelativeRow, merge.Last.Row);
        return lastRelativeRow + 1;
    }

    /// <summary>
    /// 查找实体列表最终尾部和连续分组小计需要跳过的工作表行段。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="sheet">源工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <returns>按起始行排序且不重叠的跳过行段。</returns>
    internal static IReadOnlyList<(int StartRow, int EndRow)> FindFooterSkipSpans<TEntity>(
        ISheet sheet, ExcelEntityListRegion<TEntity> region) where TEntity : class, new()
    {
        var firstRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var lastRow = region.End?.Row ?? sheet.LastRowNum;
        var spans = new List<(int StartRow, int EndRow)>();
        var pageSubtotal = region.PageSubtotal;
        if (pageSubtotal != null && region.PageBreakRows is int rowsPerPage)
        {
            var matches = FindMarkerRows(sheet, region, pageSubtotal.MarkerText, firstRow, lastRow);
            var cursor = firstRow;
            foreach (var markerRow in matches)
            {
                var expected = cursor + rowsPerPage + pageSubtotal.GapRows;
                if (markerRow != expected)
                    throw new BingOfficesConfigurationException(
                        $"列表区域分页小计标记位置无效: {region.SheetName}!{region.Start.Address}",
                        stage: BingOfficesStage.Plan);
                spans.Add((markerRow - pageSubtotal.GapRows,
                    markerRow + GetFooterHeight(pageSubtotal) - 1));
                cursor = markerRow + GetFooterHeight(pageSubtotal);
            }
            var finalMarker = region.Footer == null
                ? -1
                : FindFooterMarkerRow(sheet, region).GetValueOrDefault(-1);
            var detailEnd = region.Footer == null
                ? lastRow + 1
                : finalMarker - region.Footer.GapRows;
            if (matches.Count > 0 && detailEnd <= cursor)
                throw new BingOfficesConfigurationException(
                    $"列表区域分页小计之后缺少明细: {region.SheetName}!{region.Start.Address}",
                    stage: BingOfficesStage.Plan);
            if (detailEnd - cursor > rowsPerPage)
                throw new BingOfficesConfigurationException(
                    $"列表区域缺少分页小计标记: {region.SheetName}!{region.Start.Address}",
                    stage: BingOfficesStage.Plan);
        }
        var group = region.GroupSubtotal;
        if (group != null)
        {
            var matches = new List<int>();
            for (var rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
            {
                var cell = sheet.GetRow(rowIndex)?.GetCell(region.Start.Column);
                var value = cell == null ? null : NpoiExcelImporter.GetRawStringValue(cell);
                if (string.Equals(NpoiExcelImporter.NormalizeText(value, ExcelWhitespacePolicy.Trim),
                        group.Footer.MarkerText, StringComparison.Ordinal))
                    matches.Add(rowIndex);
            }
            foreach (var markerRow in matches)
            {
                if (markerRow < firstRow + group.Footer.GapRows)
                    throw new BingOfficesConfigurationException(
                        $"列表区域分组小计标记位于明细起始位置之前: {region.SheetName}!{region.Start.Address}",
                        stage: BingOfficesStage.Plan);
                spans.Add((markerRow - group.Footer.GapRows,
                    markerRow + GetFooterHeight(group.Footer) - 1));
            }
        }
        if (region.Footer != null)
        {
            var markerRow = FindFooterMarkerRow(sheet, region);
            spans.Add((markerRow.Value - region.Footer.GapRows,
                markerRow.Value + GetFooterHeight(region.Footer) - 1));
        }
        var ordered = spans.OrderBy(span => span.StartRow).ThenBy(span => span.EndRow).ToArray();
        for (var index = 1; index < ordered.Length; index++)
        {
            if (ordered[index].StartRow <= ordered[index - 1].EndRow)
                throw new BingOfficesConfigurationException(
                    $"列表区域尾部与分组小计行段重叠: {region.SheetName}!{region.Start.Address}",
                    stage: BingOfficesStage.Plan);
        }
        return ordered;
    }

    /// <summary>
    /// 查找指定范围内与尾部标记精确匹配的行。
    /// </summary>
    /// <typeparam name="TEntity">根实体类型。</typeparam>
    /// <param name="sheet">源工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="markerText">已规范化的尾部标记。</param>
    /// <param name="firstRow">起始行的零基索引。</param>
    /// <param name="lastRow">结束行的零基索引。</param>
    /// <returns>按工作表顺序排列的匹配行索引。</returns>
    private static IReadOnlyList<int> FindMarkerRows<TEntity>(ISheet sheet,
        ExcelEntityListRegion<TEntity> region, string markerText, int firstRow, int lastRow)
        where TEntity : class, new()
    {
        if (string.IsNullOrWhiteSpace(markerText))
            return Array.Empty<int>();
        var matches = new List<int>();
        for (var rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
        {
            var cell = sheet.GetRow(rowIndex)?.GetCell(region.Start.Column);
            var value = cell == null ? null : NpoiExcelImporter.GetRawStringValue(cell);
            if (string.Equals(NpoiExcelImporter.NormalizeText(value, ExcelWhitespacePolicy.Trim),
                    markerText, StringComparison.Ordinal))
                matches.Add(rowIndex);
        }
        return matches;
    }

    /// <summary>
    /// 判断两个 NPOI 合并区域是否完全相同。
    /// </summary>
    /// <param name="left">左侧合并区域。</param>
    /// <param name="right">右侧合并区域。</param>
    /// <returns>边界相同时返回 <see langword="true" />。</returns>
    private static bool IsExact(CellRangeAddress left, CellRangeAddress right) =>
        left.FirstRow == right.FirstRow && left.LastRow == right.LastRow
        && left.FirstColumn == right.FirstColumn && left.LastColumn == right.LastColumn;
}

/// <summary>
/// NPOI 单个实体导出执行器。
/// </summary>
internal sealed class NpoiEntityExportExecutor
{
    /// <summary>
    /// 记录实体列表区域在运行时占用的工作表矩形。
    /// </summary>
    private sealed record EntityListFootprint(string SheetName, long FirstRow, long LastRow,
        long FirstColumn, long LastColumn)
    {
        /// <summary>
        /// 判断当前矩形是否与另一个矩形相交。
        /// </summary>
        /// <param name="other">待比较的矩形。</param>
        /// <returns>两个矩形相交时返回 <see langword="true" />。</returns>
        public bool Intersects(EntityListFootprint other) =>
            string.Equals(SheetName, other.SheetName, StringComparison.OrdinalIgnoreCase)
            && FirstRow <= other.LastRow && other.FirstRow <= LastRow
            && FirstColumn <= other.LastColumn && other.FirstColumn <= LastColumn;
    }

    /// <summary>
    /// NPOI 导出映射计划构建器。
    /// </summary>
    private readonly NpoiExportPlanBuilder _planBuilder;

    /// <summary>
    /// NPOI 工作表数据写入器。
    /// </summary>
    private readonly NpoiExportSheetWriter _sheetWriter = new NpoiExportSheetWriter();

    /// <summary>
    /// 实体属性值转换器集合。
    /// </summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;

    /// <summary>
    /// 初始化一个 <see cref="NpoiEntityExportExecutor" /> 类型的实例。
    /// </summary>
    /// <param name="mappingPlanFactory">映射计划工厂。</param>
    /// <param name="valueConverters">实体属性值转换器集合。</param>
    internal NpoiEntityExportExecutor(IExcelMappingPlanFactory mappingPlanFactory,
        IReadOnlyList<IExcelValueConverter> valueConverters)
    {
        _planBuilder = new NpoiExportPlanBuilder(mappingPlanFactory ?? throw new ArgumentNullException(nameof(mappingPlanFactory)));
        _valueConverters = valueConverters ?? Array.Empty<IExcelValueConverter>();
    }

    /// <summary>
    /// 按实体布局将固定单元格、列表区域和合并区域写入工作簿。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="entity">待导出的实体。</param>
    /// <param name="layout">实体布局定义。</param>
    /// <param name="template">是否在既有模板上执行写入。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    internal void Write<TEntity>(IWorkbook workbook, TEntity entity, ExcelEntityLayout<TEntity> layout,
        bool template, CancellationToken cancellationToken) where TEntity : class, new()
    {
        if (entity == null)
            throw new ArgumentNullException(nameof(entity));
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));
        var sheetNames = layout.Cells.Select(item => item.SheetName)
            .Concat(layout.ListRegions.Select(item => item.SheetName))
            .Concat(layout.Merges.Select(item => item.SheetName))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        foreach (var name in sheetNames)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = NpoiEntityLayoutSupport.ResolveSheet(workbook, name, template);
            NpoiEntityLayoutSupport.PreflightMerges(sheet, layout.Merges, true, false);
        }
        foreach (var binding in layout.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = NpoiEntityLayoutSupport.ResolveSheet(workbook, binding.SheetName, template);
            var resolvedBinding = binding;
            if (binding.AnchorName != null)
            {
                var named = NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook, binding.SheetName,
                    binding.AnchorName);
                sheet = named.Sheet;
                resolvedBinding = binding.WithResolvedReference(named.Reference);
            }
            var anchor = NpoiEntityLayoutSupport.ResolveAnchor(sheet, resolvedBinding.Reference);
            var cell = NpoiEntityLayoutSupport.GetCell(sheet, anchor, true);
            var value = resolvedBinding.Getter(entity);
            var column = NpoiEntityLayoutSupport.CreateCellPlan(_planBuilder.MappingPlanFactory, resolvedBinding,
                MappingDirection.Export);
            var converted = column.ConvertTo(value, sheet.SheetName, anchor.Row + 1, anchor.Column + 1,
                CultureInfo.InvariantCulture);
            NpoiEntityLayoutSupport.ValidateCellExport(column, value, converted, sheet.SheetName, anchor);
            column.WriteValue(cell, converted);
        }
        var footprints = new List<EntityListFootprint>();
        foreach (var region in layout.ListRegions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = NpoiEntityLayoutSupport.ResolveSheet(workbook, region.SheetName, template);
            var resolvedRegion = region;
            if (region.AnchorName != null)
            {
                var named = NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook, region.SheetName,
                    region.AnchorName);
                sheet = named.Sheet;
                resolvedRegion = region.WithResolvedStart(named.Reference);
            }
            NpoiEntityLayoutSupport.ValidateCellRegionConflict(workbook, layout, resolvedRegion);
            var values = resolvedRegion.Getter(entity)?.Cast<object>().ToArray() ?? Array.Empty<object>();
            WriteRegion(workbook, sheet, entity, resolvedRegion, values, footprints, cancellationToken);
        }
    }

    /// <summary>
    /// 根据列表区域的运行时元素类型分派强类型写入逻辑。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="entity">待导出的实体。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="values">列表元素集合。</param>
    /// <param name="footprints">已写入或预检的列表区域范围。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void WriteRegion<TEntity>(IWorkbook workbook, ISheet sheet, TEntity entity,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<object> values,
        ICollection<EntityListFootprint> footprints, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var method = GetType().GetMethod(nameof(WriteTypedRegion), System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.NonPublic).MakeGenericMethod(typeof(TEntity), region.ItemType);
        InvokeUnwrapped(method, this, new object[] { workbook, sheet, region, values, footprints,
            cancellationToken });
    }

    /// <summary>
    /// 将强类型列表元素写入布局声明的工作表区域。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="values">列表元素集合。</param>
    /// <param name="footprints">已写入或预检的列表区域范围。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void WriteTypedRegion<TEntity, TItem>(IWorkbook workbook, ISheet sheet,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<object> values,
        ICollection<EntityListFootprint> footprints, CancellationToken cancellationToken)
        where TEntity : class, new() where TItem : class, new()
    {
        var data = values.Cast<TItem>().ToArray();
        var groups = region.DynamicColumnGroups ?? Array.Empty<IExcelEntityDynamicColumnGroup>();
        var groupDefinitions = groups.SelectMany(group => group.Definitions).ToArray();
        var request = ExcelExport.Workbook(builder => builder.AddSheet(region.SheetName, data, configure =>
        {
            configure.HeaderRowIndex(0).DataRowStartIndex(1);
            if (region.MappingDocument != null)
                configure.Mapping(region.MappingDocument);
            else if (region.MappingConfiguration != null)
                configure.Mapping(region.MappingConfiguration);
            if (groupDefinitions.Length > 0)
                configure.DynamicColumns(CreateDynamicGetter<TItem>(groups), groupDefinitions)
                    .UnknownDynamicValues(region.UnknownDynamicValues);
        }));
        var sheetRequest = request.Sheets[0];
        var mapping = _planBuilder.Create(request)[sheetRequest];
        var dynamicProperty = groups.Count == 0 ? GetDynamicProperty<TItem>(mapping) : null;
        var dynamicDefinitions = groupDefinitions.Length > 0
            ? groupDefinitions : CreateDynamicDefinitions(mapping);
        var columns = groups.Count == 0
            ? NpoiExportColumnPlanner.CreateColumns<TItem>(mapping, dynamicDefinitions,
                region.CalculatedColumns)
            : NpoiExportColumnPlanner.CreateEntityColumns<TItem>(mapping, dynamicDefinitions, groups,
                region.CalculatedColumns);
        if (region.GroupSubtotal != null)
        {
            WriteGroupedTypedRegion(workbook, sheet, region, data, groups, dynamicProperty,
                dynamicDefinitions, columns, footprints, cancellationToken);
            return;
        }
        NpoiEntityLayoutSupport.ValidateListRegionBounds(region, columns.Count, data.Length, region.IncludeHeader);
        var dataEndRow = region.PageSubtotal != null
            ? NpoiEntityLayoutSupport.ValidatePagedFooterBounds(workbook, region, columns.Count,
                data.Length)
            : region.Start.Row + (region.IncludeHeader ? 1 : 0) + data.Length;
        NpoiEntityLayoutSupport.ValidateFooterBounds(workbook, region, region.Footer, dataEndRow);
        RegisterRuntimeFootprint(region, columns.Count, data.Length, footprints);
        cancellationToken.ThrowIfCancellationRequested();
        if (region.PageSubtotal != null)
        {
            WritePagedTypedRegion(workbook, sheet, region, data, groups, dynamicProperty,
                dynamicDefinitions, columns, cancellationToken);
            return;
        }
        if (region.IncludeHeader)
        {
            if (dynamicDefinitions.Count > 0 && groups.Count == 0)
            {
                var dynamicGetter = CreateDynamicGetter<TItem>(dynamicProperty);
                request = ExcelExport.Workbook(builder => builder.AddSheet(region.SheetName, data, configure =>
                {
                    configure.HeaderRowIndex(0).DataRowStartIndex(1);
                    if (region.MappingDocument != null)
                        configure.Mapping(region.MappingDocument);
                    else if (region.MappingConfiguration != null)
                        configure.Mapping(region.MappingConfiguration);
                    configure.DynamicColumns(dynamicGetter, dynamicDefinitions)
                        .UnknownDynamicValues(region.UnknownDynamicValues);
                }));
                sheetRequest = request.Sheets[0];
            }
            _sheetWriter.Write<TItem>(workbook, sheetRequest, cancellationToken, mapping, columns,
                region.Start.Row, region.Start.Column);
            WriteFooter(workbook, sheet, region.Footer, values,
                dataEndRow, region.Start.Column, cancellationToken, region.FooterAnchorName);
            NpoiEntityLayoutSupport.ApplyPageBreaks(sheet, region, data.Length);
            return;
        }
        var rowIndex = region.Start.Row;
        var calculatedItems = columns.Any(column => column.IsCalculated)
            ? data.Cast<object>().ToArray()
            : null;
        var calculatedEvaluators = columns.Where(column => column.IsCalculated)
            .ToDictionary(column => column.Key,
                column => column.Calculated.CreateEvaluator(calculatedItems, sheet.SheetName),
                StringComparer.OrdinalIgnoreCase);
        for (var itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            var item = data[itemIndex];
            cancellationToken.ThrowIfCancellationRequested();
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var dynamicValues = groups.Count == 0
                ? dynamicProperty?.GetValue(item) as IDictionary<string, object>
                : MergeDynamicValues(item, groups);
            if (region.UnknownDynamicValues == ExcelUnknownDynamicValuePolicy.Fail)
            {
                var knownKeys = new HashSet<string>(dynamicDefinitions.Select(definition => definition.Key),
                    StringComparer.OrdinalIgnoreCase);
                var unknown = dynamicValues?.Keys.FirstOrDefault(key => !knownKeys.Contains(key));
                if (unknown != null)
                    throw new BingOfficesConfigurationException($"动态值未声明: {unknown}",
                        stage: BingOfficesStage.Validate);
            }
            for (var index = 0; index < columns.Count; index++)
            {
                var column = columns[index];
                var cell = row.GetCell(region.Start.Column + index) ?? row.CreateCell(region.Start.Column + index);
                NpoiExportSheetWriter.WriteCell(cell, item, column, dynamicValues, sheet.SheetName,
                    rowIndex + 1, region.Start.Column + index + 1, CultureInfo.InvariantCulture,
                    itemIndex, calculatedEvaluators.TryGetValue(column.Key, out var evaluator)
                        ? evaluator : null);
                if (column.BodyStyle != null)
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, column.BodyStyle);
            }
            rowIndex++;
        }
        WriteFooter(workbook, sheet, region.Footer, values, rowIndex, region.Start.Column,
            cancellationToken, region.FooterAnchorName);
        NpoiEntityLayoutSupport.ApplyPageBreaks(sheet, region, data.Length);
    }

    /// <summary>
    /// 将带分页小计的强类型列表写入 NPOI 工作表。
    /// </summary>
    private static void WritePagedTypedRegion<TEntity, TItem>(IWorkbook workbook, ISheet sheet,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<TItem> data,
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups, PropertyInfo dynamicProperty,
        IReadOnlyList<ExcelDynamicColumnDefinition> dynamicDefinitions,
        IReadOnlyList<ExcelColumnPlan> columns, CancellationToken cancellationToken)
        where TEntity : class, new() where TItem : class, new()
    {
        if (region.PageSubtotal == null || region.PageBreakRows is not int rowsPerPage)
            throw new BingOfficesConfigurationException("分页小计缺少分页配置。",
                stage: BingOfficesStage.Plan);
        var plannedRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var plannedIndex = 0;
        while (plannedIndex < data.Count)
        {
            var pageCount = Math.Min(rowsPerPage, data.Count - plannedIndex);
            plannedRow += pageCount;
            plannedIndex += pageCount;
            if (plannedIndex < data.Count)
            {
                NpoiEntityLayoutSupport.PreflightFooterMerges(sheet, region.PageSubtotal,
                    plannedRow, region.Start.Column);
                plannedRow += region.PageSubtotal.GapRows
                    + NpoiEntityLayoutSupport.GetFooterHeight(region.PageSubtotal);
            }
        }
        NpoiEntityLayoutSupport.PreflightFooterMerges(sheet, region.Footer, plannedRow,
            region.Start.Column);
        if (region.IncludeHeader)
        {
            var header = sheet.GetRow(region.Start.Row) ?? sheet.CreateRow(region.Start.Row);
            for (var index = 0; index < columns.Count; index++)
            {
                var cell = header.GetCell(region.Start.Column + index)
                    ?? header.CreateCell(region.Start.Column + index);
                cell.SetCellValue(columns[index].Title);
                if (columns[index].HeaderStyle != null)
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle,
                        columns[index].HeaderStyle);
            }
        }
        var calculatedItems = columns.Any(column => column.IsCalculated)
            ? data.Cast<object>().ToArray() : null;
        var calculatedEvaluators = columns.Where(column => column.IsCalculated)
            .ToDictionary(column => column.Key,
                column => column.Calculated.CreateEvaluator(calculatedItems, sheet.SheetName),
                StringComparer.OrdinalIgnoreCase);
        var rowIndex = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var itemIndex = 0;
        var detailRows = new List<(int FirstRow, int LastRow)>();
        while (itemIndex < data.Count)
        {
            var firstDetailRow = rowIndex + 1;
            var pageStart = itemIndex;
            var pageCount = Math.Min(rowsPerPage, data.Count - itemIndex);
            for (var pageIndex = 0; pageIndex < pageCount; pageIndex++)
            {
                var item = data[itemIndex];
                cancellationToken.ThrowIfCancellationRequested();
                if (item == null)
                    throw new BingOfficesConfigurationException("实体列表区域包含 null 项。",
                        stage: BingOfficesStage.Plan);
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                var dynamicValues = groups.Count == 0
                    ? dynamicProperty?.GetValue(item) as IDictionary<string, object>
                    : MergeDynamicValues(item, groups);
                if (region.UnknownDynamicValues == ExcelUnknownDynamicValuePolicy.Fail)
                {
                    var knownKeys = new HashSet<string>(dynamicDefinitions.Select(definition => definition.Key),
                        StringComparer.OrdinalIgnoreCase);
                    var unknown = dynamicValues?.Keys.FirstOrDefault(key => !knownKeys.Contains(key));
                    if (unknown != null)
                        throw new BingOfficesConfigurationException($"动态值未声明: {unknown}",
                            stage: BingOfficesStage.Validate);
                }
                for (var index = 0; index < columns.Count; index++)
                {
                    var column = columns[index];
                    var cell = row.GetCell(region.Start.Column + index)
                        ?? row.CreateCell(region.Start.Column + index);
                    NpoiExportSheetWriter.WriteCell(cell, item, column, dynamicValues, sheet.SheetName,
                        rowIndex + 1, region.Start.Column + index + 1, CultureInfo.InvariantCulture,
                        itemIndex, calculatedEvaluators.TryGetValue(column.Key, out var evaluator)
                            ? evaluator : null);
                    if (column.BodyStyle != null)
                        cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, column.BodyStyle);
                }
                rowIndex++;
                itemIndex++;
            }
            detailRows.Add((firstDetailRow, rowIndex));
            if (itemIndex < data.Count)
            {
                var pageItems = data.Skip(pageStart).Take(pageCount).Cast<object>().ToArray();
                WriteFooter(workbook, sheet, region.PageSubtotal, pageItems, rowIndex,
                    region.Start.Column, cancellationToken);
                rowIndex += region.PageSubtotal.GapRows
                    + NpoiEntityLayoutSupport.GetFooterHeight(region.PageSubtotal);
                NpoiEntityLayoutSupport.ApplyPageBreak(sheet, rowIndex - 1);
            }
        }
        WriteFooter(workbook, sheet, region.Footer, data.Cast<object>().ToArray(), rowIndex,
            region.Start.Column, cancellationToken, region.FooterAnchorName, detailRows);
    }

    /// <summary>
    /// 将带连续分组小计的列表项写入 NPOI 工作表。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="data">列表元素集合。</param>
    /// <param name="groups">动态列组集合。</param>
    /// <param name="dynamicProperty">单字典动态属性。</param>
    /// <param name="dynamicDefinitions">动态列定义。</param>
    /// <param name="columns">最终列计划。</param>
    /// <param name="footprints">已写入或预检的列表区域范围。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void WriteGroupedTypedRegion<TEntity, TItem>(IWorkbook workbook, ISheet sheet,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<TItem> data,
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups, PropertyInfo dynamicProperty,
        IReadOnlyList<ExcelDynamicColumnDefinition> dynamicDefinitions,
        IReadOnlyList<ExcelColumnPlan> columns, ICollection<EntityListFootprint> footprints,
        CancellationToken cancellationToken)
        where TEntity : class, new() where TItem : class, new()
    {
        var subtotal = region.GroupSubtotal;
        var firstDataRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var rowIndex = firstDataRow;
        var calculatedItems = columns.Any(column => column.IsCalculated)
            ? data.Cast<object>().ToArray() : null;
        var calculatedEvaluators = columns.Where(column => column.IsCalculated)
            .ToDictionary(column => column.Key,
                column => column.Calculated.CreateEvaluator(calculatedItems, sheet.SheetName),
                StringComparer.OrdinalIgnoreCase);
        var groupRanges = GetGroupedRanges(data, subtotal);
        var previewRow = firstDataRow;
        foreach (var range in groupRanges)
        {
            var dataEndRow = previewRow + range.Count;
            NpoiEntityLayoutSupport.ValidateFooterBounds(workbook, region, subtotal.Footer, dataEndRow);
            previewRow = dataEndRow + subtotal.Footer.GapRows
                + NpoiEntityLayoutSupport.GetFooterHeight(subtotal.Footer);
        }
        NpoiEntityLayoutSupport.ValidateFooterBounds(workbook, region, region.Footer, previewRow);
        var lastRow = previewRow + (region.Footer == null ? -1
            : region.Footer.GapRows + NpoiEntityLayoutSupport.GetFooterHeight(region.Footer) - 1);
        if (region.End.HasValue && (lastRow > region.End.Value.Row
            || region.Start.Column + columns.Count - 1 > region.End.Value.Column))
            throw new BingOfficesConfigurationException(
                $"列表区域分组小计超出声明边界: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        RegisterRuntimeFootprint(region, columns.Count, data.Count, footprints, data.Cast<object>().ToArray());
        cancellationToken.ThrowIfCancellationRequested();

        if (region.IncludeHeader)
        {
            var header = sheet.GetRow(region.Start.Row) ?? sheet.CreateRow(region.Start.Row);
            for (var index = 0; index < columns.Count; index++)
            {
                var cell = header.GetCell(region.Start.Column + index)
                    ?? header.CreateCell(region.Start.Column + index);
                cell.SetCellValue(columns[index].Title);
                if (columns[index].HeaderStyle != null)
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle,
                        columns[index].HeaderStyle);
            }
        }

        void WriteItem(TItem item, int itemIndex)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var dynamicValues = groups.Count == 0
                ? dynamicProperty?.GetValue(item) as IDictionary<string, object>
                : MergeDynamicValues(item, groups);
            if (region.UnknownDynamicValues == ExcelUnknownDynamicValuePolicy.Fail)
            {
                var knownKeys = new HashSet<string>(dynamicDefinitions.Select(definition => definition.Key),
                    StringComparer.OrdinalIgnoreCase);
                var unknown = dynamicValues?.Keys.FirstOrDefault(key => !knownKeys.Contains(key));
                if (unknown != null)
                    throw new BingOfficesConfigurationException($"动态值未声明: {unknown}",
                        stage: BingOfficesStage.Validate);
            }
            for (var index = 0; index < columns.Count; index++)
            {
                var column = columns[index];
                var cell = row.GetCell(region.Start.Column + index)
                    ?? row.CreateCell(region.Start.Column + index);
                NpoiExportSheetWriter.WriteCell(cell, item, column, dynamicValues, sheet.SheetName,
                    rowIndex + 1, region.Start.Column + index + 1, CultureInfo.InvariantCulture,
                    itemIndex, calculatedEvaluators.TryGetValue(column.Key, out var evaluator)
                        ? evaluator : null);
                if (column.BodyStyle != null)
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, column.BodyStyle);
            }
            rowIndex++;
        }

        var itemIndexOffset = 0;
        var detailRows = new List<(int FirstRow, int LastRow)>();
        foreach (var range in groupRanges)
        {
            var firstDetailRow = rowIndex + 1;
            var groupItems = data.Skip(range.Start).Take(range.Count).Cast<object>().ToArray();
            for (var itemIndex = 0; itemIndex < range.Count; itemIndex++)
                WriteItem(data[range.Start + itemIndex], itemIndexOffset + itemIndex);
            detailRows.Add((firstDetailRow, rowIndex));
            WriteFooter(workbook, sheet, subtotal.Footer, groupItems, rowIndex,
                region.Start.Column, cancellationToken);
            rowIndex += subtotal.Footer.GapRows + NpoiEntityLayoutSupport.GetFooterHeight(subtotal.Footer);
            itemIndexOffset += range.Count;
        }
        WriteFooter(workbook, sheet, region.Footer, data.Cast<object>().ToArray(), rowIndex,
            region.Start.Column, cancellationToken, region.FooterAnchorName, detailRows);
    }

    /// <summary>
    /// 按相邻分组键计算列表元素的连续区间。
    /// </summary>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    /// <param name="data">列表元素集合。</param>
    /// <param name="subtotal">分组小计定义。</param>
    /// <returns>连续分组的起始索引和元素数量。</returns>
    private static IReadOnlyList<(int Start, int Count)> GetGroupedRanges<TItem>(
        IReadOnlyList<TItem> data, IExcelEntityGroupSubtotal subtotal) where TItem : class, new()
    {
        var ranges = new List<(int Start, int Count)>();
        if (data == null || data.Count == 0)
            return ranges;
        var start = 0;
        var key = subtotal.GetKey(data[0]);
        for (var index = 1; index < data.Count; index++)
        {
            var next = subtotal.GetKey(data[index]);
            if (Equals(key, next))
                continue;
            ranges.Add((start, index - start));
            start = index;
            key = next;
        }
        ranges.Add((start, data.Count - start));
        return ranges;
    }

    /// <summary>
    /// 预检并登记列表区域的运行时占用范围。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="region">列表区域定义。</param>
    /// <param name="columnCount">映射计划的列数。</param>
    /// <param name="itemCount">实际明细行数。</param>
    /// <param name="footprints">已登记的列表区域范围。</param>
    /// <param name="values">列表项快照；配置分组小计时用于计算运行时行数。</param>
    private static void RegisterRuntimeFootprint<TEntity>(ExcelEntityListRegion<TEntity> region,
        int columnCount, int itemCount, ICollection<EntityListFootprint> footprints,
        IReadOnlyList<object> values = null)
        where TEntity : class, new()
    {
        var dataRows = (long)itemCount + (region.IncludeHeader ? 1 : 0);
        var hasData = dataRows > 0;
        var firstRow = (long)region.Start.Row;
        var lastRow = hasData ? firstRow + dataRows - 1 : firstRow - 1;
        var firstColumn = (long)region.Start.Column;
        var lastColumn = hasData ? firstColumn + Math.Max(0, columnCount - 1) : firstColumn - 1;
        if (region.GroupSubtotal != null)
        {
            var row = (long)region.Start.Row + (region.IncludeHeader ? 1 : 0);
            foreach (var range in GetGroupedRanges(values ?? Array.Empty<object>(), region.GroupSubtotal))
                row += range.Count + region.GroupSubtotal.Footer.GapRows
                    + NpoiEntityLayoutSupport.GetFooterHeight(region.GroupSubtotal.Footer);
            lastRow = Math.Max(lastRow, row - 1);
            foreach (var cell in region.GroupSubtotal.Footer.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
                lastColumn = Math.Max(lastColumn, firstColumn + cell.Reference.Column);
            foreach (var merge in region.GroupSubtotal.Footer.Merges ?? Array.Empty<ExcelEntityCellRange>())
                lastColumn = Math.Max(lastColumn, firstColumn + merge.Last.Column);
        }
        if (region.PageSubtotal != null && region.PageBreakRows is int rowsPerPage)
        {
            var row = (long)region.Start.Row + (region.IncludeHeader ? 1 : 0);
            var written = 0;
            while (written < itemCount)
            {
                var count = Math.Min(rowsPerPage, itemCount - written);
                row += count;
                written += count;
                if (written >= itemCount)
                    break;
                row += region.PageSubtotal.GapRows
                    + NpoiEntityLayoutSupport.GetFooterHeight(region.PageSubtotal);
                lastRow = Math.Max(lastRow, row - 1);
                foreach (var cell in region.PageSubtotal.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
                    lastColumn = Math.Max(lastColumn, firstColumn + cell.Reference.Column);
                foreach (var merge in region.PageSubtotal.Merges ?? Array.Empty<ExcelEntityCellRange>())
                    lastColumn = Math.Max(lastColumn, firstColumn + merge.Last.Column);
            }
        }
        if (region.Footer != null)
        {
            var footerBaseRow = region.GroupSubtotal != null || region.PageSubtotal != null
                ? lastRow + 1 : firstRow + dataRows;
            var footerRow = footerBaseRow + region.Footer.GapRows;
            lastRow = Math.Max(lastRow, footerRow);
            lastColumn = Math.Max(lastColumn, firstColumn);
            foreach (var cell in region.Footer.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
            {
                lastRow = Math.Max(lastRow, footerRow + cell.Reference.Row);
                lastColumn = Math.Max(lastColumn, firstColumn + cell.Reference.Column);
            }
            foreach (var merge in region.Footer.Merges ?? Array.Empty<ExcelEntityCellRange>())
            {
                lastRow = Math.Max(lastRow, footerRow + merge.Last.Row);
                lastColumn = Math.Max(lastColumn, firstColumn + merge.Last.Column);
            }
        }
        if (!hasData && region.Footer == null)
            return;
        var current = new EntityListFootprint(region.SheetName, firstRow, lastRow, firstColumn, lastColumn);
        if (footprints.Any(existing => current.Intersects(existing)))
            throw new BingOfficesConfigurationException(
                $"实体列表区域运行时范围与已有区域重叠: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        footprints.Add(current);
    }

    /// <summary>
    /// 将列表尾部标记、聚合值和相对合并区域写入 NPOI 工作表。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="footer">列表尾部定义。</param>
    /// <param name="items">已物化的列表项快照。</param>
    /// <param name="dataEndRow">明细结束后的零基行索引。</param>
    /// <param name="startColumn">列表区域起始零基列索引。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="anchorName">可选的最终尾部命名锚点。</param>
    /// <param name="detailRows">最终尾部使用的一基明细行段。</param>
    private static void WriteFooter(IWorkbook workbook, ISheet sheet, IExcelEntityListFooter footer,
        IReadOnlyList<object> items, int dataEndRow, int startColumn, CancellationToken cancellationToken,
        string anchorName = null, IReadOnlyList<(int FirstRow, int LastRow)> detailRows = null)
    {
        if (footer == null)
            return;
        cancellationToken.ThrowIfCancellationRequested();
        var footerRow = dataEndRow + footer.GapRows;
        var marker = sheet.GetRow(footerRow) ?? sheet.CreateRow(footerRow);
        var markerCell = marker.GetCell(startColumn) ?? marker.CreateCell(startColumn);
        markerCell.SetCellValue(footer.MarkerText, footer.MarkerStyle?.NumberFormat);
        if (footer.MarkerStyle != null)
            markerCell.CellStyle = NpoiStyleCache.Compose(workbook, markerCell.CellStyle, footer.MarkerStyle);
        foreach (var definition in footer.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = sheet.GetRow(footerRow + definition.Reference.Row)
                ?? sheet.CreateRow(footerRow + definition.Reference.Row);
            var cell = row.GetCell(startColumn + definition.Reference.Column)
                ?? row.CreateCell(startColumn + definition.Reference.Column);
            var value = definition.Evaluate(items);
            var formulaA1 = (value as ExcelEntityFooterFormulaValue)?.FormulaA1;
            if (value is ExcelEntityFooterDetailSumValue detailSum)
            {
                var rows = detailRows ?? (items.Count == 0
                    ? Array.Empty<(int FirstRow, int LastRow)>()
                    : new[] { (dataEndRow - items.Count + 1, dataEndRow) });
                formulaA1 = detailSum.ToFormulaA1(rows);
            }
            if (formulaA1 != null)
            {
                cell.SetCellFormula(formulaA1);
                if (definition.Style != null)
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, definition.Style);
                if (!string.IsNullOrWhiteSpace(definition.NumberFormat))
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle,
                        new ExcelCellStyle { NumberFormat = definition.NumberFormat });
            }
            else
            {
                cell.SetCellValue(value, definition.NumberFormat ?? definition.Style?.NumberFormat);
                if (definition.Style != null)
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, definition.Style);
            }
        }
        var merges = (footer.Merges ?? Array.Empty<ExcelEntityCellRange>()).Select(merge =>
            new CellRangeAddress(footerRow + merge.First.Row, footerRow + merge.Last.Row,
                startColumn + merge.First.Column, startColumn + merge.Last.Column)).ToArray();
        NpoiEntityLayoutSupport.PreflightFooterMerges(sheet, merges);
        if (anchorName != null)
            UpdateFooterAnchor(workbook, sheet, anchorName, footerRow, startColumn);
    }

    /// <summary>
    /// 将工作簿名称定位到最终尾部标记单元格。
    /// </summary>
    /// <param name="workbook">待导出的工作簿。</param>
    /// <param name="sheet">尾部所在工作表。</param>
    /// <param name="anchorName">命名锚点名称。</param>
    /// <param name="row">标记的零基行索引。</param>
    /// <param name="column">标记的零基列索引。</param>
    private static void UpdateFooterAnchor(IWorkbook workbook, ISheet sheet, string anchorName,
        int row, int column)
    {
        var matches = workbook.GetAllNames().Where(name =>
            string.Equals(name.NameName, anchorName, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (matches.Length > 1)
            throw new BingOfficesConfigurationException(
                $"最终尾部命名锚点不唯一: {sheet.SheetName}!{anchorName}",
                stage: BingOfficesStage.Plan);
        if (matches.Length == 1)
            NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook, sheet.SheetName, anchorName);
        var name = matches.Length == 1 ? matches[0] : workbook.CreateName();
        if (matches.Length == 0)
            name.NameName = anchorName;
        var escapedSheet = sheet.SheetName.Replace("'", "''");
        name.RefersToFormula = $"'{escapedSheet}'!{new CellReference(row, column, true, true).FormatAsString()}";
    }

    /// <summary>
    /// 获取列表项上声明动态列的字典属性。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="mapping">列表项映射计划。</param>
    /// <returns>动态字典属性；没有动态列时为 <see langword="null" />。</returns>
    private static PropertyInfo GetDynamicProperty<TItem>(IExcelMappingPlan mapping)
        where TItem : class, new()
    {
        var properties = (mapping?.Columns ?? Array.Empty<IExcelMappingColumn>())
            .Where(column => column.IsDynamicColumn).ToArray();
        if (properties.Length == 0)
            return null;
        if (properties.Length > 1)
            throw new BingOfficesConfigurationException(
                $"实体列表项 {typeof(TItem).FullName} 只能声明一个动态列字典属性。",
                stage: BingOfficesStage.Plan);
        var property = typeof(TItem).GetProperty(properties[0].Name,
            BindingFlags.Instance | BindingFlags.Public);
        if (property == null || !typeof(IDictionary<string, object>).IsAssignableFrom(property.PropertyType))
            throw new BingOfficesConfigurationException(
                $"动态列属性必须实现 IDictionary<string, object>: {properties[0].Name}",
                stage: BingOfficesStage.Plan);
        return property;
    }

    /// <summary>
    /// 将映射计划中的动态列转换为实体布局导出定义。
    /// </summary>
    /// <param name="mapping">列表项映射计划。</param>
    /// <returns>按映射顺序排列的动态列定义。</returns>
    private static IReadOnlyList<ExcelDynamicColumnDefinition> CreateDynamicDefinitions(
        IExcelMappingPlan mapping) => (mapping?.DynamicColumns ?? Array.Empty<IExcelDynamicMappingColumn>())
        .Select(column => NpoiExportColumnPlanner.CreateDynamicDefinition(column, null, mapping))
        .ToArray();

    /// <summary>
    /// 创建列表项动态字典的表达式读取器。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="property">动态字典属性。</param>
    /// <returns>可供标准导出请求使用的动态值表达式。</returns>
    private static Expression<Func<TItem, IDictionary<string, object>>> CreateDynamicGetter<TItem>(
        PropertyInfo property) where TItem : class, new()
    {
        if (property == null)
            throw new ArgumentNullException(nameof(property));
        var parameter = Expression.Parameter(typeof(TItem), "item");
        Expression body = Expression.Property(parameter, property);
        if (body.Type != typeof(IDictionary<string, object>))
            body = Expression.Convert(body, typeof(IDictionary<string, object>));
        return Expression.Lambda<Func<TItem, IDictionary<string, object>>>(body, parameter);
    }

    /// <summary>
    /// 创建多个实体动态列组的合并字典读取器。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="groups">动态列组集合。</param>
    /// <returns>将各组动态值合并到同一字典的表达式。</returns>
    private static Expression<Func<TItem, IDictionary<string, object>>> CreateDynamicGetter<TItem>(
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups) where TItem : class, new()
    {
        if (groups == null || groups.Count == 0)
            throw new ArgumentException("动态列组不能为空。", nameof(groups));
        var parameter = Expression.Parameter(typeof(TItem), "item");
        var method = typeof(NpoiEntityExportExecutor).GetMethod(nameof(MergeDynamicValues),
            BindingFlags.Static | BindingFlags.NonPublic).MakeGenericMethod(typeof(TItem));
        var body = Expression.Call(method, parameter, Expression.Constant(groups));
        return Expression.Lambda<Func<TItem, IDictionary<string, object>>>(body, parameter);
    }

    /// <summary>
    /// 合并一个列表项的多个动态值字典。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="item">列表项。</param>
    /// <param name="groups">动态列组集合。</param>
    /// <returns>按动态列键合并的字典。</returns>
    private static IDictionary<string, object> MergeDynamicValues<TItem>(TItem item,
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups) where TItem : class, new()
    {
        var values = new Dictionary<string, object>(StringComparer.Ordinal);
        foreach (var group in groups)
        {
            var source = group.Getter(item);
            if (source == null)
                continue;
            foreach (var pair in source)
                values[pair.Key] = pair.Value;
        }
        return values;
    }

    /// <summary>
    /// 调用反射方法，并保留内部异常的原始堆栈后重新抛出。
    /// </summary>
    /// <param name="method">待调用的方法。</param>
    /// <param name="target">方法目标实例。</param>
    /// <param name="arguments">方法参数。</param>
    private static void InvokeUnwrapped(MethodInfo method, object target, object[] arguments)
    {
        try
        {
            method.Invoke(target, arguments);
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }
}

/// <summary>
/// NPOI 单个实体导入执行器。
/// </summary>
internal sealed class NpoiEntityImportExecutor
{
    /// <summary>
    /// NPOI 导入映射计划构建器。
    /// </summary>
    private readonly NpoiImportPlanBuilder _planBuilder;

    /// <summary>
    /// NPOI 工作表导入执行器。
    /// </summary>
    private readonly NpoiImportSheetExecutor _sheetExecutor = new NpoiImportSheetExecutor(new NpoiImportRowMaterializer());

    /// <summary>
    /// 单元格值转换器集合。
    /// </summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;

    /// <summary>
    /// 初始化一个 <see cref="NpoiEntityImportExecutor" /> 类型的实例。
    /// </summary>
    /// <remarks>
    /// 此重载使用空的值转换器集合。
    /// </remarks>
    /// <param name="mappingPlanFactory">映射计划工厂。</param>
    internal NpoiEntityImportExecutor(IExcelMappingPlanFactory mappingPlanFactory)
    {
        _planBuilder = new NpoiImportPlanBuilder(mappingPlanFactory ?? throw new ArgumentNullException(nameof(mappingPlanFactory)));
        _valueConverters = Array.Empty<IExcelValueConverter>();
    }

    /// <summary>
    /// 初始化一个 <see cref="NpoiEntityImportExecutor" /> 类型的实例。
    /// </summary>
    /// <remarks>
    /// 此重载使用调用方提供的值转换器集合。
    /// </remarks>
    /// <param name="mappingPlanFactory">映射计划工厂。</param>
    /// <param name="valueConverters">单元格值转换器集合。</param>
    internal NpoiEntityImportExecutor(IExcelMappingPlanFactory mappingPlanFactory,
        IReadOnlyList<IExcelValueConverter> valueConverters)
    {
        _planBuilder = new NpoiImportPlanBuilder(mappingPlanFactory ?? throw new ArgumentNullException(nameof(mappingPlanFactory)));
        _valueConverters = valueConverters ?? Array.Empty<IExcelValueConverter>();
    }

    /// <summary>
    /// 按实体布局从工作簿读取固定单元格、列表区域和合并区域。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="workbook">源工作簿。</param>
    /// <param name="entity">接收导入值的实体。</param>
    /// <param name="layout">实体布局定义。</param>
    /// <param name="requireTemplateMerges">是否要求声明的合并区域已存在。</param>
    /// <param name="limits">实体导入资源限制。</param>
    /// <param name="validationFailureMode">当前行校验失败后的继续策略。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    /// <returns>包含实体、工作表结果和结构化错误的导入结果。</returns>
    internal ExcelEntityImportResult<TEntity> Read<TEntity>(IWorkbook workbook, TEntity entity,
        ExcelEntityLayout<TEntity> layout, bool requireTemplateMerges, ExcelResourceLimits limits,
        ExcelValidationFailureMode validationFailureMode,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        limits?.Validate();
        var errors = new ExcelImportErrorCollector(limits?.MaxErrors);
        var sheetResults = new List<ExcelSheetImportResult>();
        var runtime = new ExcelImportRuntime(limits);
        foreach (var name in layout.Cells.Select(item => item.SheetName)
                     .Concat(layout.ListRegions.Select(item => item.SheetName))
                     .Concat(layout.Merges.Select(item => item.SheetName))
                     .Distinct(StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = workbook.GetSheet(name);
            if (sheet == null)
                throw new BingOfficesConfigurationException($"工作簿缺少请求的 Sheet: {name}", stage: BingOfficesStage.Plan);
            NpoiEntityLayoutSupport.PreflightMerges(sheet, layout.Merges, false, requireTemplateMerges);
        }
        foreach (var binding in layout.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = workbook.GetSheet(binding.SheetName);
            var resolvedBinding = binding;
            if (binding.AnchorName != null)
            {
                var named = NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook, binding.SheetName,
                    binding.AnchorName);
                sheet = named.Sheet;
                resolvedBinding = binding.WithResolvedReference(named.Reference);
            }
            var anchor = NpoiEntityLayoutSupport.ResolveAnchor(sheet, resolvedBinding.Reference);
            if (resolvedBinding.Setter == null)
                throw new BingOfficesConfigurationException($"实体属性不可写入: {resolvedBinding.PropertyName}",
                    stage: BingOfficesStage.Plan);
            try
            {
                var column = NpoiEntityLayoutSupport.CreateCellPlan(_planBuilder.MappingPlanFactory, resolvedBinding,
                    MappingDirection.Import);
                var cellValue = NpoiExcelImporter.ReadCellValue(NpoiEntityLayoutSupport.GetCell(sheet, anchor, false),
                    workbook.IsDate1904());
                var value = NpoiExcelImporter.NormalizeText(cellValue.Text,
                    column.Property.ImportWhitespace ?? ExcelWhitespacePolicy.Trim);
                var normalizedCellValue = new ExcelCellValue(cellValue.Value, value, cellValue.Kind,
                    cellValue.CachedKind, cellValue.Formula, cellValue.ErrorCode, cellValue.FormatIndex,
                    cellValue.IsDate1904);
                if (!NpoiEntityLayoutSupport.ValidateCellImportRaw(column, normalizedCellValue, value,
                        sheet.SheetName, anchor, errors))
                    continue;
                var converted = column.ConvertFrom(value, normalizedCellValue, sheet.SheetName,
                    anchor.Row + 1, anchor.Column + 1, CultureInfo.InvariantCulture);
                if (!NpoiEntityLayoutSupport.ValidateCellImportConverted(column, normalizedCellValue, value,
                        converted, sheet.SheetName, anchor, errors))
                    continue;
                resolvedBinding.Setter(entity, converted);
            }
            catch (BingOfficesException)
            {
                throw;
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.ValueConversion, exception.Message,
                    sheet.SheetName, anchor.Row + 1, anchor.Column + 1, binding.PropertyName,
                    binding.PropertyName, null));
            }
        }
        foreach (var region in layout.ListRegions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = workbook.GetSheet(region.SheetName);
            var resolvedRegion = region;
            if (region.AnchorName != null)
            {
                var named = NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook, region.SheetName,
                    region.AnchorName);
                sheet = named.Sheet;
                resolvedRegion = region.WithResolvedStart(named.Reference);
            }
            NpoiEntityLayoutSupport.ValidateCellRegionConflict(workbook, layout, resolvedRegion);
            ReadRegion(workbook, sheet, entity, resolvedRegion, errors, sheetResults, runtime,
                limits, validationFailureMode, cancellationToken);
        }
        var sourceLocations = new Dictionary<object, SourceLocation>();
        foreach (var relation in layout.Relations ?? Array.Empty<ExcelRelationRequest>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            NpoiRelationBinder.Bind(entity, relation, errors, sourceLocations, cancellationToken);
        }
        return new ExcelEntityImportResult<TEntity>(entity, errors.Errors, sheetResults, errors.IsTruncated,
            errors.MaxErrors);
    }

    /// <summary>
    /// 根据列表区域的运行时元素类型分派强类型读取逻辑。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="workbook">源工作簿。</param>
    /// <param name="sheet">源工作表。</param>
    /// <param name="entity">接收列表数据的实体。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="errors">导入错误收集器。</param>
    /// <param name="sheetResults">工作表结果集合。</param>
    /// <param name="runtime">导入运行时状态。</param>
    /// <param name="limits">实体导入资源限制。</param>
    /// <param name="validationFailureMode">当前行校验失败后的继续策略。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void ReadRegion<TEntity>(IWorkbook workbook, ISheet sheet, TEntity entity,
        ExcelEntityListRegion<TEntity> region, ExcelImportErrorCollector errors,
        ICollection<ExcelSheetImportResult> sheetResults, ExcelImportRuntime runtime,
        ExcelResourceLimits limits,
        ExcelValidationFailureMode validationFailureMode,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        var method = GetType().GetMethod(nameof(ReadTypedRegion), System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.NonPublic).MakeGenericMethod(typeof(TEntity), region.ItemType);
        InvokeUnwrapped(method, this, new object[] { workbook, sheet, entity, region, errors, sheetResults, runtime,
            limits, validationFailureMode, cancellationToken });
    }

    /// <summary>
    /// 从布局区域读取强类型列表，并将临时工作表行号映射回源工作表。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    /// <param name="workbook">源工作簿。</param>
    /// <param name="source">源工作表。</param>
    /// <param name="entity">接收列表数据的实体。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="errors">导入错误收集器。</param>
    /// <param name="sheetResults">工作表结果集合。</param>
    /// <param name="runtime">导入运行时状态。</param>
    /// <param name="limits">实体导入资源限制。</param>
    /// <param name="validationFailureMode">当前行校验失败后的继续策略。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void ReadTypedRegion<TEntity, TItem>(IWorkbook workbook, ISheet source, TEntity entity,
        ExcelEntityListRegion<TEntity> region, ExcelImportErrorCollector errors,
        ICollection<ExcelSheetImportResult> sheetResults, ExcelImportRuntime runtime,
        ExcelResourceLimits limits,
        ExcelValidationFailureMode validationFailureMode,
        CancellationToken cancellationToken) where TEntity : class, new() where TItem : class, new()
    {
        var groups = region.DynamicColumnGroups ?? Array.Empty<IExcelEntityDynamicColumnGroup>();
        var groupDefinitions = groups.SelectMany(group => group.Definitions).ToArray();
        var request = ExcelImport.Workbook<RegionRoot<TItem>>(builder => builder.Sheet(region.SheetName,
            root => root.Items, configure =>
            {
                configure.HeaderRowIndex(0).DataRowStartIndex(1);
                if (region.MappingDocument != null)
                    configure.Mapping(region.MappingDocument);
                else if (region.MappingConfiguration != null)
                    configure.Mapping(region.MappingConfiguration);
                if (groupDefinitions.Length > 0)
                    configure.DynamicColumns(CreateDynamicTargetGetter<TItem>(groups), groupDefinitions);
                configure.FailOnUnknownDynamicColumns(
                    region.UnknownDynamicValues == ExcelUnknownDynamicValuePolicy.Fail);
            }));
        var resolved = new NpoiResolvedSheet(request.Sheets[0], 0, region.SheetName);
        var mapping = _planBuilder.Create(new[] { resolved })[request.Sheets[0]];
        var dynamicProperty = groups.Count == 0 ? GetDynamicProperty<TItem>(mapping) : null;
        var dynamicDefinitions = groupDefinitions.Length > 0
            ? groupDefinitions : CreateDynamicDefinitions(mapping);
        var columns = groups.Count == 0
            ? NpoiExportColumnPlanner.CreateColumns<TItem>(mapping, dynamicDefinitions,
                region.CalculatedColumns)
            : NpoiExportColumnPlanner.CreateEntityColumns<TItem>(mapping, dynamicDefinitions, groups,
                region.CalculatedColumns);
        var dynamicTargets = new List<(TItem Item, IDictionary<string, object> Values)>();
        var itemCount = region.End.HasValue
            ? region.End.Value.Row - region.Start.Row + (region.IncludeHeader ? 0 : 1)
            : 0;
        NpoiEntityLayoutSupport.ValidateListRegionBounds(region, columns.Count, itemCount, region.IncludeHeader);
        var temp = CreateRegionSheet<TEntity, TItem>(workbook, source, region, columns.Count, columns,
            workbook.IsDate1904(),
            cancellationToken);
        try
        {
            var options = new ExcelImportExecutionOptions<TItem>
            {
                HeaderRowIndex = 0,
                DataRowIndex = 1,
                MaxReadColumns = Math.Max(1, columns.Count),
                MappingPlan = mapping,
                Culture = CultureInfo.InvariantCulture,
                IsDate1904 = workbook.IsDate1904(),
                HeaderWhitespace = ExcelWhitespacePolicy.Trim,
                BodyWhitespace = ExcelWhitespacePolicy.Trim,
                ValidationFailureMode = validationFailureMode,
                RequireExpectedHeaders = true,
                ValidationMode = ExcelImportValidationMode.ConfiguredRules,
                MaxTrackedUniqueValues = limits?.MaxTrackedUniqueValues,
                UniqueComparison = limits?.UniqueComparison ?? StringComparison.OrdinalIgnoreCase,
                StopAtFirstEmptyRow = false,
                DynamicColumns = dynamicDefinitions,
                DynamicTargetGetter = groups.Count == 0 && dynamicProperty != null
                    ? item => dynamicProperty.GetValue(item)
                    : groups.Count == 0 ? null : item => CreateTemporaryDynamicValues(item, dynamicTargets),
                DynamicPropertyNames = groups.Count == 0 ? null :
                    new HashSet<string>(groups.Select(group => group.Property.Name), StringComparer.OrdinalIgnoreCase),
                DynamicColumnPropertyNames = groups.Count == 0 ? null : groupDefinitions
                    .ToDictionary(definition => definition.Key,
                        definition => groups.Single(group => group.Definitions.Any(column =>
                            string.Equals(column.Key, definition.Key, StringComparison.OrdinalIgnoreCase)))
                            .Property.Name, StringComparer.OrdinalIgnoreCase)
                , ErrorSheetName = region.SheetName
            };
            var items = new List<TItem>();
            var rows = new List<int>();
            var child = errors.CreateChild();
            _sheetExecutor.Execute(temp, options, items, child, runtime, cancellationToken, rows);
            if (groups.Count > 0)
                ApplyDynamicGroups(dynamicTargets, groups);
            AssignItems(entity, region, items);
            var sourceRows = rows.Select(row => region.Start.Row + row - (region.IncludeHeader ? 0 : 1)).ToArray();
            sheetResults.Add(new ExcelSheetImportResult(source.SheetName, typeof(TItem), sourceRows, child.Errors));
        }
        finally
        {
            workbook.RemoveSheetAt(workbook.GetSheetIndex(temp));
        }
    }

    /// <summary>
    /// 创建实体布局动态列组的导入目标表达式。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="groups">动态列组集合。</param>
    /// <returns>供标准导入请求保存动态定义的目标表达式。</returns>
    private static Expression<Func<TItem, IDictionary<string, object>>> CreateDynamicTargetGetter<TItem>(
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups) where TItem : class, new()
    {
        var parameter = Expression.Parameter(typeof(TItem), "item");
        var method = typeof(NpoiEntityImportExecutor).GetMethod(nameof(CreateDynamicTarget),
            BindingFlags.Static | BindingFlags.NonPublic).MakeGenericMethod(typeof(TItem));
        var body = Expression.Call(method, parameter, Expression.Constant(groups));
        return Expression.Lambda<Func<TItem, IDictionary<string, object>>>(body, parameter);
    }

    /// <summary>
    /// 创建标准导入请求保存的临时动态目标字典。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="item">列表项。</param>
    /// <param name="groups">动态列组集合。</param>
    /// <returns>新的动态值字典。</returns>
    private static IDictionary<string, object> CreateDynamicTarget<TItem>(TItem item,
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups) where TItem : class, new() =>
        new Dictionary<string, object>(StringComparer.Ordinal);

    /// <summary>
    /// 创建单行导入使用的临时动态字典，并保留其所属实体。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="item">标准导入器传入的列表项。</param>
    /// <param name="targets">临时目标集合。</param>
    /// <returns>供标准导入器写入的临时字典。</returns>
    private static IDictionary<string, object> CreateTemporaryDynamicValues<TItem>(object item,
        ICollection<(TItem Item, IDictionary<string, object> Values)> targets) where TItem : class, new()
    {
        var values = new Dictionary<string, object>(StringComparer.Ordinal);
        targets.Add(((TItem)item, values));
        return values;
    }

    /// <summary>
    /// 将标准导入器收集的动态值分发回各实体动态列组。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="targets">实体与临时动态字典集合。</param>
    /// <param name="groups">动态列组集合。</param>
    private static void ApplyDynamicGroups<TItem>(
        IEnumerable<(TItem Item, IDictionary<string, object> Values)> targets,
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups) where TItem : class, new()
    {
        foreach (var target in targets)
        {
            foreach (var group in groups)
            {
                var keys = new HashSet<string>(group.Definitions.Select(definition => definition.Key),
                    StringComparer.OrdinalIgnoreCase);
                var values = target.Values.Where(pair => keys.Contains(pair.Key)).ToArray();
                if (values.Length == 0)
                    continue;
                var destination = group.Getter(target.Item);
                if (destination == null)
                {
                    if (group.Setter == null)
                        throw new BingOfficesConfigurationException(
                            $"实体动态列属性不可写入: {group.Property.Name}", stage: BingOfficesStage.Plan);
                    try
                    {
                        var type = group.Property.PropertyType;
                        var instance = type.IsInterface || type.IsAbstract
                            ? new Dictionary<string, object>(StringComparer.Ordinal)
                            : Activator.CreateInstance(type);
                        if (instance is not IDictionary<string, object> created)
                            throw new InvalidOperationException();
                        destination = created;
                        group.Setter(target.Item, destination);
                    }
                    catch (Exception exception) when (exception is not BingOfficesException
                        && exception is not OperationCanceledException
                        && exception is not OutOfMemoryException && exception is not StackOverflowException)
                    {
                        throw new BingOfficesConfigurationException(
                            $"实体动态列属性无法创建字典实例: {group.Property.Name}", exception,
                            BingOfficesStage.Plan);
                    }
                }
                foreach (var pair in values)
                    destination[pair.Key] = pair.Value;
            }
        }
    }

    /// <summary>
    /// 获取映射计划中声明的单个动态字典属性。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="mapping">列表项映射计划。</param>
    /// <returns>动态字典属性；未声明时返回 <see langword="null" />。</returns>
    private static PropertyInfo GetDynamicProperty<TItem>(IExcelMappingPlan mapping)
        where TItem : class, new()
    {
        var properties = (mapping?.Columns ?? Array.Empty<IExcelMappingColumn>())
            .Where(column => column.IsDynamicColumn).ToArray();
        if (properties.Length == 0)
            return null;
        if (properties.Length > 1)
            throw new BingOfficesConfigurationException(
                $"导入实体列表项 {typeof(TItem).FullName} 只能声明一个动态列字典属性。",
                stage: BingOfficesStage.Plan);
        var property = typeof(TItem).GetProperty(properties[0].Name,
            BindingFlags.Instance | BindingFlags.Public);
        if (property == null || !typeof(IDictionary<string, object>).IsAssignableFrom(property.PropertyType))
            throw new BingOfficesConfigurationException(
                $"动态列属性必须实现 IDictionary<string, object>: {properties[0].Name}",
                stage: BingOfficesStage.Plan);
        return property;
    }

    /// <summary>
    /// 将映射计划中的动态列转换为实体布局导入定义。
    /// </summary>
    /// <param name="mapping">列表项映射计划。</param>
    /// <returns>按映射顺序排列的动态列定义。</returns>
    private static IReadOnlyList<ExcelDynamicColumnDefinition> CreateDynamicDefinitions(
        IExcelMappingPlan mapping) => (mapping?.DynamicColumns ?? Array.Empty<IExcelDynamicMappingColumn>())
        .Select(column => NpoiExportColumnPlanner.CreateDynamicDefinition(column, null, mapping))
        .ToArray();

    /// <summary>
    /// 将源列表区域投影到标准零基标题和数据行布局的临时工作表。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    /// <param name="workbook">承载临时工作表的工作簿。</param>
    /// <param name="source">源工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="width">区域列数。</param>
    /// <param name="columns">导入列计划。</param>
    /// <param name="isDate1904">工作簿是否使用 1904 日期系统。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    /// <returns>供标准工作表导入执行器读取的临时工作表。</returns>
    private static ISheet CreateRegionSheet<TEntity, TItem>(IWorkbook workbook, ISheet source,
        ExcelEntityListRegion<TEntity> region, int width, IReadOnlyList<ExcelColumnPlan> columns,
        bool isDate1904, CancellationToken cancellationToken) where TEntity : class, new()
        where TItem : class, new()
    {
        var temp = workbook.CreateSheet("__BingEnt" + Guid.NewGuid().ToString("N").Substring(0, 12));
        var header = temp.CreateRow(0);
        var sourceHeader = region.IncludeHeader ? source.GetRow(region.Start.Row) : null;
        for (var index = 0; index < width; index++)
        {
            var headerCell = header.CreateCell(index);
            if (sourceHeader != null)
                CopyCell(sourceHeader.GetCell(region.Start.Column + index), headerCell, isDate1904);
            else
                headerCell.SetCellValue(columns[index].Title);
        }
        var startRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var footerMarkerRow = NpoiEntityLayoutSupport.FindFooterMarkerRow(source, region);
        var lastRow = footerMarkerRow.HasValue
            ? footerMarkerRow.Value - 1 - (region.Footer?.GapRows ?? 0)
            : (region.End?.Row ?? source.LastRowNum);
        var skipSpans = region.GroupSubtotal == null && region.PageSubtotal == null
            ? Array.Empty<(int StartRow, int EndRow)>()
            : NpoiEntityLayoutSupport.FindFooterSkipSpans(source, region)
                .Where(span => span.StartRow <= lastRow).ToArray();
        var outputRow = 1;
        for (var rowIndex = startRow; rowIndex <= lastRow; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (skipSpans.Any(span => rowIndex >= span.StartRow && rowIndex <= span.EndRow))
                continue;
            var sourceRow = source.GetRow(rowIndex);
            var targetRow = temp.CreateRow(outputRow);
            for (var column = 0; column < width; column++)
                CopyCell(sourceRow?.GetCell(region.Start.Column + column), targetRow.CreateCell(column), isDate1904);
            outputRow++;
        }
        return temp;
    }

    /// <summary>
    /// 将源单元格的样式和值复制到临时工作表单元格。
    /// </summary>
    /// <param name="source">源单元格。</param>
    /// <param name="target">目标单元格。</param>
    /// <param name="isDate1904">工作簿是否使用 1904 日期系统。</param>
    private static void CopyCell(ICell source, ICell target, bool isDate1904)
    {
        if (source == null)
            return;
        target.CellStyle = source.CellStyle;
        var value = NpoiExcelImporter.ReadCellValue(source, isDate1904);
        if (value.Kind == ExcelCellKind.Empty)
            return;
        target.SetCellValue(value.Value ?? value.Text);
    }

    /// <summary>
    /// 将导入的列表元素赋给实体集合属性或填充现有集合实例。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    /// <param name="entity">接收列表数据的实体。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="items">已导入的列表元素。</param>
    private static void AssignItems<TEntity, TItem>(TEntity entity, ExcelEntityListRegion<TEntity> region,
        IReadOnlyList<TItem> items) where TEntity : class, new() where TItem : class, new()
    {
        if (region.Setter != null)
        {
            region.Setter(entity, items.ToList());
            return;
        }
        var current = region.Getter(entity) as ICollection<TItem>;
        if (current == null)
            throw new BingOfficesConfigurationException($"实体集合属性不可写入: {region.PropertyName}",
                stage: BingOfficesStage.Plan);
        current.Clear();
        foreach (var item in items)
            current.Add(item);
    }

    /// <summary>
    /// 为标准工作表导入管线提供列表属性的临时根对象。
    /// </summary>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    private sealed class RegionRoot<TItem> where TItem : class, new()
    {
        /// <summary>
        /// 获取接收导入数据的列表集合。
        /// </summary>
        public ICollection<TItem> Items { get; } = new List<TItem>();
    }

    /// <summary>
    /// 调用反射方法，并保留内部异常的原始堆栈后重新抛出。
    /// </summary>
    /// <param name="method">待调用的方法。</param>
    /// <param name="target">方法目标实例。</param>
    /// <param name="arguments">方法参数。</param>
    private static void InvokeUnwrapped(MethodInfo method, object target, object[] arguments)
    {
        try
        {
            method.Invoke(target, arguments);
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }
}
