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

namespace Bing.Offices.Npoi.Entities;

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
