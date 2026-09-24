using System.Collections;
using System.Collections.Concurrent;
using System.Globalization;
using System.Reflection;
using System.Runtime.ExceptionServices;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Providers;
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
}

/// <summary>
/// NPOI 单个实体导出执行器。
/// </summary>
internal sealed class NpoiEntityExportExecutor
{
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
            var anchor = NpoiEntityLayoutSupport.ResolveAnchor(sheet, binding.Reference);
            var cell = NpoiEntityLayoutSupport.GetCell(sheet, anchor, true);
            var value = binding.Getter(entity);
            var column = NpoiEntityLayoutSupport.CreateCellPlan(_planBuilder.MappingPlanFactory, binding,
                MappingDirection.Export);
            var converted = column.ConvertTo(value, sheet.SheetName, anchor.Row + 1, anchor.Column + 1,
                CultureInfo.InvariantCulture);
            NpoiEntityLayoutSupport.ValidateCellExport(column, value, converted, sheet.SheetName, anchor);
            column.WriteValue(cell, converted);
        }
        foreach (var region in layout.ListRegions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = NpoiEntityLayoutSupport.ResolveSheet(workbook, region.SheetName, template);
            var values = region.Getter(entity)?.Cast<object>().ToArray() ?? Array.Empty<object>();
            WriteRegion(workbook, sheet, entity, region, values, cancellationToken);
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
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void WriteRegion<TEntity>(IWorkbook workbook, ISheet sheet, TEntity entity,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<object> values, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var method = GetType().GetMethod(nameof(WriteTypedRegion), System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.NonPublic).MakeGenericMethod(typeof(TEntity), region.ItemType);
        InvokeUnwrapped(method, this, new object[] { workbook, sheet, region, values, cancellationToken });
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
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void WriteTypedRegion<TEntity, TItem>(IWorkbook workbook, ISheet sheet,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<object> values, CancellationToken cancellationToken)
        where TEntity : class, new() where TItem : class, new()
    {
        var data = values.Cast<TItem>().ToArray();
        var request = ExcelExport.Workbook(builder => builder.AddSheet(region.SheetName, data, configure =>
        {
            configure.HeaderRowIndex(0).DataRowStartIndex(1);
            if (region.MappingDocument != null)
                configure.Mapping(region.MappingDocument);
            else if (region.MappingConfiguration != null)
                configure.Mapping(region.MappingConfiguration);
        }));
        var sheetRequest = request.Sheets[0];
        var mapping = _planBuilder.Create(request)[sheetRequest];
        if (mapping.DynamicColumns.Count > 0)
            throw new BingOfficesUnsupportedFeatureException("实体列表区域暂不支持动态列。", provider: "NPOI",
                operation: BingOfficesOperation.Export, stage: BingOfficesStage.Plan);
        var columns = NpoiExportColumnPlanner.CreateColumns<TItem>(mapping,
            Array.Empty<ExcelDynamicColumnDefinition>());
        NpoiEntityLayoutSupport.ValidateListRegionBounds(region, columns.Count, data.Length, region.IncludeHeader);
        if (region.IncludeHeader)
        {
            _sheetWriter.Write<TItem>(workbook, sheetRequest, cancellationToken, mapping, columns,
                region.Start.Row, region.Start.Column);
            return;
        }
        var rowIndex = region.Start.Row;
        foreach (var item in data)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            for (var index = 0; index < columns.Count; index++)
            {
                var column = columns[index];
                var cell = row.GetCell(region.Start.Column + index) ?? row.CreateCell(region.Start.Column + index);
                var value = column.ConvertTo(column.Getter(item), sheet.SheetName, rowIndex + 1,
                    region.Start.Column + index + 1, CultureInfo.InvariantCulture);
                cell.SetCellValue(value, column.Formatter);
            }
            rowIndex++;
        }
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
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    /// <returns>包含实体、工作表结果和结构化错误的导入结果。</returns>
    internal ExcelEntityImportResult<TEntity> Read<TEntity>(IWorkbook workbook, TEntity entity,
        ExcelEntityLayout<TEntity> layout, bool requireTemplateMerges, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var errors = new ExcelImportErrorCollector(null);
        var sheetResults = new List<ExcelSheetImportResult>();
        var runtime = new ExcelImportRuntime(null);
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
            var anchor = NpoiEntityLayoutSupport.ResolveAnchor(sheet, binding.Reference);
            if (binding.Setter == null)
                throw new BingOfficesConfigurationException($"实体属性不可写入: {binding.PropertyName}",
                    stage: BingOfficesStage.Plan);
            try
            {
                var column = NpoiEntityLayoutSupport.CreateCellPlan(_planBuilder.MappingPlanFactory, binding,
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
                binding.Setter(entity, converted);
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
            ReadRegion(workbook, sheet, entity, region, errors, sheetResults, runtime, cancellationToken);
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
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void ReadRegion<TEntity>(IWorkbook workbook, ISheet sheet, TEntity entity,
        ExcelEntityListRegion<TEntity> region, ExcelImportErrorCollector errors,
        ICollection<ExcelSheetImportResult> sheetResults, ExcelImportRuntime runtime,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        var method = GetType().GetMethod(nameof(ReadTypedRegion), System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.NonPublic).MakeGenericMethod(typeof(TEntity), region.ItemType);
        InvokeUnwrapped(method, this, new object[] { workbook, sheet, entity, region, errors, sheetResults, runtime,
            cancellationToken });
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
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void ReadTypedRegion<TEntity, TItem>(IWorkbook workbook, ISheet source, TEntity entity,
        ExcelEntityListRegion<TEntity> region, ExcelImportErrorCollector errors,
        ICollection<ExcelSheetImportResult> sheetResults, ExcelImportRuntime runtime,
        CancellationToken cancellationToken) where TEntity : class, new() where TItem : class, new()
    {
        var request = ExcelImport.Workbook<RegionRoot<TItem>>(builder => builder.Sheet(region.SheetName,
            root => root.Items, configure =>
            {
                configure.HeaderRowIndex(0).DataRowStartIndex(1);
                if (region.MappingDocument != null)
                    configure.Mapping(region.MappingDocument);
                else if (region.MappingConfiguration != null)
                    configure.Mapping(region.MappingConfiguration);
            }));
        var resolved = new NpoiResolvedSheet(request.Sheets[0], 0, region.SheetName);
        var mapping = _planBuilder.Create(new[] { resolved })[request.Sheets[0]];
        if (mapping.DynamicColumns.Count > 0)
            throw new BingOfficesUnsupportedFeatureException("实体列表区域暂不支持动态列。", provider: "NPOI",
                operation: BingOfficesOperation.Import, stage: BingOfficesStage.Plan);
        var columns = NpoiExportColumnPlanner.CreateColumns<TItem>(mapping,
            Array.Empty<ExcelDynamicColumnDefinition>());
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
                ValidationFailureMode = ExcelValidationFailureMode.StopOnFirstFailure,
                RequireExpectedHeaders = true,
                ValidationMode = ExcelImportValidationMode.ConfiguredRules,
                StopAtFirstEmptyRow = false
            };
            var items = new List<TItem>();
            var rows = new List<int>();
            var child = errors.CreateChild();
            _sheetExecutor.Execute(temp, options, items, child, runtime, cancellationToken, rows);
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
        var lastRow = region.End?.Row ?? source.LastRowNum;
        var outputRow = 1;
        for (var rowIndex = startRow; rowIndex <= lastRow; rowIndex++, outputRow++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sourceRow = source.GetRow(rowIndex);
            var targetRow = temp.CreateRow(outputRow);
            for (var column = 0; column < width; column++)
                CopyCell(sourceRow?.GetCell(region.Start.Column + column), targetRow.CreateCell(column), isDate1904);
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
