using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using Bing.Offices.Conversions;
using Bing.Offices.Mappings;
using Bing.Offices.MiniExcel.Internals;
using Bing.Offices.Providers;
using Bing.Offices.Validations;

namespace Bing.Offices.Imports;

/// <summary>
/// 负责 MiniExcel 行对象到目标实体的物化、转换和校验。
/// </summary>
/// <remarks>
/// 该类型只消费已完成的列计划，不参与工作表读取、关系绑定或文件资源提交。
/// </remarks>
internal static class MiniExcelRowMaterializer
{
    /// <summary>
    /// 缓存按实体类型编译的实例工厂。
    /// </summary>
    private static readonly ConcurrentDictionary<Type, Func<object>> ItemFactories =
        new ConcurrentDictionary<Type, Func<object>>();

    /// <summary>
    /// 将一行原始值转换为实体，执行校验并收集行级错误。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="row">当前行的表头和值。</param>
    /// <param name="headers">工作表表头顺序。</param>
    /// <param name="sheetName">工作表实际名称。</param>
    /// <param name="request">当前工作表导入配置。</param>
    /// <param name="plan">当前行类型的映射计划。</param>
    /// <param name="bindings">已解析的列绑定。</param>
    /// <param name="rowNumber">当前行的一基物理行号。</param>
    /// <param name="items">接收成功实体的集合。</param>
    /// <param name="rows">接收成功行索引的集合。</param>
    /// <param name="sheetErrors">接收当前工作表错误的集合。</param>
    /// <param name="unique">当前工作表的唯一性跟踪器。</param>
    /// <param name="workbookRequest">当前工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消行处理的令牌。</param>
    /// <param name="itemType">行实体的运行时类型。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="rawDateSerials">按物理行列索引的原始日期 serial 集合。</param>
    internal static void MaterializeRow<TWorkbook>(IDictionary<string, object> row,
        IReadOnlyList<string> headers, string sheetName, ExcelSheetImportRequest request,
        IExcelMappingPlan plan, IReadOnlyList<MiniExcelExcelImporter.ColumnBinding> bindings,
        int rowNumber, IList items, ICollection<int> rows, ICollection<ExcelImportError> sheetErrors,
        UniqueTracker unique, ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        CancellationToken cancellationToken, Type itemType, bool isDate1904,
        IReadOnlyDictionary<long, double> rawDateSerials) where TWorkbook : class, new()
    {
        var item = ItemFactories.GetOrAdd(itemType, CreateItemFactory)();
        var valid = true;
        var configuredValidationEnabled = IsConfiguredValidationEnabled(workbookRequest.ValidationMode);
        if (configuredValidationEnabled)
            unique.BeginRow();
        var dynamicValues = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < bindings.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var binding = bindings[index];
            row.TryGetValue(binding.Header, out var raw);
            var text = MiniExcelSheetPlanBuilder.Normalize(MiniExcelValueAdapter.ToText(raw, request.Culture),
                binding.Column.ImportWhitespace ?? request.BodyWhitespace);
            var columnIndex = binding.ColumnIndex;
            var cell = CreateRawDateCell(rawDateSerials, rowNumber, columnIndex, text, isDate1904,
                binding.Property.PropertyType);
            try
            {
                if (configuredValidationEnabled)
                    ValidateBindings(binding.Column.ValidationBindings, text, null, sheetName, rowNumber,
                        columnIndex, binding.Column.Name, raw, request.Culture, binding.Property.PropertyType,
                        unique, binding.Column.IsUnique, binding.Column.UniqueIgnoreEmpty,
                        sheetErrors, isDate1904: isDate1904, cell: cell);
                var converted = MiniExcelValueAdapter.ConvertFrom(raw, binding.Column, binding.Property,
                    sheetName, rowNumber, columnIndex, request.Culture, isDate1904, cell);
                if (configuredValidationEnabled)
                    ValidateBindings(binding.Column.ValidationBindings, text, converted, sheetName, rowNumber,
                        columnIndex, binding.Column.Name, raw, request.Culture, binding.Property.PropertyType,
                        unique, binding.Column.IsUnique, binding.Column.UniqueIgnoreEmpty,
                        sheetErrors, rawOnly: false, isDate1904: isDate1904, cell: cell);
                binding.Setter(item, converted);
            }
            catch (MiniExcelExcelImporter.MiniExcelRowException exception)
            {
                AddSheetError(sheetErrors, workbookRequest, exception.Error);
                valid = false;
                if (request.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                    break;
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                AddSheetError(sheetErrors, workbookRequest, new ExcelImportError(
                    ExcelImportErrorCode.ValueConversion, exception.Message, sheetName, rowNumber,
                    columnIndex, binding.Column.Name, rawValue: raw));
                valid = false;
                if (request.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                    break;
            }
        }

        foreach (var dynamic in plan.DynamicColumns)
        {
            var header = MiniExcelSheetPlanBuilder.FindHeader(row.Keys, dynamic.Title, dynamic.Aliases,
                request.HeaderComparison, request.HeaderWhitespace);
            if (header == null)
                continue;
            row.TryGetValue(header, out var raw);
            var columnIndex = MiniExcelSheetPlanBuilder.FindPhysicalColumnIndex(headers, header,
                request.ReadColumnRange?.StartIndex ?? 0);
            var cell = CreateRawDateCell(rawDateSerials, rowNumber, columnIndex,
                MiniExcelValueAdapter.ToText(raw, request.Culture), isDate1904,
                MiniExcelValueAdapter.ResolveDynamicType(dynamic.DataTypeName));
            try
            {
                var converted = MiniExcelValueAdapter.ConvertDynamicFrom(raw, dynamic, sheetName,
                    rowNumber, columnIndex, request.Culture, isDate1904, cell);
                dynamicValues[dynamic.Key] = converted;
                if (configuredValidationEnabled)
                    ValidateBindings(dynamic.ValidationBindings, MiniExcelValueAdapter.ToText(raw, request.Culture),
                        converted, sheetName, rowNumber, columnIndex, dynamic.Key, raw, request.Culture,
                        MiniExcelValueAdapter.ResolveDynamicType(dynamic.DataTypeName), unique, dynamic.IsUnique,
                        dynamic.UniqueIgnoreEmpty, sheetErrors, rawOnly: false, isDate1904: isDate1904,
                        cell: cell);
            }
            catch (MiniExcelExcelImporter.MiniExcelRowException exception)
            {
                AddSheetError(sheetErrors, workbookRequest, exception.Error);
                valid = false;
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                AddSheetError(sheetErrors, workbookRequest, new ExcelImportError(
                    ExcelImportErrorCode.ValueConversion, exception.Message, sheetName, rowNumber, columnIndex,
                    dynamic.Key, rawValue: raw));
                valid = false;
            }
        }
        if (valid)
        {
            SetDynamicValues(item, request, dynamicValues);
            items.Add(item);
            rows.Add(rowNumber - 1);
            if (configuredValidationEnabled)
                unique.CommitRow();
        }
        else if (configuredValidationEnabled)
            unique.RollbackRow();
    }

    /// <summary>
    /// 编译实体的公开无参构造器。
    /// </summary>
    /// <param name="itemType">待创建实例的实体类型。</param>
    /// <returns>创建实体实例的委托。</returns>
    internal static Func<object> CreateItemFactory(Type itemType)
    {
        var body = Expression.Convert(Expression.New(itemType), typeof(object));
        return Expression.Lambda<Func<object>>(body).Compile();
    }

    /// <summary>
    /// 编译实体属性写入器，并在 Sheet 计划构建时复用。
    /// </summary>
    /// <param name="property">待写入的实体属性。</param>
    /// <returns>将值写入属性的委托。</returns>
    internal static Action<object, object> CreatePropertySetter(PropertyInfo property)
    {
        var instance = Expression.Parameter(typeof(object), "instance");
        var value = Expression.Parameter(typeof(object), "value");
        var assignment = Expression.Assign(
            Expression.Property(Expression.Convert(instance, property.DeclaringType), property),
            Expression.Convert(value, property.PropertyType));
        return Expression.Lambda<Action<object, object>>(assignment, instance, value).Compile();
    }

    /// <summary>
    /// 为日期属性恢复原始 Excel 数值单元格。
    /// </summary>
    /// <param name="rawDateSerials">按物理坐标保存的原始日期序列值。</param>
    /// <param name="rowNumber">单元格所在的一基行号。</param>
    /// <param name="columnNumber">单元格所在的一基列号。</param>
    /// <param name="text">单元格的文本缓存值。</param>
    /// <param name="isDate1904">是否使用 1904 日期系统。</param>
    /// <param name="propertyType">目标属性类型。</param>
    /// <returns>日期属性对应的原始单元格值；不适用或无法恢复时返回 <see langword="null"/>。</returns>
    private static ExcelCellValue CreateRawDateCell(IReadOnlyDictionary<long, double> rawDateSerials,
        int rowNumber, int columnNumber, string text, bool isDate1904, Type propertyType)
    {
        var targetType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
        if (targetType != typeof(DateTime) && targetType != typeof(DateTimeOffset))
            return null;
        if (rawDateSerials != null
            && rawDateSerials.TryGetValue(MiniExcelRawDateSerialReader.CreateKey(rowNumber, columnNumber),
                out var serial))
            return new ExcelCellValue(serial, text, ExcelCellKind.Number, isDate1904: isDate1904);
        return null;
    }

    /// <summary>
    /// 按原始值或转换值执行列校验和唯一性校验。
    /// </summary>
    /// <param name="bindings">待执行的列校验绑定。</param>
    /// <param name="text">单元格文本值。</param>
    /// <param name="converted">转换后的值。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowNumber">单元格所在的一基行号。</param>
    /// <param name="columnNumber">单元格所在的一基列号。</param>
    /// <param name="propertyName">目标属性名称。</param>
    /// <param name="raw">未转换的原始值。</param>
    /// <param name="culture">转换和校验使用的区域性。</param>
    /// <param name="propertyType">目标属性类型。</param>
    /// <param name="unique">唯一性跟踪器。</param>
    /// <param name="isUnique">是否执行唯一性校验。</param>
    /// <param name="ignoreEmpty">是否忽略空值的唯一性校验。</param>
    /// <param name="errors">接收导入错误的集合。</param>
    /// <param name="rawOnly">是否仅执行原始值校验。</param>
    /// <param name="isDate1904">是否使用 1904 日期系统。</param>
    /// <param name="cell">与当前值对应的公共单元格描述。</param>
    private static void ValidateBindings(IReadOnlyList<IExcelValidationBinding> bindings, string text,
        object converted, string sheetName, int rowNumber, int columnNumber, string propertyName,
        object raw, CultureInfo culture, Type propertyType, UniqueTracker unique,
        bool isUnique, bool ignoreEmpty, ICollection<ExcelImportError> errors,
        bool rawOnly = true, bool isDate1904 = false, ExcelCellValue cell = null)
    {
        if (bindings == null)
            bindings = Array.Empty<IExcelValidationBinding>();
        foreach (var binding in bindings)
        {
            if (!rawOnly && binding.IsRaw)
                continue;
            if (rawOnly && !binding.IsRaw)
                continue;
            if (!rawOnly && binding.Kind == ExcelValidationBindingKind.Unique)
                continue;
            if (!binding.Validate(new ExcelValidationContext(text, sheetName, rowNumber, columnNumber,
                propertyName, converted, propertyType,
                cell ?? MiniExcelValueAdapter.CreateCell(raw, text, isDate1904), culture)))
            {
                throw new MiniExcelExcelImporter.MiniExcelRowException(new ExcelImportError(
                    MiniExcelValueAdapter.GetValidationCode(binding), binding.ErrorMessage, sheetName,
                    rowNumber, columnNumber, propertyName, rawValue: raw));
            }
        }
        if (!rawOnly && isUnique && !unique.TryReserve(propertyName, text, false, ignoreEmpty, rowNumber))
        {
            unique.TryGetFirstRowNumber(propertyName, text, out var firstRow);
            throw new MiniExcelExcelImporter.MiniExcelRowException(new ExcelImportError(
                ExcelImportErrorCode.Validation, "重复数据。", sheetName, rowNumber, columnNumber,
                propertyName, rawValue: raw, firstRowNumber: firstRow == 0 ? null : firstRow));
        }
    }

    /// <summary>
    /// 将动态列值写入实体的动态目标成员。
    /// </summary>
    /// <param name="item">接收动态列值的实体。</param>
    /// <param name="request">当前工作表导入配置。</param>
    /// <param name="values">待写入的动态列值。</param>
    private static void SetDynamicValues(object item, ExcelSheetImportRequest request,
        IDictionary<string, object> values)
    {
        if (request.DynamicTarget == null || values.Count == 0)
            return;
        var member = (request.DynamicTarget as LambdaExpression)?.Body as MemberExpression;
        if (member?.Member is PropertyInfo property && property.CanWrite)
        {
            var current = property.GetValue(item) as IDictionary<string, object>;
            if (current == null)
            {
                current = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                property.SetValue(item, current);
            }
            foreach (var pair in values)
                current[pair.Key] = pair.Value;
            return;
        }
        var target = request.DynamicTargetGetter?.Invoke(item) as IDictionary<string, object>;
        if (target != null)
            foreach (var pair in values)
                target[pair.Key] = pair.Value;
    }

    /// <summary>
    /// 判断当前模式是否启用配置级校验。
    /// </summary>
    /// <param name="mode">当前导入校验模式。</param>
    /// <returns>模式包含配置规则时返回 true，否则返回 false。</returns>
    private static bool IsConfiguredValidationEnabled(ExcelImportValidationMode mode) =>
        mode == ExcelImportValidationMode.ConfiguredRules
        || mode == ExcelImportValidationMode.ConfiguredAndWorkbook;

    /// <summary>
    /// 在未达到错误上限时追加一个工作表错误。
    /// </summary>
    /// <typeparam name="TWorkbook">工作簿根实体类型。</typeparam>
    /// <param name="errors">接收错误的集合。</param>
    /// <param name="request">当前工作簿导入请求。</param>
    /// <param name="error">待追加的错误。</param>
    private static void AddSheetError<TWorkbook>(ICollection<ExcelImportError> errors,
        ExcelWorkbookImportRequest<TWorkbook> request, ExcelImportError error)
        where TWorkbook : class, new()
    {
        if (request.ResourceLimits?.MaxErrors is int maximum && errors.Count >= maximum)
            return;
        errors.Add(error);
    }
}
