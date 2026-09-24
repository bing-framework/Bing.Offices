using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;

namespace Bing.Offices.Imports;

/// <summary>
/// 负责 MiniExcel 工作表表头、列绑定和物理坐标计划。
/// </summary>
/// <remarks>
/// 该职责在行枚举开始前完成，行物化器只消费已解析的绑定，避免逐行重复查找表头。
/// </remarks>
internal static class MiniExcelSheetPlanBuilder
{
    /// <summary>
    /// 根据固定映射计划计算潜在日期列，供 RawDate 补偿读取使用。
    /// </summary>
    /// <typeparam name="TItem">工作表行实体类型。</typeparam>
    /// <param name="plan">当前工作表的映射计划。</param>
    /// <param name="request">当前工作表的导入配置。</param>
    /// <returns>可确定的日期列号集合；布局不确定时返回 <see langword="null"/>。</returns>
    internal static IReadOnlyCollection<int> ResolveDateColumns<TItem>(IExcelMappingPlan plan,
        ExcelSheetImportRequest request) where TItem : class, new()
    {
        var result = new HashSet<int>();
        var physicalColumn = (request.ReadColumnRange?.StartIndex ?? 0) + 1;
        var hasAmbiguousLayout = request.DynamicColumns.Count > 0
            || !request.RequireExpectedHeaders
            || !request.FailOnUnknownDynamicColumns
            || plan.Columns.Any(column => column.Ignored || column.IsDynamicColumn);
        foreach (var column in plan.Columns.Where(column => !column.IsDynamicColumn))
        {
            var property = typeof(TItem).GetProperty(column.Name,
                BindingFlags.Instance | BindingFlags.Public);
            if (property != null && IsNavigationOrDynamicContainer(property.PropertyType))
            {
                hasAmbiguousLayout = true;
                physicalColumn++;
                continue;
            }
            var propertyType = property?.PropertyType;
            var targetType = Nullable.GetUnderlyingType(propertyType ?? typeof(object)) ?? propertyType;
            if (!column.Ignored && (targetType == typeof(DateTime) || targetType == typeof(DateTimeOffset)))
                result.Add(physicalColumn);
            physicalColumn++;
        }
        if (result.Count == 0)
            return Array.Empty<int>();
        // 表头可能包含未知列、忽略列或动态布局；无法在计划期确定物理坐标时回退旧扫描。
        return hasAmbiguousLayout ? null : result;
    }

    /// <summary>
    /// 根据映射计划和表头创建可写入的列绑定。
    /// </summary>
    /// <typeparam name="TItem">工作表行实体类型。</typeparam>
    /// <param name="plan">当前工作表的映射计划。</param>
    /// <param name="headers">工作表表头。</param>
    /// <param name="request">当前工作表的导入配置。</param>
    /// <returns>已匹配并可写入实体属性的列绑定。</returns>
    internal static List<MiniExcelExcelImporter.ColumnBinding> BuildBindings<TItem>(IExcelMappingPlan plan,
        string[] headers, ExcelSheetImportRequest request) where TItem : class, new()
    {
        var bindings = new List<MiniExcelExcelImporter.ColumnBinding>();
        foreach (var column in plan.Columns.Where(column => !column.Ignored && !column.IsDynamicColumn))
        {
            var property = typeof(TItem).GetProperty(column.Name,
                BindingFlags.Instance | BindingFlags.Public);
            if (property != null && IsNavigationOrDynamicContainer(property.PropertyType))
                continue;
            var header = FindHeader(headers, column.Title, column.Aliases, request.HeaderComparison,
                request.HeaderWhitespace);
            if (header == null)
            {
                if (request.RequireExpectedHeaders)
                    throw new MiniExcelExcelImporter.MiniExcelSheetException(
                        $"Sheet {request.Name} 缺少表头: {column.Title}");
                continue;
            }
            if (property == null || !property.CanWrite)
            {
                throw new MiniExcelExcelImporter.MiniExcelSheetException($"属性不可写入: {column.Name}");
            }
            var physicalColumnIndex = FindPhysicalColumnIndex(headers, header,
                request.ReadColumnRange?.StartIndex ?? 0);
            bindings.Add(new MiniExcelExcelImporter.ColumnBinding(column, property, header,
                physicalColumnIndex, MiniExcelRowMaterializer.CreatePropertySetter(property)));
        }
        return bindings;
    }

    /// <summary>
    /// 判断属性是否为关系集合或动态字典容器。
    /// </summary>
    /// <param name="propertyType">待判断的属性类型。</param>
    /// <returns>属性类型是关系集合或动态字典容器时返回 true，否则返回 false。</returns>
    internal static bool IsNavigationOrDynamicContainer(Type propertyType) =>
        typeof(IDictionary<string, object>).IsAssignableFrom(propertyType)
        || (propertyType != typeof(string) && typeof(IEnumerable).IsAssignableFrom(propertyType));

    /// <summary>
    /// 校验表头中是否存在未声明的动态列。
    /// </summary>
    /// <param name="headers">已读取的表头集合。</param>
    /// <param name="bindings">已建立的固定列绑定。</param>
    /// <param name="plan">当前工作表的映射计划。</param>
    /// <param name="request">当前工作表的导入配置。</param>
    internal static void ValidateUnknownHeaders(string[] headers,
        IReadOnlyList<MiniExcelExcelImporter.ColumnBinding> bindings, IExcelMappingPlan plan,
        ExcelSheetImportRequest request)
    {
        if (!request.FailOnUnknownDynamicColumns)
            return;
        var known = new HashSet<string>(bindings.Select(binding => binding.Header),
            request.HeaderComparison == ExcelNameComparison.Ordinal
                ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase);
        foreach (var column in plan.DynamicColumns)
        {
            foreach (var header in headers)
            {
                if (FindHeader(new[] { header }, column.Title, column.Aliases,
                        request.HeaderComparison, request.HeaderWhitespace) != null)
                    known.Add(header);
            }
        }
        var unknown = headers.FirstOrDefault(header => !known.Contains(header));
        if (unknown != null)
            throw new MiniExcelExcelImporter.MiniExcelSheetException(
                $"Sheet {request.Name} 包含未声明动态列: {unknown}");
    }

    /// <summary>
    /// 查找表头对应的一基物理列号。
    /// </summary>
    /// <param name="headers">工作表表头。</param>
    /// <param name="header">待定位的表头文本。</param>
    /// <param name="startIndex">读取区域的零基起始列号。</param>
    /// <returns>表头对应的一基物理列号。</returns>
    internal static int FindPhysicalColumnIndex(IReadOnlyList<string> headers, string header, int startIndex)
    {
        for (var index = 0; index < headers.Count; index++)
        {
            if (string.Equals(headers[index], header, StringComparison.Ordinal))
                return startIndex + index + 1;
        }

        throw new MiniExcelExcelImporter.MiniExcelSheetException($"动态列表头无法定位物理列: {header}");
    }

    /// <summary>
    /// 按标题、别名和空白规则查找匹配表头。
    /// </summary>
    /// <param name="headers">待搜索的表头集合。</param>
    /// <param name="title">主标题。</param>
    /// <param name="aliases">可接受的别名集合。</param>
    /// <param name="comparison">名称比较方式。</param>
    /// <param name="whitespace">表头空白处理方式。</param>
    /// <returns>匹配到的原始表头；未匹配时返回 <see langword="null"/>。</returns>
    internal static string FindHeader(IEnumerable<string> headers, string title, IReadOnlyList<string> aliases,
        ExcelNameComparison comparison, ExcelWhitespacePolicy whitespace)
    {
        var expected = new[] { title }.Concat(aliases ?? Array.Empty<string>());
        foreach (var header in headers)
        {
            var normalized = Normalize(header, whitespace);
            foreach (var candidate in expected)
            {
                if (string.Equals(normalized, Normalize(candidate, whitespace),
                    comparison == ExcelNameComparison.Ordinal ? StringComparison.Ordinal
                        : StringComparison.OrdinalIgnoreCase))
                    return header;
            }
        }
        return null;
    }

    /// <summary>
    /// 按指定规则规范化文本空白。
    /// </summary>
    /// <param name="value">待规范化的文本。</param>
    /// <param name="policy">空白处理规则。</param>
    /// <returns>按规则处理后的文本。</returns>
    internal static string Normalize(string value, ExcelWhitespacePolicy policy)
    {
        value ??= string.Empty;
        return policy switch
        {
            ExcelWhitespacePolicy.Preserve => value,
            ExcelWhitespacePolicy.Trim => value.Trim(),
            ExcelWhitespacePolicy.RemoveAll => new string(value.Where(character => !char.IsWhiteSpace(character)).ToArray()),
            _ => throw new ArgumentOutOfRangeException(nameof(policy))
        };
    }

    /// <summary>
    /// 根据读取区域计算 MiniExcel 查询起始单元格。
    /// </summary>
    /// <param name="request">当前工作表的导入配置。</param>
    /// <returns>MiniExcel 查询使用的 A1 起始单元格地址。</returns>
    internal static string CreateStartCell(ExcelSheetImportRequest request)
    {
        if (request.HeaderRowIndex == 0 && request.ReadColumnRange == null)
            return "A1";
        var column = request.ReadColumnRange?.StartIndex ?? 0;
        var letters = string.Empty;
        do
        {
            letters = (char)('A' + column % 26) + letters;
            column = column / 26 - 1;
        } while (column >= 0);
        return letters + (request.HeaderRowIndex + 1).ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 校验表头列数不超过工作表读取限制。
    /// </summary>
    /// <param name="count">已读取的表头列数。</param>
    /// <param name="request">当前工作表的导入配置。</param>
    internal static void ValidateHeaderCount(int count, ExcelSheetImportRequest request)
    {
        if (count > request.MaxReadColumns)
            throw new MiniExcelExcelImporter.MiniExcelSheetException(
                $"Sheet {request.Name} 的表头列数超过限制: {request.MaxReadColumns}");
    }
}
