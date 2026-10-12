using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.XSSF.UserModel;
using NPOI.OpenXml4Net.OPC;

namespace Bing.Offices.Npoi.Exports.PivotTables;

/// <summary>
/// 为 NPOI 透视表建立类型化共享项和完整记录快照。
/// </summary>
internal static class NpoiPivotCacheWriter
{
    /// <summary>
    /// 缓存定义指向完整记录部件的工作簿关系类型。
    /// </summary>
    private const string RecordsRelationshipType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";

    /// <summary>
    /// 完成透视缓存定义并生成记录快照。
    /// </summary>
    /// <param name="pivot">已建立缓存及包关系的原生透视表。</param>
    /// <param name="headers">保持来源顺序的字段标题。</param>
    /// <param name="rows">保持来源顺序的字面值记录，不包含表头。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>待写入实际缓存记录部件的完整索引快照。</returns>
    /// <remarks>
    /// 来源值须已经归一为字符串、有限 double、布尔、DateTime 或空值。
    /// 筛选不改变缓存来源；本方法不修改来源工作表或原生透视字段角色。
    /// </remarks>
    internal static NpoiPivotCacheRecords Complete(XSSFPivotTable pivot, IReadOnlyList<string> headers,
        IReadOnlyList<IReadOnlyList<object>> rows, CancellationToken token)
    {
        if (pivot == null)
            throw new ArgumentNullException(nameof(pivot));
        if (headers == null)
            throw new ArgumentNullException(nameof(headers));
        if (rows == null)
            throw new ArgumentNullException(nameof(rows));
        token.ThrowIfCancellationRequested();
        if (headers.Count == 0 || headers.Any(string.IsNullOrWhiteSpace)
            || headers.Distinct(StringComparer.OrdinalIgnoreCase).Count() != headers.Count)
            throw new ArgumentException("缓存字段标题必须非空且唯一。", nameof(headers));
        if (rows.Count == 0)
            throw new ArgumentException("透视缓存至少需要一条来源记录。", nameof(rows));
        var definition = pivot.GetPivotCacheDefinition()?.GetCTPivotCacheDefinition()
            ?? throw new InvalidOperationException("原生透视表没有缓存定义。");
        var partName = GetRecordPartName(pivot, definition);
        var fields = new CT_CacheFields { cacheField = new List<CT_CacheField>() };
        var recordIndexes = new uint[rows.Count][];
        for (var row = 0; row < rows.Count; row++)
        {
            token.ThrowIfCancellationRequested();
            if (rows[row] == null || rows[row].Count != headers.Count)
                throw new ArgumentException("缓存来源行必须包含全部来源字段。", nameof(rows));
            recordIndexes[row] = new uint[headers.Count];
        }
        for (var column = 0; column < headers.Count; column++)
        {
            token.ThrowIfCancellationRequested();
            var items = new CT_SharedItems { Items = new List<object>() };
            var values = new List<object>();
            var indexes = new Dictionary<object, uint>();
            uint? blankIndex = null;
            for (var row = 0; row < rows.Count; row++)
            {
                token.ThrowIfCancellationRequested();
                var value = rows[row][column];
                uint index;
                if (value == null)
                {
                    if (!blankIndex.HasValue)
                    {
                        blankIndex = checked((uint)items.Items.Count);
                        items.Items.Add(new CT_Missing());
                        values.Add(null);
                    }
                    index = blankIndex.Value;
                }
                else if (!indexes.TryGetValue(value, out index))
                {
                    var item = ToCacheItem(value);
                    index = checked((uint)items.Items.Count);
                    items.Items.Add(item);
                    values.Add(value);
                    indexes.Add(value, index);
                }
                recordIndexes[row][column] = index;
            }
            SetSharedItemFlags(items, values);
            items.count = checked((uint)items.Items.Count);
            items.countSpecified = true;
            fields.cacheField.Add(new CT_CacheField { name = NpoiPivotTextEncoding.Encode(headers[column]), sharedItems = items });
        }
        fields.count = checked((uint)fields.cacheField.Count);
        fields.countSpecified = true;
        var counts = fields.cacheField.Select(field => field.sharedItems.Items.Count).ToArray();
        var records = new NpoiPivotCacheRecords(partName, counts,
            recordIndexes.Select(row => (IReadOnlyList<uint>)row).ToArray());
        var fieldItems = CreateFieldItems(pivot, counts, token);
        token.ThrowIfCancellationRequested();
        // 在完整构建和校验后替换缓存，失败或取消不留下半份共享项。
        definition.cacheFields = fields;
        definition.recordCount = checked((uint)rows.Count);
        definition.recordCountSpecified = true;
        definition.saveData = true;
        foreach (var replacement in fieldItems)
            replacement.Field.items = replacement.Items;
        return records;
    }

    /// <summary>
    /// 从缓存定义的包关系取得实际记录部件地址。
    /// </summary>
    /// <param name="pivot">已建立包关系的原生透视表。</param>
    /// <param name="definition">包含记录关系标识的缓存定义。</param>
    /// <returns>待写入的记录部件绝对包地址。</returns>
    private static string GetRecordPartName(XSSFPivotTable pivot, CT_PivotCacheDefinition definition)
    {
        var part = pivot.GetPivotCacheDefinition().GetPackagePart();
        if (part == null || string.IsNullOrWhiteSpace(definition.id))
            throw new InvalidOperationException("原生透视表没有缓存记录关系。");
        var relationship = part.GetRelationship(definition.id);
        if (relationship == null || relationship.RelationshipType != RecordsRelationshipType
            || relationship.TargetMode != TargetMode.Internal)
            throw new InvalidOperationException("原生透视缓存记录关系无效。");
        // 创建时的 records 对象可能没有 PackagePart；缓存定义关系才是实际包地址。
        return part.GetRelatedPart(relationship).PartName.Name;
    }

    /// <summary>
    /// 为已配置轴字段建立合法的共享项引用。
    /// </summary>
    /// <param name="pivot">当前原生透视表。</param>
    /// <param name="sharedItemCounts">各字段去重后的共享项数量。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>完整构建后待替换的字段项目列表。</returns>
    private static IReadOnlyList<(CT_PivotField Field, CT_Items Items)> CreateFieldItems(
        XSSFPivotTable pivot, IReadOnlyList<int> sharedItemCounts, CancellationToken token)
    {
        var table = pivot.GetCTPivotTableDefinition();
        var fields = table.pivotFields?.pivotField;
        if (fields == null || fields.Count != sharedItemCounts.Count)
            throw new ArgumentException("原生透视字段数量与缓存来源不一致。", nameof(pivot));
        var replacements = new List<(CT_PivotField Field, CT_Items Items)>();
        for (var column = 0; column < fields.Count; column++)
        {
            token.ThrowIfCancellationRequested();
            var field = fields[column];
            if (!field.axis.HasValue || field.axis == ST_Axis.axisValues)
                continue;
            var items = new CT_Items { item = new List<CT_Item>() };
            if (field.axis.HasValue && field.axis != ST_Axis.axisValues)
            {
                for (var index = 0; index < sharedItemCounts[column]; index++)
                {
                    token.ThrowIfCancellationRequested();
                    items.item.Add(new CT_Item { t = ST_ItemType.data, x = checked((uint)index), xSpecified = true });
                }
                if (field.axis != ST_Axis.axisPage && field.defaultSubtotal)
                    items.item.Add(new CT_Item { t = ST_ItemType.@default });
            }
            items.count = checked((uint)items.item.Count);
            replacements.Add((field, items));
        }
        return replacements;
    }

    /// <summary>
    /// 将来源字面值转换为类型化缓存共享项。
    /// </summary>
    /// <param name="value">非空的来源值。</param>
    /// <returns>用于缓存定义的共享项。</returns>
    private static object ToCacheItem(object value) => value switch
    {
        string text => new CT_String { v = NpoiPivotTextEncoding.Encode(text) },
        double number when double.IsFinite(number) => new CT_Number { v = number },
        bool boolean => new CT_Boolean { v = boolean },
        DateTime date => new CT_DateTime { v = date },
        _ => throw new ArgumentException("缓存来源仅接受文本、有限 double、布尔、日期或空值。", nameof(value))
    };

    /// <summary>
    /// 根据共享项类型配置字段标志。
    /// </summary>
    /// <param name="items">完整字段共享项。</param>
    /// <param name="values">尚未转义的原始共享项值。</param>
    private static void SetSharedItemFlags(CT_SharedItems items, IReadOnlyList<object> values)
    {
        var strings = items.Items.OfType<CT_String>().ToArray();
        var numbers = items.Items.OfType<CT_Number>().ToArray();
        var hasBoolean = items.Items.Any(item => item is CT_Boolean);
        var hasDate = items.Items.Any(item => item is CT_DateTime);
        items.containsBlank = items.Items.Any(item => item is CT_Missing);
        items.containsString = strings.Length > 0;
        items.containsSemiMixedTypes = items.containsString;
        items.containsNumber = numbers.Length > 0;
        items.containsDate = hasDate;
        items.containsNonDate = items.containsString || items.containsNumber || hasBoolean;
        items.containsMixedTypes = new[] { items.containsString, items.containsNumber, hasBoolean, hasDate }
            .Count(hasType => hasType) > 1;
        items.containsInteger = items.containsNumber && !items.containsMixedTypes
            && numbers.All(item => Math.Truncate(item.v) == item.v);
        items.longText = values.OfType<string>().Any(text => text.Length > 255);
    }
}
