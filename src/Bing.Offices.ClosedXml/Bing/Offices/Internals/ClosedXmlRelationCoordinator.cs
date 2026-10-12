using System.Collections;
using System.Reflection;
using Bing.Offices.Imports;

namespace Bing.Offices.ClosedXml.Internals;

/// <summary>
/// 执行公共 Workbook/Entity 的父子集合关系绑定。
/// </summary>
internal static class ClosedXmlRelationCoordinator
{
    /// <summary>
    /// 按公共关系请求将子集合绑定到父实体。
    /// </summary>
    /// <param name="root">包含父集合和子集合的根对象。</param>
    /// <param name="relations">待处理的关系请求。</param>
    /// <param name="errors">用于收集关系错误的集合。</param>
    /// <param name="maxErrors">允许记录的最大错误数；为空表示不限制。</param>
    /// <param name="relationKeyBudget">关系父键索引的剩余预算。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="sourceLocations">实体到源工作表位置的引用映射。</param>
    /// <param name="sourceSheets">按列表集合索引的导入结果。</param>
    /// <returns>达到错误数量上限时返回 <see langword="true" />。</returns>
    internal static bool Bind(object root, IReadOnlyList<ExcelRelationRequest> relations,
        ICollection<ExcelImportError> errors, int? maxErrors,
        ClosedXmlRelationKeyBudget relationKeyBudget, CancellationToken cancellationToken,
        IReadOnlyDictionary<object, ClosedXmlSourceLocation> sourceLocations = null,
        IReadOnlyDictionary<object, ExcelSheetImportResult> sourceSheets = null)
    {
        if (root == null || relations == null || relations.Count == 0)
            return false;
        relationKeyBudget ??= new ClosedXmlRelationKeyBudget(null);
        var truncated = false;
        foreach (var relation in relations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var parentCollection = relation.Parents(root);
            var parents = (parentCollection as IEnumerable)?.Cast<object>()
                ?? Enumerable.Empty<object>();
            var childCollection = relation.Children(root);
            var children = (childCollection as IEnumerable)?.Cast<object>()
                ?? Enumerable.Empty<object>();
            var parentIndex = new RelationKeyIndex(relation);
            var parentRowIndex = 0;
            foreach (var parent in parents)
            {
                var sourceLocation = GetSourceLocation(sourceLocations, sourceSheets,
                    parentCollection, parent, parentRowIndex++);
                cancellationToken.ThrowIfCancellationRequested();
                if (Reached(errors, maxErrors))
                {
                    truncated = true;
                    break;
                }
                var parentKey = relation.ParentKey.DynamicInvoke(parent);
                if (parentKey == null)
                {
                    AddError(errors, maxErrors, CreateError(ExcelImportErrorCode.Relationship,
                        "父项关联键为空。", relation.ParentType, sourceLocation));
                    continue;
                }
                if (parentIndex.Contains(parentKey))
                {
                    AddError(errors, maxErrors, CreateError(ExcelImportErrorCode.Relationship,
                        "父项关联键重复。", relation.ParentType, sourceLocation,
                        parentKey));
                    continue;
                }
                if (!relationKeyBudget.TryReserve())
                {
                    AddError(errors, maxErrors, CreateError(ExcelImportErrorCode.ResourceLimit,
                        "关系父键索引数量超过限制。", relation.ParentType, sourceLocation));
                    truncated = true;
                    break;
                }
                parentIndex.TryAdd(parentKey, parent);
            }
            if (truncated)
                return true;
            var childRowIndex = 0;
            foreach (var child in children)
            {
                var sourceLocation = GetSourceLocation(sourceLocations, sourceSheets,
                    childCollection, child, childRowIndex++);
                cancellationToken.ThrowIfCancellationRequested();
                if (Reached(errors, maxErrors))
                {
                    truncated = true;
                    break;
                }
                var childKey = relation.ChildKey.DynamicInvoke(child);
                if (childKey == null)
                {
                    AddError(errors, maxErrors, CreateError(ExcelImportErrorCode.Relationship,
                        "子项关联键为空。", relation.ChildType, sourceLocation));
                    continue;
                }
                if (!parentIndex.TryGetValue(childKey, out var parent) || parent == null)
                {
                    AddError(errors, maxErrors, CreateError(ExcelImportErrorCode.Relationship,
                        "无法绑定父子关系。", relation.ChildType, sourceLocation,
                        childKey));
                    continue;
                }
                var navigation = relation.Navigation(parent);
                if (!TryAdd(navigation, relation.ChildType, child))
                {
                    AddError(errors, maxErrors, CreateError(ExcelImportErrorCode.Relationship,
                        "父实体子集合导航属性不可写。", relation.ChildType, sourceLocation,
                        childKey));
                    continue;
                }
            }
        }
        return truncated || maxErrors.HasValue && errors.Count >= maxErrors.Value;
    }

    /// <summary>
    /// 尝试向导航集合加入子实体。
    /// </summary>
    /// <param name="navigation">父实体上的导航集合。</param>
    /// <param name="childType">子实体类型。</param>
    /// <param name="child">待加入的子实体。</param>
    /// <returns>导航集合支持写入并已加入时返回 <see langword="true" />。</returns>
    private static bool TryAdd(object navigation, Type childType, object child)
    {
        if (navigation is IList list)
        {
            list.Add(child);
            return true;
        }
        if (navigation == null)
            return false;
        var add = navigation.GetType().GetMethod("Add", new[] { childType });
        if (add == null)
        {
            add = navigation.GetType().GetInterfaces()
                .Where(type => type.IsGenericType
                    && type.GetGenericTypeDefinition() == typeof(ICollection<>))
                .Select(type => type.GetMethod("Add"))
                .FirstOrDefault(method => method != null);
        }
        if (add == null)
            return false;
        add.Invoke(navigation, new[] { child });
        return true;
    }

    /// <summary>
    /// 维护关系键到父实体的索引，并复用请求配置的比较器。
    /// </summary>
    private sealed class RelationKeyIndex
    {
        /// <summary>
        /// 关系键到父实体的映射。
        /// </summary>
        private readonly Dictionary<object, object> _parents;

        /// <summary>
        /// 初始化关系键索引。
        /// </summary>
        /// <param name="relation">关系请求。</param>
        public RelationKeyIndex(ExcelRelationRequest relation)
        {
            _parents = new Dictionary<object, object>(CreateComparer(relation));
        }

        /// <summary>
        /// 尝试加入一个父键。
        /// </summary>
        /// <param name="key">父项关联键。</param>
        /// <param name="parent">父实体。</param>
        /// <returns>键不存在并成功加入时返回 <see langword="true" />。</returns>
        public bool TryAdd(object key, object parent) => _parents.TryAdd(key, parent);

        /// <summary>
        /// 检查关系键是否已经建立索引。
        /// </summary>
        /// <param name="key">待检查的父项关联键。</param>
        /// <returns>已存在时返回 <see langword="true" />。</returns>
        public bool Contains(object key) => _parents.ContainsKey(key);

        /// <summary>
        /// 按子项关联键查找父实体。
        /// </summary>
        /// <param name="key">子项关联键。</param>
        /// <param name="parent">找到的父实体。</param>
        /// <returns>找到匹配父实体时返回 <see langword="true" />。</returns>
        public bool TryGetValue(object key, out object? parent) => _parents.TryGetValue(key, out parent);

        /// <summary>
        /// 创建关系键的对象比较器适配器。
        /// </summary>
        /// <param name="relation">关系请求。</param>
        /// <returns>与关系配置一致的对象比较器。</returns>
        private static IEqualityComparer<object> CreateComparer(ExcelRelationRequest relation)
        {
            if (relation.Comparer == null)
                return EqualityComparer<object>.Default;
            var comparerInterface = relation.Comparer.GetType().GetInterfaces()
                .FirstOrDefault(type => type.IsGenericType
                    && type.GetGenericTypeDefinition() == typeof(IEqualityComparer<>));
            var equalsMethod = comparerInterface?.GetMethod(nameof(IEqualityComparer<object>.Equals));
            var hashCodeMethod = comparerInterface?.GetMethod(nameof(IEqualityComparer<object>.GetHashCode));
            return equalsMethod == null || hashCodeMethod == null
                ? EqualityComparer<object>.Default
                : new ReflectionRelationComparer(relation.Comparer, equalsMethod, hashCodeMethod);
        }
    }

    /// <summary>
    /// 将关系请求的泛型比较器适配到对象字典。
    /// </summary>
    private sealed class ReflectionRelationComparer : IEqualityComparer<object>
    {
        /// <summary>
        /// 原始关系比较器。
        /// </summary>
        private readonly object _comparer;

        /// <summary>
        /// 泛型比较器的相等方法。
        /// </summary>
        private readonly MethodInfo _equalsMethod;

        /// <summary>
        /// 泛型比较器的哈希方法。
        /// </summary>
        private readonly MethodInfo _hashCodeMethod;

        /// <summary>
        /// 初始化比较器适配器。
        /// </summary>
        /// <param name="comparer">原始关系比较器。</param>
        /// <param name="equalsMethod">泛型相等方法。</param>
        /// <param name="hashCodeMethod">泛型哈希方法。</param>
        public ReflectionRelationComparer(object comparer, MethodInfo equalsMethod, MethodInfo hashCodeMethod)
        {
            _comparer = comparer;
            _equalsMethod = equalsMethod;
            _hashCodeMethod = hashCodeMethod;
        }

        /// <inheritdoc />
        public new bool Equals(object? x, object? y) =>
            (bool)_equalsMethod.Invoke(_comparer, new[] { x, y })!;

        /// <inheritdoc />
        public int GetHashCode(object obj) =>
            (int)_hashCodeMethod.Invoke(_comparer, new[] { obj })!;
    }

    /// <summary>
    /// 在未达到错误上限时记录一个关系错误。
    /// </summary>
    /// <param name="errors">错误集合。</param>
    /// <param name="maxErrors">允许记录的最大错误数；为空表示不限制。</param>
    /// <param name="error">待记录的错误。</param>
    private static void AddError(ICollection<ExcelImportError> errors, int? maxErrors,
        ExcelImportError error)
    {
        if (errors == null || maxErrors.HasValue && errors.Count >= maxErrors.Value)
            return;
        errors.Add(error);
    }

    /// <summary>
    /// 使用实体来源位置创建关系错误。
    /// </summary>
    /// <param name="code">错误代码。</param>
    /// <param name="message">错误消息。</param>
    /// <param name="itemType">产生错误的实体类型。</param>
    /// <param name="sourceLocation">产生错误的实体来源位置。</param>
    /// <param name="rawValue">导致错误的原始关联键。</param>
    /// <returns>包含源工作表和行号的导入错误。</returns>
    private static ExcelImportError CreateError(ExcelImportErrorCode code, string message,
        Type itemType, RelationSourceLocation sourceLocation, object rawValue = null)
    {
        return new ExcelImportError(code, message, sourceLocation.SheetName ?? string.Empty,
            sourceLocation.RowIndex, 0, itemType.Name, rawValue: rawValue);
    }

    private static RelationSourceLocation GetSourceLocation(
        IReadOnlyDictionary<object, ClosedXmlSourceLocation> sourceLocations,
        IReadOnlyDictionary<object, ExcelSheetImportResult> sourceSheets,
        object collection, object item, int index)
    {
        if (sourceLocations != null && item != null
            && sourceLocations.TryGetValue(item, out var location))
            return new RelationSourceLocation(location.SheetName, location.RowIndex);
        if (sourceSheets != null && collection != null && index >= 0
            && sourceSheets.TryGetValue(collection, out var sheet)
            && index < sheet.SourceRows.Count)
            return new RelationSourceLocation(sheet.Name, sheet.SourceRows[index] + 1);
        return default;
    }

    private readonly struct RelationSourceLocation
    {
        internal RelationSourceLocation(string sheetName, int rowIndex)
        {
            SheetName = sheetName;
            RowIndex = rowIndex;
        }

        internal string SheetName { get; }
        internal int RowIndex { get; }
    }

    /// <summary>
    /// 判断错误集合是否已达到配置上限。
    /// </summary>
    /// <param name="errors">错误集合。</param>
    /// <param name="maxErrors">允许记录的最大错误数。</param>
    /// <returns>已达到上限时返回 <see langword="true" />。</returns>
    private static bool Reached(ICollection<ExcelImportError> errors, int? maxErrors) =>
        maxErrors.HasValue && errors != null && errors.Count >= maxErrors.Value;
}
