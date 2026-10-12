using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 在实体摘要导入期间验证父子键关系，而不保留实体集合或导航集合。
/// </summary>
internal sealed class ExcelEntityForwardRelationTracker
{
    private readonly IReadOnlyList<RelationState> _relations;
    private readonly Func<bool> _tryReserveKey;
    private readonly string _provider;

    private ExcelEntityForwardRelationTracker(IReadOnlyList<RelationState> relations,
        Func<bool> tryReserveKey, string provider)
    {
        _relations = relations;
        _tryReserveKey = tryReserveKey ?? throw new ArgumentNullException(nameof(tryReserveKey));
        _provider = provider ?? "Core";
    }

    /// <summary>
    /// 获取关系键预算是否已经耗尽。
    /// </summary>
    internal bool IsTruncated { get; private set; }

    /// <summary>
    /// 验证摘要关系所需的列表顺序和类型唯一性。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="layout">实体布局。</param>
    internal static void ValidateSummaryLayout<TEntity>(ExcelEntityLayout<TEntity> layout)
        where TEntity : class, new()
    {
        _ = Create(layout, () => true, "Core");
    }

    /// <summary>
    /// 创建摘要关系验证器。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="layout">实体布局。</param>
    /// <param name="tryReserveKey">尝试预留父键索引预算的委托。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <returns>摘要关系验证器。</returns>
    internal static ExcelEntityForwardRelationTracker Create<TEntity>(
        ExcelEntityLayout<TEntity> layout, Func<bool> tryReserveKey, string provider)
        where TEntity : class, new()
    {
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));
        if (layout.Relations == null || layout.Relations.Count == 0)
            return new ExcelEntityForwardRelationTracker(Array.Empty<RelationState>(),
                tryReserveKey ?? (() => true), provider);

        var regions = layout.ListRegions.Select((region, index) =>
            new RegionState(region.ItemType, index)).ToArray();
        var states = new List<RelationState>(layout.Relations.Count);
        foreach (var relation in layout.Relations)
        {
            var parents = regions.Where(region => region.ItemType == relation.ParentType).ToArray();
            var children = regions.Where(region => region.ItemType == relation.ChildType).ToArray();
            if (parents.Length != 1 || children.Length != 1)
                throw new BingOfficesConfigurationException(
                    $"摘要关系必须分别对应唯一的父列表和子列表: {relation.ParentType.Name} -> {relation.ChildType.Name}。",
                    stage: BingOfficesStage.Plan);
            if (relation.ParentType == relation.ChildType || parents[0].Index >= children[0].Index)
                throw new BingOfficesConfigurationException(
                    "摘要关系要求父列表位于子列表之前，且父子类型不能相同。",
                    stage: BingOfficesStage.Plan);
            states.Add(new RelationState(relation));
        }
        return new ExcelEntityForwardRelationTracker(states, tryReserveKey ?? (() => true), provider);
    }

    /// <summary>
    /// 记录一个已通过字段校验的列表项并返回关系错误。
    /// </summary>
    /// <param name="itemType">列表项类型。</param>
    /// <param name="item">已转换的列表项。</param>
    /// <param name="sheetName">来源工作表名称。</param>
    /// <param name="rowIndex">来源行的零基索引。</param>
    /// <returns>当前列表项产生的结构化关系错误。</returns>
    internal IReadOnlyList<ExcelImportError> Track(Type itemType, object item,
        string sheetName, int rowIndex)
    {
        if (itemType == null || item == null || _relations.Count == 0)
            return Array.Empty<ExcelImportError>();

        List<ExcelImportError> errors = null;
        foreach (var relation in _relations)
        {
            if (relation.ParentType == itemType)
                TrackParent(relation, item, sheetName, rowIndex, ref errors);
            else if (relation.ChildType == itemType)
                TrackChild(relation, item, sheetName, rowIndex, ref errors);
        }
        return errors ?? (IReadOnlyList<ExcelImportError>)Array.Empty<ExcelImportError>();
    }

    private void TrackParent(RelationState relation, object item, string sheetName,
        int rowIndex, ref List<ExcelImportError> errors)
    {
        var key = InvokeKey(relation.Request.ParentKey, item, "父项键选择器");
        if (key == null)
        {
            (errors ??= new List<ExcelImportError>()).Add(CreateError(
                ExcelImportErrorCode.Relationship, "摘要关系父键为空。",
                relation.ParentType, sheetName, rowIndex, key));
            return;
        }
        if (relation.ParentKeys.Contains(key))
        {
            (errors ??= new List<ExcelImportError>()).Add(CreateError(
                ExcelImportErrorCode.Relationship, "摘要关系父键重复。",
                relation.ParentType, sheetName, rowIndex, key));
            return;
        }
        if (!_tryReserveKey())
        {
            IsTruncated = true;
            (errors ??= new List<ExcelImportError>()).Add(CreateError(
                ExcelImportErrorCode.ResourceLimit,
                "摘要关系父键索引数量超过限制。", relation.ParentType, sheetName, rowIndex, key));
            return;
        }
        relation.ParentKeys.Add(key);
    }

    private void TrackChild(RelationState relation, object item, string sheetName,
        int rowIndex, ref List<ExcelImportError> errors)
    {
        var key = InvokeKey(relation.Request.ChildKey, item, "子项键选择器");
        if (key == null)
        {
            (errors ??= new List<ExcelImportError>()).Add(CreateError(
                ExcelImportErrorCode.Relationship, "摘要关系子键为空。",
                relation.ChildType, sheetName, rowIndex, key));
            return;
        }
        if (!relation.ParentKeys.Contains(key))
            (errors ??= new List<ExcelImportError>()).Add(CreateError(
                ExcelImportErrorCode.Relationship, "摘要关系无法匹配父项。",
                relation.ChildType, sheetName, rowIndex, key));
    }

    private object InvokeKey(Delegate selector, object item, string name)
    {
        try
        {
            return selector.DynamicInvoke(item);
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            if (exception.InnerException is OperationCanceledException)
                ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw new BingOfficesImportException($"Excel {name}执行失败。", exception.InnerException,
                _provider, BingOfficesStage.Validate, code: BingOfficesErrorCode.UserExtensionFailed);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesImportException($"Excel {name}执行失败。", exception,
                _provider, BingOfficesStage.Validate, code: BingOfficesErrorCode.UserExtensionFailed);
        }
    }

    private static ExcelImportError CreateError(ExcelImportErrorCode code, string message,
        Type itemType, string sheetName, int rowIndex, object rawValue) =>
        new(code, message, sheetName, rowIndex + 1, 0, itemType.Name, itemType.Name,
            rawValue: rawValue);

    private sealed class RelationState
    {
        internal RelationState(ExcelRelationRequest request)
        {
            Request = request;
            ParentType = request.ParentType;
            ChildType = request.ChildType;
            ParentKeys = new HashSet<object>(CreateComparer(request.Comparer));
        }

        internal ExcelRelationRequest Request { get; }
        internal Type ParentType { get; }
        internal Type ChildType { get; }
        internal HashSet<object> ParentKeys { get; }
    }

    private readonly struct RegionState
    {
        internal RegionState(Type itemType, int index)
        {
            ItemType = itemType;
            Index = index;
        }

        internal Type ItemType { get; }
        internal int Index { get; }
    }

    private static IEqualityComparer<object> CreateComparer(object comparer)
    {
        if (comparer is IEqualityComparer<object> objectComparer)
            return objectComparer;
        if (comparer == null)
            return EqualityComparer<object>.Default;
        var comparerInterface = comparer.GetType().GetInterfaces()
            .FirstOrDefault(type => type.IsGenericType
                && type.GetGenericTypeDefinition() == typeof(IEqualityComparer<>));
        var equalsMethod = comparerInterface?.GetMethod(nameof(IEqualityComparer<object>.Equals));
        var hashCodeMethod = comparerInterface?.GetMethod(nameof(IEqualityComparer<object>.GetHashCode));
        return equalsMethod == null || hashCodeMethod == null
            ? EqualityComparer<object>.Default
            : new ReflectionComparer(comparer, equalsMethod, hashCodeMethod);
    }

    private sealed class ReflectionComparer : IEqualityComparer<object>
    {
        private readonly object _comparer;
        private readonly MethodInfo _equalsMethod;
        private readonly MethodInfo _hashCodeMethod;

        internal ReflectionComparer(object comparer, MethodInfo equalsMethod, MethodInfo hashCodeMethod)
        {
            _comparer = comparer;
            _equalsMethod = equalsMethod;
            _hashCodeMethod = hashCodeMethod;
        }

        public new bool Equals(object x, object y) =>
            (bool)_equalsMethod.Invoke(_comparer, new[] { x, y });

        public int GetHashCode(object obj) =>
            (int)_hashCodeMethod.Invoke(_comparer, new[] { obj });
    }
}
