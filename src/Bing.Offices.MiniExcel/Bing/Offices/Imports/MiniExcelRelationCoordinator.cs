using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.Exceptions;

namespace Bing.Offices.Imports;

/// <summary>
/// 负责在工作簿实体之间执行关系绑定。
/// </summary>
/// <remarks>
/// 该协调器独立于工作表读取和行物化，保持父键索引、首个父项优先和取消语义。
/// </remarks>
internal static class MiniExcelRelationCoordinator
{
    /// <summary>
    /// 表示按具体工作簿类型执行关系绑定的委托。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="root">已导入的工作簿根实体。</param>
    /// <param name="relation">待执行的关系绑定请求。</param>
    /// <param name="errors">接收关系绑定错误的集合。</param>
    /// <param name="request">当前工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消关系绑定的令牌。</param>
    private delegate void RelationBindInvoker<TWorkbook>(TWorkbook root, ExcelRelationRequest relation,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> request,
        CancellationToken cancellationToken) where TWorkbook : class, new();

    /// <summary>
    /// 缓存按工作簿、父子实体和键类型构造的关系绑定委托。
    /// </summary>
    private static readonly ConcurrentDictionary<(Type Workbook, Type Parent, Type Child, Type Key), Delegate>
        RelationBindInvokerCache = new();

    /// <summary>
    /// 按请求将导入的父子实体绑定到导航集合。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="root">已导入的工作簿根实体。</param>
    /// <param name="relations">待执行的关系绑定请求。</param>
    /// <param name="errors">接收关系绑定错误的集合。</param>
    /// <param name="request">当前工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消关系绑定的令牌。</param>
    internal static void Bind<TWorkbook>(TWorkbook root, IReadOnlyList<ExcelRelationRequest> relations,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> request,
        CancellationToken cancellationToken) where TWorkbook : class, new()
    {
        foreach (var relation in relations ?? Array.Empty<ExcelRelationRequest>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                var key = (typeof(TWorkbook), relation.ParentType, relation.ChildType,
                    relation.ParentKey.Method.ReturnType);
                var invoker = (RelationBindInvoker<TWorkbook>)RelationBindInvokerCache.GetOrAdd(
                    key, CreateRelationBindInvoker);
                invoker(root, relation, errors, request, cancellationToken);
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                AddError(errors, request, new ExcelImportError(ExcelImportErrorCode.Relationship,
                    exception.Message, null, 0, 0, null));
            }
        }
    }

    /// <summary>
    /// 为指定的运行时类型组合创建关系绑定委托。
    /// </summary>
    /// <param name="key">工作簿、父实体、子实体和关系键类型组合。</param>
    /// <returns>适用于指定工作簿类型的绑定委托。</returns>
    private static Delegate CreateRelationBindInvoker(
        (Type Workbook, Type Parent, Type Child, Type Key) key)
    {
        var method = typeof(MiniExcelRelationCoordinator).GetMethod(nameof(CreateTypedRelationBindInvoker),
            BindingFlags.Static | BindingFlags.NonPublic)!.MakeGenericMethod(
                key.Workbook, key.Parent, key.Child, key.Key);
        return (Delegate)method.Invoke(null, null)!;
    }

    /// <summary>
    /// 创建调用强类型关系绑定实现的泛型委托。
    /// </summary>
    /// <typeparam name="TWorkbook">工作簿根实体类型。</typeparam>
    /// <typeparam name="TParent">父实体类型。</typeparam>
    /// <typeparam name="TChild">子实体类型。</typeparam>
    /// <typeparam name="TKey">关系键类型。</typeparam>
    /// <returns>绑定指定实体和键类型的委托。</returns>
    private static RelationBindInvoker<TWorkbook> CreateTypedRelationBindInvoker<TWorkbook, TParent, TChild, TKey>()
        where TWorkbook : class, new()
        where TParent : class
        where TChild : class =>
        (root, relation, errors, request, cancellationToken) =>
            BindTypedRelations<TWorkbook, TParent, TChild, TKey>(root, relation, errors, request,
                cancellationToken);

    /// <summary>
    /// 按具体实体和键类型执行关系绑定。
    /// </summary>
    /// <typeparam name="TWorkbook">工作簿根实体类型。</typeparam>
    /// <typeparam name="TParent">父实体类型。</typeparam>
    /// <typeparam name="TChild">子实体类型。</typeparam>
    /// <typeparam name="TKey">关系键类型。</typeparam>
    /// <param name="root">已导入的工作簿根实体。</param>
    /// <param name="relation">当前关系绑定请求。</param>
    /// <param name="errors">接收关系绑定错误的集合。</param>
    /// <param name="request">当前工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消关系绑定的令牌。</param>
    private static void BindTypedRelations<TWorkbook, TParent, TChild, TKey>(TWorkbook root,
        ExcelRelationRequest relation, ICollection<ExcelImportError> errors,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken)
        where TWorkbook : class, new()
        where TParent : class
        where TChild : class
    {
        var parents = (relation.Parents(root) as IEnumerable)?.Cast<TParent>().ToArray()
            ?? Array.Empty<TParent>();
        var children = (relation.Children(root) as IEnumerable)?.Cast<TChild>().ToArray()
            ?? Array.Empty<TChild>();
        var comparer = relation.Comparer as IEqualityComparer<TKey> ?? EqualityComparer<TKey>.Default;
        var parentByKey = new Dictionary<TKey, TParent>(comparer);
        var nullParent = default(TParent);
        var hasNullParent = false;
        Exception parentKeyBoundary = null;

        foreach (var child in children)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var key = InvokeRelationKey<TChild, TKey>(relation.ChildKey, child);
            var indexBuilt = false;
            if (parentByKey.Count == 0 && !hasNullParent && parentKeyBoundary == null)
            {
                indexBuilt = true;
                foreach (var candidate in parents)
                {
                    try
                    {
                        var candidateKey = InvokeRelationKey<TParent, TKey>(relation.ParentKey, candidate);
                        if (candidateKey is null)
                        {
                            if (!hasNullParent)
                            {
                                nullParent = candidate;
                                hasNullParent = true;
                            }
                        }
                        else
                        {
                            parentByKey.TryAdd(candidateKey, candidate);
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception exception) when (exception is not OutOfMemoryException
                        && exception is not StackOverflowException)
                    {
                        // 后续父项异常只有在请求的 child key 无法由此前父项满足时才暴露，
                        // 以保留 First Parent Wins 的短路边界。
                        parentKeyBoundary ??= exception;
                    }
                }
            }

            TParent parent = null;
            var cacheHit = false;
            if (key is null && hasNullParent)
            {
                cacheHit = true;
                var cachedKey = indexBuilt ? key : InvokeRelationKey<TParent, TKey>(relation.ParentKey, nullParent);
                if (indexBuilt || RelationKeysEqual(cachedKey, key, relation.Comparer))
                    parent = nullParent;
            }
            else if (key is not null && parentByKey.TryGetValue(key, out var cachedParent))
            {
                cacheHit = true;
                var cachedKey = indexBuilt ? key : InvokeRelationKey<TParent, TKey>(relation.ParentKey, cachedParent);
                if (indexBuilt || RelationKeysEqual(cachedKey, key, relation.Comparer))
                    parent = cachedParent;
            }

            if (parent == null && cacheHit && parentKeyBoundary == null)
            {
                foreach (var candidate in parents)
                {
                    var candidateKey = InvokeRelationKey<TParent, TKey>(relation.ParentKey, candidate);
                    if (RelationKeysEqual(candidateKey, key, relation.Comparer))
                    {
                        parent = candidate;
                        break;
                    }
                }
            }

            if (parent == null && parentKeyBoundary != null)
                throw parentKeyBoundary;

            if (parent == null)
            {
                AddError(errors, request, new ExcelImportError(ExcelImportErrorCode.Relationship,
                    $"未找到关联父实体: {key}", null, 0, 0, null, rawValue: key));
                continue;
            }
            var navigation = relation.Navigation(parent);
            if (navigation == null)
                continue;
            AddToCollection<TChild>(navigation, child);
        }
    }

    /// <summary>
    /// 按公开的 ICollection 合同追加关系子项，兼容 HashSet 等非 IList 集合。
    /// </summary>
    /// <typeparam name="TItem">关系子实体类型。</typeparam>
    /// <param name="collection">目标导航集合。</param>
    /// <param name="item">待追加的关系子实体。</param>
    private static void AddToCollection<TItem>(object collection, TItem item)
        where TItem : class
    {
        if (collection is ICollection<TItem> typedCollection)
        {
            typedCollection.Add(item);
            return;
        }
        throw new BingOfficesConfigurationException(
            $"关系导航集合 {collection.GetType().FullName} 未实现 ICollection<{typeof(TItem).FullName}>。",
            stage: BingOfficesStage.Plan);
    }

    /// <summary>
    /// 调用关系键选择器并保留原始键选择异常边界。
    /// </summary>
    /// <typeparam name="TItem">参与关系绑定的实体类型。</typeparam>
    /// <typeparam name="TKey">关系键类型。</typeparam>
    /// <param name="keySelector">关系键选择器委托。</param>
    /// <param name="item">待读取关系键的实体。</param>
    /// <returns>实体对应的关系键。</returns>
    private static TKey InvokeRelationKey<TItem, TKey>(Delegate keySelector, TItem item)
        where TItem : class
    {
        try
        {
            return ((Func<TItem, TKey>)keySelector)(item);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            // DynamicInvoke previously exposed TargetInvocationException to the public error boundary.
            throw new TargetInvocationException(exception);
        }
    }

    /// <summary>
    /// 按指定比较器比较父子关系键。
    /// </summary>
    /// <param name="left">左侧关系键。</param>
    /// <param name="right">右侧关系键。</param>
    /// <param name="comparer">非泛型、泛型或反射比较器。</param>
    /// <returns>关系键相等时返回 true，否则返回 false。</returns>
    private static bool RelationKeysEqual(object left, object right, object comparer)
    {
        if (comparer is System.Collections.IEqualityComparer nonGeneric)
            return nonGeneric.Equals(left, right);
        if (comparer is IEqualityComparer<object> objectComparer)
            return objectComparer.Equals(left, right);
        if (comparer != null)
        {
            var leftType = left?.GetType() ?? right?.GetType();
            if (leftType != null)
            {
                var equals = comparer.GetType().GetMethod(nameof(object.Equals),
                    BindingFlags.Instance | BindingFlags.Public, binder: null,
                    types: new[] { leftType, leftType }, modifiers: null);
                if (equals != null)
                    return equals.Invoke(comparer, new[] { left, right }) is true;
            }
        }
        return Equals(left, right);
    }

    /// <summary>
    /// 在未达到错误上限时追加一个关系绑定错误。
    /// </summary>
    /// <typeparam name="TWorkbook">工作簿根实体类型。</typeparam>
    /// <param name="errors">接收错误的集合。</param>
    /// <param name="request">当前工作簿导入请求。</param>
    /// <param name="error">待追加的错误。</param>
    /// <returns>成功追加时返回 true；达到错误上限时返回 false。</returns>
    private static bool AddError<TWorkbook>(ICollection<ExcelImportError> errors,
        ExcelWorkbookImportRequest<TWorkbook> request, ExcelImportError error)
        where TWorkbook : class, new()
    {
        if (request.ResourceLimits?.MaxErrors is int maximum && errors.Count >= maximum)
            return false;
        errors.Add(error);
        return true;
    }
}
