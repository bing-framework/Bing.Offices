using System.Collections;
using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.ExceptionServices;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.ExcelDataReader;

/// <summary>
/// 负责在工作簿分批导入完成后执行跨工作表关系绑定。
/// </summary>
/// <typeparam name="TWorkbook">Workbook 根模型类型。</typeparam>
/// <remarks>
/// 分批回调仍按物理工作表顺序交付；关系绑定延迟到所有可读取工作表完成后，
/// 因此父子工作表可以乱序出现，且父项失败不会进入关系索引。
/// </remarks>
internal sealed class ExcelDataReaderBatchRelationCoordinator<TWorkbook>
    where TWorkbook : class, new()
{
    /// <summary>
    /// 执行指定父子类型和键类型的关系绑定。
    /// </summary>
    /// <param name="root">工作簿根模型。</param>
    /// <param name="relation">关系定义。</param>
    /// <param name="sheetName">子工作表名称。</param>
    /// <param name="errors">接收关系错误的集合。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private delegate void RelationBindInvoker(
        object root, ExcelRelationRequest relation, string sheetName,
        ICollection<ExcelImportError> errors, CancellationToken cancellationToken);

    /// <summary>
    /// 按父子类型和键类型缓存关系绑定委托。
    /// </summary>
    private static readonly ConcurrentDictionary<(Type Parent, Type Child, Type Key), RelationBindInvoker>
        RelationBindInvokerCache = new();

    /// <summary>
    /// 保存关系绑定期间的工作簿根模型。
    /// </summary>
    private readonly TWorkbook _root = new TWorkbook();
    /// <summary>
    /// 获取工作簿中配置的关系定义。
    /// </summary>
    private readonly IReadOnlyList<ExcelRelationRequest> _relations;
    /// <summary>
    /// 获取工作簿中配置的工作表请求。
    /// </summary>
    private readonly IReadOnlyList<ExcelSheetImportRequest> _sheets;
    /// <summary>
    /// 保存参与关系绑定的目标集合引用。
    /// </summary>
    private readonly List<object> _relationCollections;

    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderBatchRelationCoordinator{TWorkbook}" /> 类型的实例。
    /// </summary>
    /// <param name="request">工作簿导入请求。</param>
    internal ExcelDataReaderBatchRelationCoordinator(ExcelWorkbookImportRequest<TWorkbook> request)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        _relations = request.Relations ?? Array.Empty<ExcelRelationRequest>();
        _sheets = request.Sheets ?? Array.Empty<ExcelSheetImportRequest>();
        _relationCollections = new List<object>();
        foreach (var relation in _relations)
        {
            if (relation == null)
                throw new BingOfficesConfigurationException("关系请求不能为空。", stage: BingOfficesStage.Plan);
            var parents = relation.Parents(_root);
            var children = relation.Children(_root);
            if (parents == null || children == null)
                throw new BingOfficesConfigurationException(
                    $"关系集合不能为空: {relation.ParentType.Name} -> {relation.ChildType.Name}。",
                    stage: BingOfficesStage.Plan);
            AddRelationCollection(parents);
            AddRelationCollection(children);
        }
    }

    /// <summary>
    /// 在批次回调前保存关系涉及的成功实体。
    /// </summary>
    /// <param name="sheet">当前批次对应的 Sheet 请求。</param>
    /// <param name="batch">待交付的 Workbook 批次。</param>
    internal void AcceptBatch(ExcelSheetImportRequest sheet, ExcelWorkbookImportBatch batch)
    {
        if (sheet == null || batch == null || batch.Items.Count == 0)
            return;
        var target = sheet.Target(_root);
        if (!ContainsRelationCollection(target))
            return;
        AddItems(target, batch.Items);
    }

    /// <summary>
    /// 在所有 Sheet 读取完成后绑定关系并按子 Sheet 分组返回错误。
    /// </summary>
    /// <param name="maxErrors">本阶段允许追加的关系错误数量；null 表示不限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>关系错误分组以及是否因错误预算截断。</returns>
    internal ExcelDataReaderRelationBindingResult Complete(int? maxErrors,
        CancellationToken cancellationToken)
    {
        var groups = new List<ExcelDataReaderRelationErrorGroup>();
        var acceptedErrors = 0;
        var truncated = false;
        ExcelSheetImportRequest truncatedSheet = null;
        foreach (var relation in _relations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var childSheet = ResolveChildSheet(relation);
            var relationErrors = new List<ExcelImportError>();
            try
            {
                var invoker = RelationBindInvokerCache.GetOrAdd(
                    (relation.ParentType, relation.ChildType, relation.ParentKey.Method.ReturnType),
                    CreateRelationBindInvoker);
                invoker(_root, relation, childSheet.Name, relationErrors, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception) when (exception is not OutOfMemoryException
                && exception is not StackOverflowException)
            {
                relationErrors.Add(CreateRelationError(childSheet.Name, exception.Message));
            }

            if (relationErrors.Count == 0)
                continue;
            var remaining = maxErrors.HasValue ? maxErrors.Value - acceptedErrors : relationErrors.Count;
            if (remaining <= 0)
            {
                truncated = true;
                truncatedSheet = childSheet;
                break;
            }
            if (relationErrors.Count > remaining)
            {
                relationErrors = relationErrors.Take(remaining).ToList();
                truncated = true;
                truncatedSheet = childSheet;
            }
            acceptedErrors += relationErrors.Count;
            groups.Add(new ExcelDataReaderRelationErrorGroup(childSheet, relationErrors));
            if (truncated)
                break;
        }
        return new ExcelDataReaderRelationBindingResult(groups, truncated, truncatedSheet);
    }

    /// <summary>
    /// 保存尚未记录的关系目标集合引用。
    /// </summary>
    /// <param name="collection">关系目标集合。</param>
    private void AddRelationCollection(object collection)
    {
        if (!_relationCollections.Any(existing => ReferenceEquals(existing, collection)))
            _relationCollections.Add(collection);
    }

    /// <summary>
    /// 判断对象是否是关系目标集合。
    /// </summary>
    /// <param name="target">待判断的对象。</param>
    /// <returns>对象属于关系目标集合时为 <see langword="true" />；否则为 <see langword="false" />。</returns>
    private bool ContainsRelationCollection(object target)
        => target != null && _relationCollections.Any(collection => ReferenceEquals(collection, target));

    /// <summary>
    /// 查找关系对应的子工作表请求。
    /// </summary>
    /// <param name="relation">关系定义。</param>
    /// <returns>匹配的子工作表请求。</returns>
    private ExcelSheetImportRequest ResolveChildSheet(ExcelRelationRequest relation)
    {
        var children = relation.Children(_root);
        var exact = _sheets.FirstOrDefault(sheet => ReferenceEquals(sheet.Target(_root), children));
        return exact ?? _sheets.FirstOrDefault(sheet => sheet.ItemType == relation.ChildType) ?? _sheets[0];
    }

    /// <summary>
    /// 将批次实体追加到关系目标集合。
    /// </summary>
    /// <param name="target">关系目标集合。</param>
    /// <param name="items">待追加的实体。</param>
    private static void AddItems(object target, IReadOnlyList<object> items)
    {
        if (target is IList list)
        {
            foreach (var item in items)
                list.Add(item);
            return;
        }
        var add = target.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public)
            .Where(method => string.Equals(method.Name, "Add", StringComparison.Ordinal)
                && method.GetParameters().Length == 1)
            .FirstOrDefault();
        if (add == null)
            throw new BingOfficesConfigurationException(
                $"关系目标集合 {target.GetType().FullName} 必须支持 Add 操作。",
                stage: BingOfficesStage.Plan);
        foreach (var item in items)
            add.Invoke(target, new[] { item });
    }

    /// <summary>
    /// 创建指定类型组合的关系绑定委托。
    /// </summary>
    /// <param name="key">父类型、子类型和键类型组合。</param>
    /// <returns>关闭泛型参数后的关系绑定委托。</returns>
    private static RelationBindInvoker CreateRelationBindInvoker(
        (Type Parent, Type Child, Type Key) key)
    {
        var method = typeof(ExcelDataReaderBatchRelationCoordinator<TWorkbook>)
            .GetMethod(nameof(BindTyped), BindingFlags.Static | BindingFlags.NonPublic)!
            .MakeGenericMethod(key.Parent, key.Child, key.Key);
        return (root, relation, sheetName, errors, cancellationToken) =>
        {
            try
            {
                method.Invoke(null, new object[] { root, relation, sheetName, errors, cancellationToken });
            }
            catch (TargetInvocationException exception) when (exception.InnerException != null)
            {
                ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            }
        };
    }

    /// <summary>
    /// 按强类型键将子实体绑定到父实体。
    /// </summary>
    /// <typeparam name="TParent">父实体类型。</typeparam>
    /// <typeparam name="TChild">子实体类型。</typeparam>
    /// <typeparam name="TKey">关联键类型。</typeparam>
    /// <param name="root">工作簿根模型。</param>
    /// <param name="relation">关系定义。</param>
    /// <param name="sheetName">子工作表名称。</param>
    /// <param name="errors">接收关系错误的集合。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void BindTyped<TParent, TChild, TKey>(object root, ExcelRelationRequest relation,
        string sheetName, ICollection<ExcelImportError> errors, CancellationToken cancellationToken)
        where TParent : class
        where TChild : class
    {
        var parents = (relation.Parents(root) as IEnumerable)?.Cast<TParent>().ToArray()
            ?? Array.Empty<TParent>();
        var children = (relation.Children(root) as IEnumerable)?.Cast<TChild>().ToArray()
            ?? Array.Empty<TChild>();
        if (children.Length == 0)
            return;
        var comparer = relation.Comparer as IEqualityComparer<TKey> ?? EqualityComparer<TKey>.Default;
        var parentByKey = new Dictionary<TKey, TParent>(comparer);
        var nullParent = default(TParent);
        var hasNullParent = false;
        foreach (var parent in parents)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var key = ((Func<TParent, TKey>)relation.ParentKey)(parent);
            if (key is null)
            {
                if (!hasNullParent)
                {
                    nullParent = parent;
                    hasNullParent = true;
                }
            }
            else
            {
                parentByKey.TryAdd(key, parent);
            }
        }

        foreach (var child in children)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var key = ((Func<TChild, TKey>)relation.ChildKey)(child);
            TParent parent = null;
            if (key is null)
                parent = hasNullParent ? nullParent : null;
            else
                parentByKey.TryGetValue(key, out parent);
            if (parent == null)
            {
                errors.Add(CreateRelationError(sheetName, $"未找到关联父实体: {key}", key));
                continue;
            }
            var navigation = relation.Navigation(parent);
            if (navigation == null)
                continue;
            AddToCollection<TChild>(navigation, child);
        }
    }

    /// <summary>
    /// 将子实体添加到父实体导航集合。
    /// </summary>
    /// <typeparam name="TItem">子实体类型。</typeparam>
    /// <param name="collection">父实体导航集合。</param>
    /// <param name="item">待添加的子实体。</param>
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
    /// 创建关系绑定错误。
    /// </summary>
    /// <param name="sheetName">关联错误所属的工作表名称。</param>
    /// <param name="message">错误消息。</param>
    /// <param name="rawValue">可选的原始键值。</param>
    /// <returns>结构化关系错误。</returns>
    private static ExcelImportError CreateRelationError(string sheetName, string message,
        object rawValue = null)
        => new ExcelImportError(ExcelImportErrorCode.Relationship, message, sheetName, 0, 0, null,
            rawValue: rawValue);
}

/// <summary>
/// 保存工作簿关系绑定阶段的结果。
/// </summary>
internal sealed class ExcelDataReaderRelationBindingResult
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderRelationBindingResult" /> 类型的实例。
    /// </summary>
    /// <param name="groups">按工作表分组的关系错误。</param>
    /// <param name="truncated">关系错误是否因预算而截断。</param>
    /// <param name="truncatedSheet">发生截断的工作表请求。</param>
    internal ExcelDataReaderRelationBindingResult(
        IReadOnlyList<ExcelDataReaderRelationErrorGroup> groups, bool truncated,
        ExcelSheetImportRequest truncatedSheet)
    {
        Groups = groups ?? Array.Empty<ExcelDataReaderRelationErrorGroup>();
        Truncated = truncated;
        TruncatedSheet = truncatedSheet;
    }

    /// <summary>
    /// 获取按工作表分组的关系错误。
    /// </summary>
    internal IReadOnlyList<ExcelDataReaderRelationErrorGroup> Groups { get; }
    /// <summary>
    /// 获取关系错误是否因预算而截断。
    /// </summary>
    internal bool Truncated { get; }
    /// <summary>
    /// 获取发生截断的工作表请求。
    /// </summary>
    internal ExcelSheetImportRequest TruncatedSheet { get; }
}

/// <summary>
/// 保存单个工作表的关系错误。
/// </summary>
internal sealed class ExcelDataReaderRelationErrorGroup
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderRelationErrorGroup" /> 类型的实例。
    /// </summary>
    /// <param name="sheet">关系错误所属的工作表请求。</param>
    /// <param name="errors">工作表关系错误。</param>
    internal ExcelDataReaderRelationErrorGroup(ExcelSheetImportRequest sheet,
        IReadOnlyList<ExcelImportError> errors)
    {
        Sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
        Errors = errors ?? Array.Empty<ExcelImportError>();
    }

    /// <summary>
    /// 获取关系错误所属的工作表请求。
    /// </summary>
    internal ExcelSheetImportRequest Sheet { get; }
    /// <summary>
    /// 获取工作表关系错误。
    /// </summary>
    internal IReadOnlyList<ExcelImportError> Errors { get; }
}
