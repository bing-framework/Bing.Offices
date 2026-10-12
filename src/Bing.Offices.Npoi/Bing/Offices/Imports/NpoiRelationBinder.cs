using System.Collections.Concurrent;
using System.Reflection;
using Bing.Offices.Exceptions;

namespace Bing.Offices.Imports;

/// <summary>
/// 执行显式 Excel 父子关系绑定，隔离导入器的关系处理职责。
/// </summary>
internal static class NpoiRelationBinder
{
    /// <summary>
    /// 执行一组已编译的 Excel 父子关系绑定。
    /// </summary>
    /// <param name="root">接收关系绑定结果的工作簿根实体。</param>
    /// <param name="request">定义父项、子项及关联键选择器的关系请求。</param>
    /// <param name="errors">接收绑定错误的导入错误收集器。</param>
    /// <param name="sourceLocations">实体到原始工作表位置的映射。</param>
    /// <param name="sourceSheets">按列表集合索引的导入结果。</param>
    /// <param name="relationKeyBudget">关系父键索引的剩余预算。</param>
    /// <param name="cancellationToken">遍历关系集合时检查的取消令牌。</param>
    private delegate void BindInvoker(object root, ExcelRelationRequest request,
        ExcelImportErrorCollector errors, IReadOnlyDictionary<object, SourceLocation> sourceLocations,
        IReadOnlyDictionary<object, ExcelSheetImportResult> sourceSheets,
        NpoiRelationKeyBudget relationKeyBudget, CancellationToken cancellationToken);

    /// <summary>
    /// 按工作簿、父项、子项和关联键类型隔离关系绑定委托缓存。
    /// </summary>
    private static readonly ConcurrentDictionary<(Type Workbook, Type Parent, Type Child, Type Key), BindInvoker>
        BindInvokers = new();

    /// <summary>
    /// 绑定一个具体的父子关系请求。
    /// </summary>
    /// <typeparam name="TWorkbook">包含父子集合的工作簿根实体类型。</typeparam>
    /// <param name="root">已完成工作表导入的工作簿根实体。</param>
    /// <param name="request">定义集合、键选择器和导航集合的关系请求。</param>
    /// <param name="errors">接收关系绑定错误的工作簿级错误收集器。</param>
    /// <param name="sourceLocations">实体到原始工作表位置的引用映射。</param>
    /// <param name="cancellationToken">遍历父项和子项时检查的取消令牌。</param>
    public static void Bind<TWorkbook>(TWorkbook root, ExcelRelationRequest request,
        ExcelImportErrorCollector errors, IReadOnlyDictionary<object, SourceLocation> sourceLocations,
        CancellationToken cancellationToken)
        where TWorkbook : class, new() =>
        Bind(root, request, errors, sourceLocations, null, null, cancellationToken);

    /// <summary>
    /// 绑定一个具体的父子关系请求并应用关系键预算。
    /// </summary>
    /// <typeparam name="TWorkbook">包含父子集合的工作簿根实体类型。</typeparam>
    /// <param name="root">已完成工作表导入的工作簿根实体。</param>
    /// <param name="request">定义集合、键选择器和导航集合的关系请求。</param>
    /// <param name="errors">接收关系绑定错误的工作簿级错误收集器。</param>
    /// <param name="sourceLocations">实体到原始工作表位置的引用映射。</param>
    /// <param name="relationKeyBudget">关系父键索引的剩余预算。</param>
    /// <param name="cancellationToken">遍历父项和子项时检查的取消令牌。</param>
    public static void Bind<TWorkbook>(TWorkbook root, ExcelRelationRequest request,
        ExcelImportErrorCollector errors, IReadOnlyDictionary<object, SourceLocation> sourceLocations,
        NpoiRelationKeyBudget relationKeyBudget, CancellationToken cancellationToken)
        where TWorkbook : class, new() =>
        Bind(root, request, errors, sourceLocations, null, relationKeyBudget, cancellationToken);

    /// <summary>
    /// 按列表结果中的源行索引绑定一个关系请求。
    /// </summary>
    /// <typeparam name="TWorkbook">包含父子集合的工作簿根实体类型。</typeparam>
    /// <param name="root">已完成工作表导入的工作簿根实体。</param>
    /// <param name="request">定义集合、键选择器和导航集合的关系请求。</param>
    /// <param name="errors">接收关系错误的工作簿级错误收集器。</param>
    /// <param name="sourceSheets">按列表集合索引的导入结果。</param>
    /// <param name="relationKeyBudget">关系父键索引的剩余预算。</param>
    /// <param name="cancellationToken">遍历关系集合时检查的取消令牌。</param>
    internal static void BindWithSourceRows<TWorkbook>(TWorkbook root,
        ExcelRelationRequest request, ExcelImportErrorCollector errors,
        IReadOnlyDictionary<object, ExcelSheetImportResult> sourceSheets,
        NpoiRelationKeyBudget relationKeyBudget, CancellationToken cancellationToken)
        where TWorkbook : class, new()
        => Bind(root, request, errors, null, sourceSheets, relationKeyBudget, cancellationToken);

    private static void Bind<TWorkbook>(TWorkbook root, ExcelRelationRequest request,
        ExcelImportErrorCollector errors,
        IReadOnlyDictionary<object, SourceLocation> sourceLocations,
        IReadOnlyDictionary<object, ExcelSheetImportResult> sourceSheets,
        NpoiRelationKeyBudget relationKeyBudget, CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        relationKeyBudget ??= new NpoiRelationKeyBudget(null);
        var key = (typeof(TWorkbook), request.ParentType, request.ChildType, request.ParentKey.Method.ReturnType);
        BindInvokers.GetOrAdd(key, CreateBindInvoker)(root, request, errors, sourceLocations,
            sourceSheets, relationKeyBudget, cancellationToken);
    }

    /// <summary>
    /// 检查指定关系类型组合是否已有缓存委托，仅供职责级测试验证缓存合同。
    /// </summary>
    /// <param name="workbookType">工作簿根实体类型。</param>
    /// <param name="parentType">父项实体类型。</param>
    /// <param name="childType">子项实体类型。</param>
    /// <param name="keyType">父子关联键类型。</param>
    /// <returns>指定类型组合已有缓存委托时为 true，否则为 false。</returns>
    internal static bool IsCached(Type workbookType, Type parentType, Type childType, Type keyType) =>
        BindInvokers.ContainsKey((workbookType, parentType, childType, keyType));

    /// <summary>
    /// 为关系类型组合创建泛型绑定委托。
    /// </summary>
    /// <param name="key">工作簿、父项、子项和关联键类型组合。</param>
    /// <returns>绑定到指定类型组合的关系处理委托。</returns>
    private static BindInvoker CreateBindInvoker((Type Workbook, Type Parent, Type Child, Type Key) key)
    {
        var method = typeof(NpoiRelationBinder).GetMethod(nameof(CreateTypedBindInvoker),
            BindingFlags.Static | BindingFlags.NonPublic)!.MakeGenericMethod(
            key.Workbook, key.Parent, key.Child, key.Key);
        return (BindInvoker)method.Invoke(null, null)!;
    }

    /// <summary>
    /// 创建调用具体泛型关系绑定核心方法的委托。
    /// </summary>
    /// <typeparam name="TWorkbook">工作簿根实体类型。</typeparam>
    /// <typeparam name="TParent">父项实体类型。</typeparam>
    /// <typeparam name="TChild">子项实体类型。</typeparam>
    /// <typeparam name="TKey">父子关联键类型。</typeparam>
    /// <returns>调用具体泛型绑定核心方法的委托。</returns>
    private static BindInvoker CreateTypedBindInvoker<TWorkbook, TParent, TChild, TKey>()
        where TWorkbook : class, new()
        where TParent : class
        where TChild : class =>
        (root, request, errors, sourceLocations, sourceSheets, relationKeyBudget, cancellationToken) =>
            BindCore<TWorkbook, TParent, TChild, TKey>((TWorkbook)root, request, errors,
                sourceLocations, sourceSheets, relationKeyBudget, cancellationToken);

    /// <summary>
    /// 使用具体泛型类型执行父子键关联并写入导航集合。
    /// </summary>
    /// <typeparam name="TWorkbook">包含父子集合的工作簿根实体类型。</typeparam>
    /// <typeparam name="TParent">父项实体类型。</typeparam>
    /// <typeparam name="TChild">子项实体类型。</typeparam>
    /// <typeparam name="TKey">父子关联键类型。</typeparam>
    /// <param name="root">工作簿根实体。</param>
    /// <param name="request">关系绑定请求。</param>
    /// <param name="errors">接收键缺失、重复和未匹配错误的收集器。</param>
    /// <param name="sourceLocations">实体来源位置映射。</param>
    /// <param name="sourceSheets">按列表集合索引的导入结果。</param>
    /// <param name="relationKeyBudget">关系父键索引的剩余预算。</param>
    /// <param name="cancellationToken">遍历过程检查的取消令牌。</param>
    private static void BindCore<TWorkbook, TParent, TChild, TKey>(TWorkbook root,
        ExcelRelationRequest request, ExcelImportErrorCollector errors,
        IReadOnlyDictionary<object, SourceLocation> sourceLocations,
        IReadOnlyDictionary<object, ExcelSheetImportResult> sourceSheets,
        NpoiRelationKeyBudget relationKeyBudget, CancellationToken cancellationToken)
        where TWorkbook : class, new()
        where TParent : class
        where TChild : class
    {
        ICollection<TParent> parents;
        ICollection<TChild> children;
        try
        {
            parents = (ICollection<TParent>)request.Parents(root);
            children = (ICollection<TChild>)request.Children(root);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesImportException("Excel 关系集合读取器执行失败。", exception, "NPOI",
                BingOfficesStage.Validate, code: BingOfficesErrorCode.UserExtensionFailed);
        }
        if (parents == null || children == null)
            throw new InvalidOperationException("关系绑定的父集合或子集合不可写入。");
        var parentByKey = new Dictionary<TKey,
            (TParent Parent, RelationSourceLocation SourceLocation)>(
                (IEqualityComparer<TKey>)request.Comparer);
        var parentIndex = 0;
        foreach (var parent in parents)
        {
            var sourceLocation = GetSourceLocation(sourceLocations, sourceSheets, parents, parent,
                parentIndex++);
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
            TKey key;
            try
            {
                key = ((Func<TParent, TKey>)request.ParentKey)(parent);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception) when (exception is not OutOfMemoryException
                && exception is not StackOverflowException)
            {
                throw new BingOfficesImportException("Excel 父项键选择器执行失败。", exception, "NPOI",
                    BingOfficesStage.Validate, code: BingOfficesErrorCode.UserExtensionFailed);
            }
            if (key == null)
            {
                errors.Add(CreateError("父项键为空。", sourceLocation, key));
                continue;
            }
            if (parentByKey.ContainsKey(key))
            {
                errors.Add(CreateError($"父项键重复: {key}", sourceLocation, key));
                continue;
            }
            if (!relationKeyBudget.TryReserve())
            {
                errors.Add(CreateResourceError("关系父键索引数量超过限制。", sourceLocation));
                errors.MarkTruncated();
                return;
            }
            parentByKey.Add(key, (parent, sourceLocation));
        }

        // 父键索引已经覆盖关系查找；导航目标按父项缓存后逐项写入，避免为所有子项建立分组列表。
        var navigationByParent = new Dictionary<TParent,
            (ICollection<TChild> Target, RelationSourceLocation SourceLocation)>();
        var childIndex = 0;
        foreach (var child in children)
        {
            var sourceLocation = GetSourceLocation(sourceLocations, sourceSheets, children, child,
                childIndex++);
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
            TKey key;
            try
            {
                key = ((Func<TChild, TKey>)request.ChildKey)(child);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception) when (exception is not OutOfMemoryException
                && exception is not StackOverflowException)
            {
                throw new BingOfficesImportException("Excel 子项键选择器执行失败。", exception, "NPOI",
                    BingOfficesStage.Validate, code: BingOfficesErrorCode.UserExtensionFailed);
            }
            if (key == null)
            {
                errors.Add(CreateError("子项键为空。", sourceLocation, key));
                continue;
            }
            if (!parentByKey.TryGetValue(key, out var parentEntry))
            {
                errors.Add(CreateError($"子项找不到父项: {key}", sourceLocation, key));
                continue;
            }
            var parent = parentEntry.Parent;

            if (!navigationByParent.TryGetValue(parent, out var navigationTarget))
            {
                ICollection<TChild> target;
                try
                {
                    target = (ICollection<TChild>)request.Navigation(parent);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception exception) when (exception is not OutOfMemoryException
                    && exception is not StackOverflowException)
                {
                    throw new BingOfficesImportException("Excel 关系导航集合读取器执行失败。", exception,
                        "NPOI", BingOfficesStage.Validate, code: BingOfficesErrorCode.UserExtensionFailed);
                }
                navigationTarget = (target, parentEntry.SourceLocation);
                navigationByParent[parent] = navigationTarget;
            }

            // 空导航集合只在首次命中该父项时报告，保持原有按父项聚合的错误数量。
            if (navigationTarget.Target == null)
                continue;
            try
            {
                navigationTarget.Target.Add(child);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception) when (exception is not OutOfMemoryException
                && exception is not StackOverflowException)
            {
                throw new BingOfficesImportException("Excel 关系导航集合写入器执行失败。", exception,
                    "NPOI", BingOfficesStage.Validate, code: BingOfficesErrorCode.UserExtensionFailed);
            }
        }

        if (errors.IsLimitReached)
        {
            errors.MarkTruncated();
            return;
        }

        foreach (var pair in navigationByParent)
        {
            if (pair.Value.Target != null)
                continue;
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
            errors.Add(CreateError("导航集合为空且不可写入。", pair.Value.SourceLocation, null));
        }
    }

    /// <summary>
    /// 使用源实体的导入位置创建关系绑定错误。
    /// </summary>
    /// <param name="message">描述关系绑定失败的消息。</param>
    /// <param name="sourceLocation">产生错误的实体来源位置。</param>
    /// <param name="key">导致错误的关联键。</param>
    /// <returns>带有可用工作表和行号的导入错误。</returns>
    private static ExcelImportError CreateError(string message,
        RelationSourceLocation sourceLocation, object key)
    {
        return new ExcelImportError(ExcelImportErrorCode.Relationship, message,
            sourceLocation.SheetName ?? string.Empty, sourceLocation.RowIndex, 0, null, "Key", null, key);
    }

    /// <summary>
    /// 根据源位置创建关系索引资源错误。
    /// </summary>
    /// <param name="message">描述资源限制的消息。</param>
    /// <param name="sourceLocation">触发资源限制的父项来源位置。</param>
    /// <returns>带有可用工作表和行号的资源错误。</returns>
    private static ExcelImportError CreateResourceError(string message,
        RelationSourceLocation sourceLocation)
    {
        return new ExcelImportError(ExcelImportErrorCode.ResourceLimit, message,
            sourceLocation.SheetName ?? string.Empty, sourceLocation.RowIndex, 0, null);
    }

    private static RelationSourceLocation GetSourceLocation(
        IReadOnlyDictionary<object, SourceLocation> sourceLocations,
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
}
