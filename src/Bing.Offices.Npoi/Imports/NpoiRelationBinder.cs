using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading;
using Bing.Offices.Imports;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// 执行显式 Excel 父子关系绑定，隔离导入器的关系处理职责。
/// </summary>
internal static class NpoiRelationBinder
{
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
        where TWorkbook : class, new()
    {
        var method = typeof(NpoiRelationBinder).GetMethod(nameof(BindCore),
            BindingFlags.Static | BindingFlags.NonPublic);
        method.MakeGenericMethod(typeof(TWorkbook), request.ParentType, request.ChildType,
            request.ParentKey.Method.ReturnType).Invoke(null,
            new object[] { root, request, errors, sourceLocations, cancellationToken });
    }

    /// <summary>使用具体泛型类型执行父子键关联并写入导航集合。</summary>
    /// <typeparam name="TWorkbook">包含父子集合的工作簿根实体类型。</typeparam>
    /// <typeparam name="TParent">父项实体类型。</typeparam>
    /// <typeparam name="TChild">子项实体类型。</typeparam>
    /// <typeparam name="TKey">父子关联键类型。</typeparam>
    /// <param name="root">工作簿根实体。</param>
    /// <param name="request">关系绑定请求。</param>
    /// <param name="errors">接收键缺失、重复和未匹配错误的收集器。</param>
    /// <param name="sourceLocations">实体来源位置映射。</param>
    /// <param name="cancellationToken">遍历过程检查的取消令牌。</param>
    private static void BindCore<TWorkbook, TParent, TChild, TKey>(TWorkbook root,
        ExcelRelationRequest request, ExcelImportErrorCollector errors,
        IReadOnlyDictionary<object, SourceLocation> sourceLocations, CancellationToken cancellationToken)
        where TWorkbook : class, new()
        where TParent : class
        where TChild : class
    {
        var parents = (ICollection<TParent>)request.Parents(root);
        var children = (ICollection<TChild>)request.Children(root);
        if (parents == null || children == null)
            throw new InvalidOperationException("关系绑定的父集合或子集合不可写入。");
        var parentByKey = new Dictionary<TKey, TParent>((IEqualityComparer<TKey>)request.Comparer);
        foreach (var parent in parents)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
            var key = ((Func<TParent, TKey>)request.ParentKey)(parent);
            if (key == null)
            {
                errors.Add(CreateError("父项键为空。", sourceLocations, parent, key));
                continue;
            }
            if (!parentByKey.TryAdd(key, parent))
                errors.Add(CreateError($"父项键重复: {key}", sourceLocations, parent, key));
        }

        var childByParent = new Dictionary<TParent, List<TChild>>();
        foreach (var child in children)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
            var key = ((Func<TChild, TKey>)request.ChildKey)(child);
            if (key == null)
            {
                errors.Add(CreateError("子项键为空。", sourceLocations, child, key));
                continue;
            }
            if (!parentByKey.TryGetValue(key, out var parent))
            {
                errors.Add(CreateError($"子项找不到父项: {key}", sourceLocations, child, key));
                continue;
            }
            if (!childByParent.TryGetValue(parent, out var list))
                childByParent[parent] = list = new List<TChild>();
            list.Add(child);
        }

        foreach (var pair in childByParent)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
            var target = (ICollection<TChild>)request.Navigation(pair.Key);
            if (target == null)
            {
                errors.Add(CreateError("导航集合为空且不可写入。", sourceLocations, pair.Key, null));
                continue;
            }
            foreach (var child in pair.Value)
                target.Add(child);
        }
    }

    /// <summary>使用源实体的导入位置创建关系绑定错误。</summary>
    /// <param name="message">描述关系绑定失败的消息。</param>
    /// <param name="sourceLocations">实体来源位置映射。</param>
    /// <param name="source">产生错误的父项或子项实体。</param>
    /// <param name="key">导致错误的关联键。</param>
    /// <returns>带有可用工作表和行号的导入错误。</returns>
    private static ExcelImportError CreateError(string message,
        IReadOnlyDictionary<object, SourceLocation> sourceLocations, object source, object key)
    {
        sourceLocations.TryGetValue(source, out var location);
        return new ExcelImportError(ExcelImportErrorCode.Relationship, message,
            location?.SheetName ?? string.Empty, location?.RowIndex ?? 0, 0, null, "Key", null, key);
    }
}
