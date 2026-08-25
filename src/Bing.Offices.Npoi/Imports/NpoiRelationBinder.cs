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

    private static ExcelImportError CreateError(string message,
        IReadOnlyDictionary<object, SourceLocation> sourceLocations, object source, object key)
    {
        sourceLocations.TryGetValue(source, out var location);
        return new ExcelImportError(ExcelImportErrorCode.Relationship, message,
            location?.SheetName ?? string.Empty, location?.RowIndex ?? 0, 0, null, "Key", null, key);
    }
}
