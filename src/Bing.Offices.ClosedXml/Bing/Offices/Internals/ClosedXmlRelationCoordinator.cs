using System.Collections;
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
    /// <param name="cancellationToken">取消令牌。</param>
    internal static void Bind(object root, IReadOnlyList<ExcelRelationRequest> relations,
        ICollection<ExcelImportError> errors, int? maxErrors, CancellationToken cancellationToken)
    {
        if (root == null || relations == null || relations.Count == 0)
            return;
        foreach (var relation in relations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var parents = (relation.Parents(root) as IEnumerable)?.Cast<object>().ToArray()
                ?? Array.Empty<object>();
            var children = (relation.Children(root) as IEnumerable)?.Cast<object>().ToArray()
                ?? Array.Empty<object>();
            var parentKeys = new List<(object Parent, object Key)>();
            foreach (var parent in parents)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (Reached(errors, maxErrors))
                    break;
                var parentKey = relation.ParentKey.DynamicInvoke(parent);
                if (parentKey == null)
                {
                    AddError(errors, maxErrors, new ExcelImportError(
                        ExcelImportErrorCode.Relationship, "父项关联键为空。", string.Empty, 0, 0,
                        relation.ParentType.Name));
                    continue;
                }
                if (parentKeys.Any(item => KeysEqual(relation, item.Key, parentKey)))
                {
                    AddError(errors, maxErrors, new ExcelImportError(
                        ExcelImportErrorCode.Relationship, "父项关联键重复。", string.Empty, 0, 0,
                        relation.ParentType.Name, rawValue: parentKey));
                    continue;
                }
                parentKeys.Add((parent, parentKey));
            }
            foreach (var child in children)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (Reached(errors, maxErrors))
                    break;
                var childKey = relation.ChildKey.DynamicInvoke(child);
                if (childKey == null)
                {
                    AddError(errors, maxErrors, new ExcelImportError(
                        ExcelImportErrorCode.Relationship, "子项关联键为空。", string.Empty, 0, 0,
                        relation.ChildType.Name));
                    continue;
                }
                var parent = parentKeys.FirstOrDefault(item => KeysEqual(relation, item.Key, childKey)).Parent;
                if (parent == null)
                {
                    AddError(errors, maxErrors, new ExcelImportError(
                        ExcelImportErrorCode.Relationship, "无法绑定父子关系。", string.Empty, 0, 0,
                        relation.ChildType.Name, rawValue: childKey));
                    continue;
                }
                var navigation = relation.Navigation(parent);
                if (!TryAdd(navigation, relation.ChildType, child))
                {
                    AddError(errors, maxErrors, new ExcelImportError(
                        ExcelImportErrorCode.Relationship, "父实体子集合导航属性不可写。", string.Empty, 0, 0,
                        relation.ChildType.Name, rawValue: childKey));
                    continue;
                }
            }
        }
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
    /// 使用关系请求指定的比较器比较两个关联键。
    /// </summary>
    /// <param name="relation">关系请求。</param>
    /// <param name="left">左侧关联键。</param>
    /// <param name="right">右侧关联键。</param>
    /// <returns>两个关联键相等时返回 <see langword="true" />。</returns>
    private static bool KeysEqual(ExcelRelationRequest relation, object left, object right)
    {
        if (relation.Comparer == null)
            return Equals(left, right);
        var comparerInterface = relation.Comparer.GetType().GetInterfaces()
            .FirstOrDefault(type => type.IsGenericType
                && type.GetGenericTypeDefinition() == typeof(IEqualityComparer<>));
        var equalsMethod = comparerInterface?.GetMethod(nameof(IEqualityComparer<object>.Equals));
        return equalsMethod == null
            ? Equals(left, right)
            : (bool)equalsMethod.Invoke(relation.Comparer, new[] { left, right });
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
    /// 判断错误集合是否已达到配置上限。
    /// </summary>
    /// <param name="errors">错误集合。</param>
    /// <param name="maxErrors">允许记录的最大错误数。</param>
    /// <returns>已达到上限时返回 <see langword="true" />。</returns>
    private static bool Reached(ICollection<ExcelImportError> errors, int? maxErrors) =>
        maxErrors.HasValue && errors != null && errors.Count >= maxErrors.Value;
}
