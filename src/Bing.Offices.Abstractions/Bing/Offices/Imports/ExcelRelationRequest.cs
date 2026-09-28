using System.Linq.Expressions;
using System.ComponentModel;
using Bing.Offices.Configurations;

namespace Bing.Offices.Imports;

/// <summary>
/// Workbook 导入父子关系执行描述。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelRelationRequest
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelRelationRequest" /> 类型的实例。
    /// </summary>
    /// <param name="parents">读取父实体集合的委托。</param>
    /// <param name="children">读取子实体集合的委托。</param>
    /// <param name="parentKey">读取父实体关联键的委托。</param>
    /// <param name="childKey">读取子实体关联键的委托。</param>
    /// <param name="navigation">写入父实体子集合的委托。</param>
    /// <param name="parentType">父实体类型。</param>
    /// <param name="childType">子实体类型。</param>
    /// <param name="comparer">关联键比较器。</param>
    private ExcelRelationRequest(Func<object, object> parents, Func<object, object> children, Delegate parentKey,
        Delegate childKey, Func<object, object> navigation, Type parentType, Type childType,
        object comparer)
    {
        Parents = parents;
        Children = children;
        ParentKey = parentKey;
        ChildKey = childKey;
        Navigation = navigation;
        ParentType = parentType;
        ChildType = childType;
        Comparer = comparer;
    }

    /// <summary>
    /// 获取父集合读取器。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> Parents { get; }
    /// <summary>
    /// 获取子集合读取器。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> Children { get; }
    /// <summary>
    /// 获取父键委托。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Delegate ParentKey { get; }
    /// <summary>
    /// 获取子键委托。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Delegate ChildKey { get; }
    /// <summary>
    /// 获取子导航属性写入器。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> Navigation { get; }
    /// <summary>
    /// 获取父实体类型。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Type ParentType { get; }
    /// <summary>
    /// 获取子实体类型。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Type ChildType { get; }
    /// <summary>
    /// 获取键比较器。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public object Comparer { get; }

    /// <summary>
    /// 根据父子表达式创建关系执行描述。
    /// </summary>
    /// <typeparam name="TWorkbook">工作簿模型类型。</typeparam>
    /// <typeparam name="TParent">父实体类型。</typeparam>
    /// <typeparam name="TChild">子实体类型。</typeparam>
    /// <typeparam name="TKey">父子实体关联键类型。</typeparam>
    /// <param name="parents">工作簿父实体集合属性表达式。</param>
    /// <param name="children">工作簿子实体集合属性表达式。</param>
    /// <param name="parentKey">读取父实体关联键的函数。</param>
    /// <param name="childKey">读取子实体关联键的函数。</param>
    /// <param name="navigation">父实体子集合属性表达式。</param>
    /// <param name="comparer">可选的关联键比较器。</param>
    /// <returns>已编译的父子关系执行描述。</returns>
    internal static ExcelRelationRequest Create<TWorkbook, TParent, TChild, TKey>(
        Expression<Func<TWorkbook, ICollection<TParent>>> parents,
        Expression<Func<TWorkbook, ICollection<TChild>>> children,
        Func<TParent, TKey> parentKey,
        Func<TChild, TKey> childKey,
        Expression<Func<TParent, ICollection<TChild>>> navigation,
        IEqualityComparer<TKey> comparer)
        where TParent : class where TChild : class
    {
        var parentGetter = parents.Compile();
        var childGetter = children.Compile();
        var navigationGetter = navigation.Compile();
        return new ExcelRelationRequest(value => parentGetter((TWorkbook)value), value => childGetter((TWorkbook)value),
            parentKey, childKey, value => navigationGetter((TParent)value), typeof(TParent), typeof(TChild),
            comparer ?? EqualityComparer<TKey>.Default);
    }
}
