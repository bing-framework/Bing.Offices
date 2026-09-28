using System.Linq.Expressions;

namespace Bing.Offices.Exports.DynamicColumns;

/// <summary>
/// 将泛型动态值读取器转换为对象字典读取器的扩展类。
/// </summary>
internal static class ExcelDynamicGetterExtensions
{
    /// <summary>
    /// 将泛型动态值读取器适配为对象读取器。
    /// </summary>
    /// <typeparam name="T">动态值所属的实体类型。</typeparam>
    /// <param name="getter">读取实体动态值的泛型委托。</param>
    /// <returns>接受对象并调用泛型读取器的委托。</returns>
    public static Func<object, IDictionary<string, object>> ToObjectDictionaryGetter<T>(
        this Func<T, IDictionary<string, object>> getter) => value => getter((T)value);
}
