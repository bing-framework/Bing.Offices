using System.Reflection;

namespace Bing.Offices.Mappings;

/// <summary>
/// 包含已编译实体属性访问器的列映射契约。
/// </summary>
internal interface IExcelCompiledMappingColumn
{
    /// <summary>
    /// 获取实体上参与映射的属性元数据。
    /// </summary>
    PropertyInfo Property { get; }

    /// <summary>
    /// 获取从实体读取属性值的委托。
    /// </summary>
    Func<object, object> Getter { get; }

    /// <summary>
    /// 获取向实体写入属性值的委托。
    /// </summary>
    Action<object, object> Setter { get; }

    /// <summary>
    /// 获取声明在属性上的特性集合。
    /// </summary>
    IReadOnlyList<Attribute> Attributes { get; }
}
