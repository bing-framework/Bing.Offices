using System.Text;
using Bing.Reflection;
using Bing.Text;

namespace Bing.Offices.Extensions;

/// <summary>
/// 类型名称和枚举元数据解析扩展。
/// </summary>
public static class TypeExtensions
{
    /// <summary>
    /// 将类型转换为不含程序集限定名的 C# 泛型类型名称。
    /// </summary>
    /// <param name="type">类型</param>
    /// <returns>类型的 C# 名称；泛型类型包含其递归解析后的类型参数。</returns>
    public static string GetCSharpTypeName(this Type type)
    {
        var sb = new StringBuilder();
        var name = type.Name;
        if (!type.IsGenericType)
            return name;
        sb.Append(name.Substring(0, name.IndexOf('`')));
        sb.Append("<");
        sb.Append(string.Join(", ", type.GetGenericArguments().Select(t => t.GetCSharpTypeName())));
        sb.Append(">");

        return sb.ToString();
    }

    /// <summary>
    /// 创建枚举显示文本到整数值的映射。
    /// </summary>
    /// <param name="type">类型</param>
    /// <returns>
    /// key：返回显示名称或描述 <br/>
    /// value：值
    /// </returns>
    public static IDictionary<string, int> GetEnumTextAndValues(this Type type)
    {
        if (!type.IsEnum)
            throw new InvalidOperationException();
        var items = type.GetEnumDefinitionList();
        var dict = new Dictionary<string, int>();
        // 枚举名 值 显示名称 描述
        foreach (var item in items) 
            dict.Add(item.Description ?? item.DisplayName ?? item.Name, item.Value);
        return dict;
    }

    /// <summary>
    /// 读取枚举名称、整数值、显示名称和描述组成的定义列表。
    /// </summary>
    /// <param name="type">类型</param>
    /// <returns>枚举定义列表；传入类型不是枚举时返回 null。</returns>
    public static IEnumerable<(string Name, int Value, string DisplayName, string Description)> GetEnumDefinitionList(this Type type)
    {
        var list = new List<(string Name, int Value, string DisplayName, string Description)>();
        var attrType = type;
        if (!attrType.IsEnum)
            return null;
        var names = Enum.GetNames(attrType);
        var values = Enum.GetValues(attrType);
        var index = 0;
        foreach (var value in values)
        {
            var name = names[index];
            var field = attrType.GetField(name);
            var displayName = TypeReflections.GetDisplayName(field);
            var des = TypeReflections.GetDescription(field);
            (string Name, int Value, string DisplayName, string Description) item = new(
                name,
                Convert.ToInt32(value),
                displayName.IsNullOrWhiteSpace() ? null : displayName,
                des.IsNullOrWhiteSpace() ? null : des
            );
            list.Add(item);
            index++;
        }

        return list;
    }

    /// <summary>
    /// 获取保留原始底层值的枚举定义列表。
    /// </summary>
    /// <param name="type">枚举类型。</param>
    /// <returns>保留枚举底层值的名称、值、显示名和描述列表。</returns>
    internal static IEnumerable<(string Name, object Value, string DisplayName, string Description)>
        GetEnumValueDefinitionList(this Type type)
    {
        if (!type.IsEnum)
            return Enumerable.Empty<(string Name, object Value, string DisplayName, string Description)>();

        var list = new List<(string Name, object Value, string DisplayName, string Description)>();
        foreach (var name in Enum.GetNames(type))
        {
            var field = type.GetField(name);
            var displayName = TypeReflections.GetDisplayName(field);
            var description = TypeReflections.GetDescription(field);
            list.Add((name, Enum.Parse(type, name),
                displayName.IsNullOrWhiteSpace() ? null : displayName,
                description.IsNullOrWhiteSpace() ? null : description));
        }

        return list;
    }

        
}
