using System.Text;
using Bing.Reflection;
using Bing.Text;

namespace Bing.Offices.Extensions;

/// <summary>
/// 类型(<see cref="Type"/>) 扩展
/// </summary>
public static class TypeExtensions
{
    /// <summary>
    /// 获取CSharp类型名称
    /// </summary>
    /// <param name="type">类型</param>
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
    /// 获取枚举列表
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
    /// 获取枚举定义列表
    /// </summary>
    /// <param name="type">类型</param>
    /// <returns>返回枚举列表元组（名称、值、显示名、描述）</returns>
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
            var field = values.GetType().GetField(value.ToString());
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

        
}