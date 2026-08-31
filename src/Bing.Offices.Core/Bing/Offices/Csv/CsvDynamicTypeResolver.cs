using System;

namespace Bing.Offices.Csv;

/// <summary>解析 CSV 动态列允许使用的 CLR 类型名称。</summary>
internal static class CsvDynamicTypeResolver
{
    /// <summary>将配置中的动态列类型名称解析为受支持的 CLR 类型。</summary>
    /// <param name="name">类型名称；为空时使用 string。</param>
    /// <returns>对应的 CLR 类型。</returns>
    public static Type Resolve(string name)
    {
        switch ((name ?? "string").ToLowerInvariant())
        {
            case "object": return typeof(object);
            case "string": return typeof(string);
            case "boolean": case "bool": return typeof(bool);
            case "byte": return typeof(byte);
            case "int16": return typeof(short);
            case "int32": case "int": return typeof(int);
            case "int64": case "long": return typeof(long);
            case "single": case "float": return typeof(float);
            case "double": return typeof(double);
            case "decimal": return typeof(decimal);
            case "datetime": return typeof(DateTime);
            case "datetimeoffset": return typeof(DateTimeOffset);
            case "guid": return typeof(Guid);
            case "bytes": return typeof(byte[]);
            default: throw new InvalidOperationException($"动态列数据类型不在允许列表中: {name}");
        }
    }
}
