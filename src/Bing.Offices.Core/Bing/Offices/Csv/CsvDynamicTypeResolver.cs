using System;

namespace Bing.Offices.Csv;

internal static class CsvDynamicTypeResolver
{
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
