using System.Reflection;

namespace Bing.Offices.Mappings;

internal interface IExcelCompiledMappingColumn
{
    PropertyInfo Property { get; }
    Func<object, object> Getter { get; }
    Action<object, object> Setter { get; }
    IReadOnlyList<Attribute> Attributes { get; }
}
