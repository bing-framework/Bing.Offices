using System;
using System.Collections.Generic;
using System.Reflection;
using Bing.Offices.Attributes;

namespace Bing.Offices.Mappings;

internal interface IExcelCompiledMappingColumn
{
    PropertyInfo Property { get; }
    Func<object, object> Getter { get; }
    Action<object, object> Setter { get; }
    IReadOnlyList<Attribute> Attributes { get; }
}
