using System;

namespace Bing.Offices.Attributes;

/// <summary>
/// 含有动态列
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class HasDynamicColumnAttribute : Attribute
{
}