using System;

// ReSharper disable once CheckNamespace
namespace Bing.Offices.Attributes;

/// <summary>
/// 自动换行特性
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public class WrapTextAttribute : DecoratorAttributeBase
{
}