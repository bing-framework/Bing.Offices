﻿using System;

// ReSharper disable once CheckNamespace
namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 合并单元格特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class MergeColumnsAttribute : DecoratorAttributeBase
    {
    }
}
