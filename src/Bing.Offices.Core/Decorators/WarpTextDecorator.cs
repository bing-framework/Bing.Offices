using System;
using System.Collections.Generic;
using System.Text;
using Bing.Offices.Abstractions.Decorators;
using Bing.Offices.Attributes;

namespace Bing.Offices.Decorators
{
    /// <summary>
    /// 自动换行装饰器
    /// </summary>
    [BindDecorator(typeof(WrapTextAttribute))]
    internal class WarpTextDecorator:IDecorator
    {
        
    }
}
