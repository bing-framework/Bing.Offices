using System.Collections.Generic;

namespace Bing.Offices.Abstractions.Decorators
{
    /// <summary>
    /// 装饰器
    /// </summary>
    public interface IDecorator
    {
        byte[] Handler<T>(byte[] workbookBytes, List<T> data, IDecoratorContext context);
    }
}
