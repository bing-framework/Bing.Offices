namespace Bing.Offices.Decorators;

/// <summary>
/// 装饰器上下文
/// </summary>
public class DecoratorContext : IDecoratorContext
{
    /// <summary>
    /// 类型装饰器信息
    /// </summary>
    public TypeDecoratorInfo TypeDecoratorInfo { get; set; }
}