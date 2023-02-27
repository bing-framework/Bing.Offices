namespace Bing.Offices.Decorators;

/// <summary>
/// 装饰器上下文
/// </summary>
public interface IDecoratorContext
{
    /// <summary>
    /// 类型装饰器信息
    /// </summary>
    TypeDecoratorInfo TypeDecoratorInfo { get; set; }
}