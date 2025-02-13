using System.Collections.Concurrent;
using System.Reflection;
using Bing.Collections;
using Bing.Offices.Attributes;
using Bing.Offices.Decorators;

namespace Bing.Offices.Factories;

/// <summary>
/// 装饰器工厂
/// </summary>
internal static class DecoratorFactory
{
    /// <summary>
    /// 装饰器字典
    /// </summary>
    private static readonly IDictionary<Type, IDecorator> DecoratorDict = new ConcurrentDictionary<Type, IDecorator>();

    /// <summary>
    /// 装饰器集合字典
    /// </summary>
    private static readonly IDictionary<Type, IList<IDecorator>> DecoratorsDict = new ConcurrentDictionary<Type, IList<IDecorator>>();

    /// <summary>
    /// 创建装饰器实例
    /// </summary>
    /// <param name="type">绑定装饰器的特性类型</param>
    public static IDecorator CreateInstance(Type type)
    {
        if (DecoratorDict.ContainsKey(type))
            return DecoratorDict[type];
        var decoratorType = Assembly.GetAssembly(type).GetTypes().ToList()
            ?.Where(t => typeof(IDecorator).IsAssignableFrom(t))?.FirstOrDefault(t =>
                t.IsDefined(typeof(BindDecoratorAttribute)) &&
                t.GetCustomAttribute<BindDecoratorAttribute>()?.DecoratorType == type);
        if (decoratorType == null)
            throw new ArgumentNullException(nameof(decoratorType), "找不到指定装饰器类型");
        var decorator = Activator.CreateInstance(decoratorType) as IDecorator;
        DecoratorDict[type] = decorator;
        return decorator;
    }


    /// <summary>
    /// 创建装饰器实例集合
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    public static IList<IDecorator> CreateInstances<T>() where T : new() => CreateInstances(typeof(T));

    /// <summary>
    /// 创建装饰器实例集合
    /// </summary>
    /// <param name="type">类型</param>
    public static IList<IDecorator> CreateInstances(Type type)
    {
        if (DecoratorsDict.ContainsKey(type))
            return DecoratorsDict[type];
        var decorators = new List<IDecorator>();
        var decoratorAttributes = new List<DecoratorAttributeBase>();
        var typeDecoratorInfo = TypeDecoratorInfoFactory.CreateInstance(type);

        decoratorAttributes.AddRange(typeDecoratorInfo.TypeDecorators);
        typeDecoratorInfo.PropertyDecoratorInfos.ForEach(x => decoratorAttributes.AddRange(x.Decorators));

        decoratorAttributes.Distinct(new DecoratorAttributeComparer()).ToList().ForEach(x =>
        {
            var decorator = CreateInstance(x.GetType());
            if (decorator != null)
                decorators.Add(decorator);
        });
        DecoratorsDict[type] = decorators;
        return decorators;
    }
}
