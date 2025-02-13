using System;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Tests;

/// <summary>
/// 注入容器
/// </summary>
public static class DependencyContainer
{
    /// <summary>
    /// 服务提供程序
    /// </summary>
    private static readonly IServiceProvider ServiceProvider;

    /// <summary>
    /// 静态构造函数
    /// </summary>
    static DependencyContainer()
    {
        var serviceCollection = new ServiceCollection();
        serviceCollection.AddNpoi();
        ServiceProvider = serviceCollection.BuildServiceProvider();
    }

    /// <summary>
    /// 解析
    /// </summary>
    /// <typeparam name="TComponent">组件类型</typeparam>
    /// <param name="component">组件</param>
    public static TComponent Resolve<TComponent>(this TComponent component) => ServiceProvider.GetRequiredService<TComponent>();
}