using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Exceptions;
using Bing.Offices.Extensions;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 映射配置加载器静态与 DI 边界合同测试。
/// </summary>
public sealed class ExcelMappingConfigurationLoaderTest
{
    /// <summary>
    /// 验证依赖注入注册的默认加载器使用已注册的观察器。
    /// </summary>
    [Fact]
    public void DependencyInjection_DefaultLoader_ShouldUseRegisteredObserver()
    {
        // Arrange
        var services = new ServiceCollection();
        var observer = new RecordingObserver();
        services.AddSingleton<IBingOfficesExceptionObserver>(observer);
        services.AddBingOfficesNpoi();
        using var provider = services.BuildServiceProvider();

        // Act
        var loader = provider.GetRequiredService<IExcelMappingConfigurationLoader>();
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            loader.FromXmlDocument("<ExcelMappingDocument>"));

        // Assert
        Assert.Same(exception, Assert.Single(observer.Exceptions));
    }

    /// <summary>
    /// 记录测试场景中的通知事件。
    /// </summary>
    private sealed class RecordingObserver : IBingOfficesExceptionObserver
    {
        /// <summary>
        /// 获取异常集合。
        /// </summary>
        public List<BingOfficesException> Exceptions { get; } = new();

        /// <inheritdoc />
        public void Observe(BingOfficesException exception) => Exceptions.Add(exception);
    }
}
