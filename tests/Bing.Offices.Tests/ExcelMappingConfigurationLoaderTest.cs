using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Exceptions;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>映射配置加载器静态与 DI 边界合同测试。</summary>
public sealed class ExcelMappingConfigurationLoaderTest
{
    /// <summary>空诊断 API 已删除，Loader 仅保留 v2 文档入口。</summary>
    [Fact]
    public void RemovedDiagnosticsSurface_ShouldNotBePublished()
    {
        var methods = typeof(ExcelMappingConfigurationLoader).GetMethods(
            BindingFlags.Public | BindingFlags.Static);

        Assert.DoesNotContain(methods, method => method.GetParameters().Any(parameter => parameter.IsOut));
        Assert.Null(typeof(ExcelMappingConfigurationLoader).Assembly.GetType(
            "Bing.Offices.Configurations.ExcelMappingDiagnostic"));
    }

    /// <summary>静态 Loader 是纯解析入口，不参与 Observer。</summary>
    [Fact]
    public void StaticLoader_InvalidDocument_ShouldThrowWithoutObservation()
    {
        // Arrange
        var observer = new RecordingObserver();

        // Act
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            ExcelMappingConfigurationLoader.FromJsonDocument("{"));

        // Assert
        Assert.Empty(observer.Exceptions);
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(BingOfficesOperation.Configuration, exception.Operation);
    }

    /// <summary>默认实现必须向 Observer 发送最终抛出的同一异常实例，重复经过 dispatcher 也只通知一次。</summary>
    [Fact]
    public void DefaultLoader_InvalidDocument_ShouldObserveSameInstanceOnce()
    {
        // Arrange
        var observer = new RecordingObserver();
        var loader = new DefaultExcelMappingConfigurationLoader(new[] { observer });

        // Act
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            loader.FromJsonDocument("{"));
        new BingOfficesExceptionDispatcher(new[] { observer }).Observe(exception);

        // Assert
        Assert.Same(exception, Assert.Single(observer.Exceptions));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
    }

    /// <summary>正式 DI 注册应把用户 Observer 注入默认 Loader。</summary>
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

    private sealed class RecordingObserver : IBingOfficesExceptionObserver
    {
        public List<BingOfficesException> Exceptions { get; } = new();

        public void Observe(BingOfficesException exception) => Exceptions.Add(exception);
    }
}
