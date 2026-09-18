using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Exceptions;
using Bing.Offices.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 映射配置加载器静态与 DI 边界合同测试。
/// </summary>
public sealed class ExcelMappingConfigurationLoaderTest
{
    /// <summary>
    /// 验证已删除的诊断 API 不再公开。
    /// </summary>
    [Fact]
    public void RemovedDiagnosticsSurface_ShouldNotBePublished()
    {
        var methods = typeof(ExcelMappingConfigurationLoader).GetMethods(
            BindingFlags.Public | BindingFlags.Static);

        Assert.DoesNotContain(methods, method => method.GetParameters().Any(parameter => parameter.IsOut));
        Assert.Null(typeof(ExcelMappingConfigurationLoader).Assembly.GetType(
            "Bing.Offices.Configurations.ExcelMappingDiagnostic"));
    }

    /// <summary>
    /// 验证静态配置解析失败时不通知异常观察器。
    /// </summary>
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

    /// <summary>
    /// 验证默认加载器仅通知一次最终抛出的异常实例。
    /// </summary>
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
