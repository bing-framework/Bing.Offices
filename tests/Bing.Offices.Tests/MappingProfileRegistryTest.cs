using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Configurations;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.ProfileFixtures;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// Mapping Profile Registry 测试。
/// </summary>
public class MappingProfileRegistryTest
{
    /// <summary>
    /// 测试 - 显式注册的 Profile 应按名称和方向解析，并保持导入导出配置隔离。
    /// </summary>
    [Fact]
    public void ExplicitRegistration_ShouldResolveDirectionalProfile()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddMappingProfile<ExplicitProfile, ImportModel, ExportModel>("orders");

        // Act
        using var provider = services.BuildServiceProvider();
        var registry = provider.GetRequiredService<IMappingProfileRegistry>();
        var profile = registry.Get<ImportModel, ExportModel>("orders");

        // Assert
        Assert.Equal("导入", profile.ImportConfiguration.Columns[0].Title);
        Assert.Equal("导出", profile.ExportConfiguration.Columns[0].Title);
    }

    /// <summary>
    /// 测试 - 相同名称与模型方向的重复注册应在注册阶段失败。
    /// </summary>
    [Fact]
    public void DuplicateRegistration_ShouldFailBeforeProviderBuild()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddMappingProfile<ExplicitProfile, ImportModel, ExportModel>("orders");

        // Act
        var action = () => services.AddMappingProfile<SecondExplicitProfile, ImportModel, ExportModel>("orders");

        // Assert
        Assert.Throws<InvalidOperationException>(action);
    }

    /// <summary>
    /// 测试 - 程序集扫描应注册实现接口的具体 Profile，且开放泛型与抽象类型不参与扫描。
    /// </summary>
    [Fact]
    public void AssemblyScan_ShouldRegisterConcreteProfiles()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddMappingProfilesFromAssembly(typeof(ScannedProfile).Assembly);

        // Act
        using var provider = services.BuildServiceProvider();
        var registry = provider.GetRequiredService<IMappingProfileRegistry>();

        // Assert
        Assert.True(registry.TryGet(typeof(ScannedProfile).FullName, MappingDirection.Import,
            typeof(ImportModel), out _));
        Assert.False(registry.TryGet(typeof(IgnoredAbstractProfile).FullName, MappingDirection.Import,
            typeof(ImportModel), out _));
        Assert.False(registry.TryGet(typeof(IgnoredOpenGenericProfile<>).FullName, MappingDirection.Import,
            typeof(ImportModel), out _));
        Assert.Throws<KeyNotFoundException>(() => registry.Get<ImportModel, ExportModel>("missing"));
    }

    /// <summary>
    /// 测试 - 重复扫描同一程序集不能依赖最后注册项覆盖已有 Profile。
    /// </summary>
    [Fact]
    public void AssemblyScan_DuplicateAssembly_ShouldFailDeterministically()
    {
        // Arrange
        var services = new ServiceCollection();

        // Act
        var action = () => services.AddMappingProfilesFromAssemblies(
            typeof(ScannedProfile).Assembly, typeof(ScannedProfile).Assembly);

        // Assert
        Assert.Throws<InvalidOperationException>(action);
    }

    /// <summary>
    /// 测试 - Registry 并发解析应返回同一不可变 Profile 快照而不抛异常。
    /// </summary>
    [Fact]
    public void Registry_ConcurrentReads_ShouldBeStable()
    {
        // Arrange
        var registry = new MappingProfileRegistry();
        var snapshot = new ExcelMappingProfile<ImportModel, ExportModel>(_ => { });
        registry.Register("concurrent", snapshot);

        // Act
        var profiles = new IMappingProfileSnapshot[64];
        Parallel.For(0, profiles.Length, index =>
            registry.TryGet("concurrent", MappingDirection.Import, typeof(ImportModel), out profiles[index]));

        // Assert
        Assert.All(profiles, profile => Assert.NotNull(profile));
        Assert.All(profiles.Skip(1), profile => Assert.Same(profiles[0], profile));
    }

    /// <summary>
    /// 测试 - 扫描两个真实程序集时，输入顺序不应改变可解析 Profile。
    /// </summary>
    [Fact]
    public void AssemblyScan_TwoAssemblies_ShouldBeOrderIndependent()
    {
        // Arrange
        using var firstProvider = BuildProvider(typeof(ScannedProfile).Assembly,
            typeof(ExternalMappingProfile).Assembly);
        using var secondProvider = BuildProvider(typeof(ExternalMappingProfile).Assembly,
            typeof(ScannedProfile).Assembly);

        // Act
        var first = firstProvider.GetRequiredService<IMappingProfileRegistry>();
        var second = secondProvider.GetRequiredService<IMappingProfileRegistry>();

        // Assert
        Assert.True(first.TryGet(typeof(ScannedProfile).FullName, MappingDirection.Import,
            typeof(ImportModel), out var firstSnapshot));
        Assert.True(second.TryGet(typeof(ScannedProfile).FullName, MappingDirection.Import,
            typeof(ImportModel), out var secondSnapshot));
        Assert.Equal(firstSnapshot.ImportType, secondSnapshot.ImportType);
        Assert.Equal(firstSnapshot.ExportType, secondSnapshot.ExportType);
        Assert.True(first.TryGet(typeof(ExternalMappingProfile).FullName, MappingDirection.Import,
            typeof(ExternalImportModel), out var firstExternal));
        Assert.True(second.TryGet(typeof(ExternalMappingProfile).FullName, MappingDirection.Import,
            typeof(ExternalImportModel), out var secondExternal));
        Assert.Equal(firstExternal.ImportType, secondExternal.ImportType);
        Assert.Equal(firstExternal.ExportType, secondExternal.ExportType);
    }

    /// <summary>
    /// 测试 - 扫描外部程序集后，测试程序集不得用相同 Profile key 覆盖已有注册。
    /// </summary>
    [Fact]
    public void AssemblyScan_CrossAssemblyDuplicateKey_ShouldFailBeforeProviderBuild()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddMappingProfilesFromAssembly(typeof(ExternalMappingProfile).Assembly);

        // Act
        var action = () => services.AddMappingProfile<ExternalDuplicateProfile, ExternalImportModel,
            ExternalExportModel>(typeof(ExternalMappingProfile).FullName);

        // Assert
        Assert.Throws<InvalidOperationException>(action);
    }

    private static ServiceProvider BuildProvider(params System.Reflection.Assembly[] assemblies)
    {
        var services = new ServiceCollection();
        services.AddMappingProfilesFromAssemblies(assemblies);
        return services.BuildServiceProvider();
    }

    public sealed class ExplicitProfile : IMappingProfile<ImportModel, ExportModel>
    {
        public void Configure(FluentSetting<ImportModel, ExportModel> setting)
        {
            setting.Import.Property(model => model.Name).HasHeader("导入");
            setting.Export.Property(model => model.Label).HasHeader("导出");
        }
    }

    public sealed class SecondExplicitProfile : IMappingProfile<ImportModel, ExportModel>
    {
        public void Configure(FluentSetting<ImportModel, ExportModel> setting)
        {
        }
    }

    public sealed class ScannedProfile : IMappingProfile<ImportModel, ExportModel>
    {
        public void Configure(FluentSetting<ImportModel, ExportModel> setting)
        {
        }
    }

    public sealed class ExternalDuplicateProfile : IMappingProfile<ExternalImportModel, ExternalExportModel>
    {
        public void Configure(FluentSetting<ExternalImportModel, ExternalExportModel> setting)
        {
        }
    }

    public abstract class IgnoredAbstractProfile : IMappingProfile<ImportModel, ExportModel>
    {
        public abstract void Configure(FluentSetting<ImportModel, ExportModel> setting);
    }

    public sealed class IgnoredOpenGenericProfile<T> : IMappingProfile<ImportModel, ExportModel>
    {
        public void Configure(FluentSetting<ImportModel, ExportModel> setting)
        {
        }
    }

    public sealed class ImportModel
    {
        public string Name { get; set; }
    }

    public sealed class ExportModel
    {
        public string Label { get; set; }
    }
}
