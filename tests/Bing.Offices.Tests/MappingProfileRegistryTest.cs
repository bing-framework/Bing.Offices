using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Bing.Offices.Configurations;
using Bing.Offices.Extensions;
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
        services.AddMappingProfile<ExplicitProfile>();

        // Act
        using var provider = services.BuildServiceProvider();
        var registry = provider.GetRequiredService<IMappingProfileRegistry>();
        Assert.True(registry.TryGetDescriptor(typeof(ExplicitProfile).FullName, MappingDirection.Import,
            typeof(ImportModel), out var import));
        Assert.True(registry.TryGetDescriptor(typeof(ExplicitProfile).FullName, MappingDirection.Export,
            typeof(ExportModel), out var export));

        // Assert
        Assert.Equal("导入", import.Configuration.Columns[0].Title);
        Assert.Equal("导出", export.Configuration.Columns[0].Title);
    }

    /// <summary>
    /// 测试 - 显式稳定 Profile 名称应替代 FullName 参与方向解析，并可通过只读 Resolver 消费。
    /// </summary>
    [Fact]
    public void ExplicitRegistration_WithStableName_ShouldResolveThroughReadOnlyResolver()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddMappingProfile<ExplicitProfile>("orders");

        // Act
        using var provider = services.BuildServiceProvider();
        var resolver = provider.GetRequiredService<IMappingProfileResolver>();

        // Assert
        Assert.True(resolver.TryGetDescriptor("orders", MappingDirection.Import,
            typeof(ImportModel), out var descriptor));
        Assert.Equal("orders", descriptor.Name);
        Assert.False(resolver.TryGetDescriptor(typeof(ExplicitProfile).FullName,
            MappingDirection.Import, typeof(ImportModel), out _));
    }

    /// <summary>
    /// 测试 - 四种 Profile 形状均应生成对应的方向 descriptor，且单向 Profile 不伪造另一方向。
    /// </summary>
    [Fact]
    public void ExplicitRegistration_AllProfileShapes_ShouldResolveExpectedDirections()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddMappingProfile<ImportOnlyProfile>();
        services.AddMappingProfile<ExportOnlyProfile>();
        services.AddMappingProfile<SameModelProfile>();
        services.AddMappingProfile<ExplicitProfile>();

        // Act
        using var provider = services.BuildServiceProvider();
        var registry = provider.GetRequiredService<IMappingProfileRegistry>();

        // Assert
        Assert.True(registry.TryGetDescriptor(typeof(ImportOnlyProfile).FullName, MappingDirection.Import,
            typeof(ImportModel), out _));
        Assert.False(registry.TryGetDescriptor(typeof(ImportOnlyProfile).FullName, MappingDirection.Export,
            typeof(ImportModel), out _));
        Assert.True(registry.TryGetDescriptor(typeof(ExportOnlyProfile).FullName, MappingDirection.Export,
            typeof(ExportModel), out _));
        Assert.False(registry.TryGetDescriptor(typeof(ExportOnlyProfile).FullName, MappingDirection.Import,
            typeof(ExportModel), out _));
        Assert.True(registry.TryGetDescriptor(typeof(SameModelProfile).FullName, MappingDirection.Import,
            typeof(ImportModel), out _));
        Assert.True(registry.TryGetDescriptor(typeof(SameModelProfile).FullName, MappingDirection.Export,
            typeof(ImportModel), out _));
        Assert.True(registry.TryGetDescriptor(typeof(ExplicitProfile).FullName, MappingDirection.Import,
            typeof(ImportModel), out _));
        Assert.True(registry.TryGetDescriptor(typeof(ExplicitProfile).FullName, MappingDirection.Export,
            typeof(ExportModel), out _));
    }

    /// <summary>
    /// 测试 - DI 注册的方向 Profile 应进入默认 Plan Factory 主链并生成对应计划。
    /// </summary>
    [Fact]
    public void DiProfileRegistry_DefaultPlanFactory_ShouldCreateDirectionalPlan()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddMappingProfile<ExplicitProfile>();
        services.AddNpoi();
        using var provider = services.BuildServiceProvider();
        var document = new ExcelMappingDocument
        {
            Import = new ExcelMappingConfiguration
            {
                Profile = typeof(ExplicitProfile).FullName
            }
        };

        // Act
        var factory = provider.GetRequiredService<Bing.Offices.Providers.IExcelMappingPlanFactory>();
        var plan = factory.Create<ImportModel>(document, MappingDirection.Import);

        // Assert
        Assert.Equal("导入", plan.Columns.Single(column => column.Name == nameof(ImportModel.Name)).Title);
    }

    /// <summary>
    /// 测试 - 一个 Profile 通过多个契约产生相同方向和模型时应报告确定性冲突。
    /// </summary>
    [Fact]
    public void MultipleContracts_SameDirectionAndModel_ShouldFailDeterministically()
    {
        // Arrange
        var services = new ServiceCollection();
        var count = services.Count;

        // Act
        var exception = Assert.Throws<InvalidOperationException>(() =>
            services.AddMappingProfile<ConflictingProfile<ImportModel>>());

        // Assert
        Assert.Contains("重复方向 descriptor", exception.Message);
        Assert.Equal(count, services.Count);
        Assert.DoesNotContain(services, descriptor => descriptor.ServiceType == typeof(ConflictingProfile<ImportModel>));
    }

    /// <summary>
    /// 测试 - 相同名称与模型方向的重复注册应在注册阶段失败。
    /// </summary>
    [Fact]
    public void DuplicateRegistration_ShouldFailBeforeProviderBuild()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddMappingProfile<ExplicitProfile>();

        // Act
        var action = () => services.AddMappingProfile<ExplicitProfile>();

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
        services.AddMappingProfiles(typeof(ScannedProfile).Assembly);

        // Act
        using var provider = services.BuildServiceProvider();
        var registry = provider.GetRequiredService<IMappingProfileRegistry>();

        // Assert
        Assert.True(registry.TryGetDescriptor(typeof(ScannedProfile).FullName, MappingDirection.Import,
            typeof(ImportModel), out _));
        Assert.False(registry.TryGetDescriptor(typeof(IgnoredAbstractProfile).FullName, MappingDirection.Import,
            typeof(ImportModel), out _));
        Assert.False(registry.TryGetDescriptor(typeof(IgnoredOpenGenericProfile<>).FullName, MappingDirection.Import,
            typeof(ImportModel), out _));
        Assert.False(registry.TryGetDescriptor("missing", MappingDirection.Import, typeof(ImportModel), out _));
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
        services.AddMappingProfiles(typeof(ScannedProfile).Assembly);
        var action = () => services.AddMappingProfiles(typeof(ScannedProfile).Assembly);

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
        registry.Register(new ProfileDescriptor("concurrent", MappingDirection.Import,
            typeof(ImportModel), new ExcelMappingConfiguration()));

        // Act
        var profiles = new ProfileDescriptor[64];
        Parallel.For(0, profiles.Length, index =>
            registry.TryGetDescriptor("concurrent", MappingDirection.Import, typeof(ImportModel), out profiles[index]));

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
        Assert.True(first.TryGetDescriptor(typeof(ScannedProfile).FullName, MappingDirection.Import,
            typeof(ImportModel), out var firstSnapshot));
        Assert.True(second.TryGetDescriptor(typeof(ScannedProfile).FullName, MappingDirection.Import,
            typeof(ImportModel), out var secondSnapshot));
        Assert.Equal(firstSnapshot.ModelType, secondSnapshot.ModelType);
        Assert.Equal(firstSnapshot.Direction, secondSnapshot.Direction);
        Assert.True(first.TryGetDescriptor(typeof(ExternalMappingProfile).FullName, MappingDirection.Import,
            typeof(ExternalImportModel), out var firstExternal));
        Assert.True(second.TryGetDescriptor(typeof(ExternalMappingProfile).FullName, MappingDirection.Import,
            typeof(ExternalImportModel), out var secondExternal));
        Assert.Equal(firstExternal.ModelType, secondExternal.ModelType);
        Assert.Equal(firstExternal.Direction, secondExternal.Direction);
    }

    /// <summary>
    /// 测试 - 扫描外部程序集后，测试程序集不得用相同 Profile key 覆盖已有注册。
    /// </summary>
    [Fact]
    public void AssemblyScan_CrossAssemblyDuplicateKey_ShouldFailBeforeProviderBuild()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddMappingProfile<ExternalMappingProfile>();

        // Act
        var action = () => services.AddMappingProfile<ExternalMappingProfile>();

        // Assert
        Assert.Throws<InvalidOperationException>(action);
    }

    /// <summary>
    /// 测试 - 程序集部分类型加载失败时，应继续注册可加载的 Profile 并保留诊断上下文。
    /// </summary>
    [Fact]
    public void AssemblyScan_WhenSomeTypesFailToLoad_ShouldKeepLoadableProfiles()
    {
        // Arrange
        var services = new ServiceCollection();
        var assembly = new ThrowingAssembly(new[] { typeof(ScannedProfile), null },
            new Exception[] { new TypeLoadException("missing dependency") });

        // Act
        services.AddMappingProfiles(assembly);
        using var provider = services.BuildServiceProvider();
        var registry = provider.GetRequiredService<IMappingProfileRegistry>();

        // Assert
        Assert.True(registry.TryGetDescriptor(typeof(ScannedProfile).FullName,
            MappingDirection.Import, typeof(ImportModel), out _));
    }

    /// <summary>
    /// 测试 - 程序集没有任何可加载类型时，应抛出包含加载诊断的异常。
    /// </summary>
    [Fact]
    public void AssemblyScan_WhenNoTypesCanLoad_ShouldThrowWithDiagnostics()
    {
        // Arrange
        var services = new ServiceCollection();
        var assembly = new ThrowingAssembly(new Type[] { null },
            new Exception[] { new TypeLoadException("all types unavailable") });

        // Act
        var exception = Assert.Throws<InvalidOperationException>(() =>
            services.AddMappingProfiles(assembly));

        // Assert
        Assert.Contains("all types unavailable", exception.Message);
        Assert.IsType<ReflectionTypeLoadException>(exception.InnerException);
    }

    private static ServiceProvider BuildProvider(params System.Reflection.Assembly[] assemblies)
    {
        var services = new ServiceCollection();
        foreach (var assembly in assemblies)
            services.AddMappingProfiles(assembly);
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

    public sealed class ConflictingProfile<T> : IImportMappingProfile<T>,
        IMappingProfile<T, ExportModel> where T : class, new()
    {
        public void Configure(ImportMappingBuilder<T> setting)
        {
        }

        public void Configure(FluentSetting<T, ExportModel> setting)
        {
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

    public sealed class ImportOnlyProfile : IImportMappingProfile<ImportModel>
    {
        public void Configure(ImportMappingBuilder<ImportModel> setting)
        {
            setting.Property(model => model.Name).HasHeader("仅导入");
        }
    }

    public sealed class ExportOnlyProfile : IExportMappingProfile<ExportModel>
    {
        public void Configure(ExportMappingBuilder<ExportModel> setting)
        {
            setting.Property(model => model.Label).HasHeader("仅导出");
        }
    }

    public sealed class SameModelProfile : IMappingProfile<ImportModel>
    {
        public void Configure(FluentSetting<ImportModel, ImportModel> setting)
        {
            setting.Import.Property(model => model.Name).HasHeader("同模型导入");
            setting.Export.Property(model => model.Name).HasHeader("同模型导出");
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

    private sealed class ThrowingAssembly : Assembly
    {
        private readonly Type[] _types;
        private readonly Exception[] _loaderExceptions;

        public ThrowingAssembly(Type[] types, Exception[] loaderExceptions)
        {
            _types = types;
            _loaderExceptions = loaderExceptions;
        }

        public override string FullName => "Bing.Offices.Tests.ThrowingAssembly";

        public override Type[] GetTypes() => throw new ReflectionTypeLoadException(_types, _loaderExceptions);
    }
}
