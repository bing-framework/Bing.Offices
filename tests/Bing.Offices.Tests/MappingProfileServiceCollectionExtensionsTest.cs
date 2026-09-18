using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// Mapping Profile 服务集合扩展的参数校验、注册原子性和程序集扫描测试。
/// </summary>
public class MappingProfileServiceCollectionExtensionsTest
{
    /// <summary>
    /// 验证空值参数应被拒绝。
    /// </summary>
    [Fact]
    public void NullArguments_ShouldBeRejected()
    {
        Assert.Throws<ArgumentNullException>(() =>
            MappingProfileServiceCollectionExtensions.AddMappingProfile<ValidProfile>(null));
        Assert.Throws<ArgumentNullException>(() =>
            MappingProfileServiceCollectionExtensions.AddMappingProfile<ValidProfile>(null, "profile"));
        Assert.Throws<ArgumentNullException>(() =>
            MappingProfileServiceCollectionExtensions.AddMappingProfiles(null, typeof(ValidProfile).Assembly));
        Assert.Throws<ArgumentNullException>(() =>
            MappingProfileServiceCollectionExtensions.AddMappingProfiles(new ServiceCollection(), null));
    }

    /// <summary>
    /// 验证显式注册会回退空白名称并返回相同服务。
    /// </summary>
    [Fact]
    public void ExplicitRegistration_ShouldFallbackBlankNameAndReturnSameServices()
    {
        var services = new ServiceCollection();

        var defaultNameResult = MappingProfileServiceCollectionExtensions.AddMappingProfile<ValidProfile>(services);
        var blankNameResult = MappingProfileServiceCollectionExtensions.AddMappingProfile<NamedProfile>(services,
            " \t");

        Assert.Same(services, defaultNameResult);
        Assert.Same(services, blankNameResult);

        using var provider = services.BuildServiceProvider();
        var registry = provider.GetRequiredService<IMappingProfileRegistry>();
        Assert.True(registry.TryGetDescriptor(typeof(ValidProfile).FullName, MappingDirection.Import,
            typeof(ProfileModel), out _));
        Assert.True(registry.TryGetDescriptor(typeof(NamedProfile).FullName, MappingDirection.Import,
            typeof(ProfileModel), out _));
    }

    /// <summary>
    /// 验证无效抽象 Profile 或冲突 Profile 会在任何注册前失败。
    /// </summary>
    [Fact]
    public void InvalidAbstractOrConflictingProfile_ShouldFailBeforeAnyRegistration()
    {
        AssertNoPartialRegistration<NoContractProfile>(services =>
            MappingProfileServiceCollectionExtensions.AddMappingProfile<NoContractProfile>(services));
        AssertNoPartialRegistration<AbstractProfile>(services =>
            MappingProfileServiceCollectionExtensions.AddMappingProfile<AbstractProfile>(services));

        var services = new ServiceCollection();
        var count = services.Count;
        var exception = Assert.Throws<InvalidOperationException>(() =>
            MappingProfileServiceCollectionExtensions.AddMappingProfile<ConflictingProfile<ProfileModel>>(services));

        Assert.Contains("重复方向 descriptor", exception.Message);
        Assert.Equal(count, services.Count);
        Assert.DoesNotContain(services, descriptor =>
            descriptor.ServiceType == typeof(ConflictingProfile<ProfileModel>));
    }

    /// <summary>
    /// 验证重复注册应失败不追加服务。
    /// </summary>
    [Fact]
    public void DuplicateRegistration_ShouldFailWithoutAppendingServices()
    {
        var services = new ServiceCollection();
        MappingProfileServiceCollectionExtensions.AddMappingProfile<ValidProfile>(services);
        var count = services.Count;

        var exception = Assert.Throws<InvalidOperationException>(() =>
            MappingProfileServiceCollectionExtensions.AddMappingProfile<ValidProfile>(services));

        Assert.Contains("Profile 注册名称重复", exception.Message);
        Assert.Equal(count, services.Count);
        Assert.Equal(1, services.Count(descriptor => descriptor.ServiceType == typeof(ValidProfile)));
    }

    /// <summary>
    /// 验证程序集扫描会返回服务并注册具体的受支持 Profile。
    /// </summary>
    [Fact]
    public void AssemblyScan_ShouldReturnServicesAndRegisterConcreteSupportedProfiles()
    {
        var services = new ServiceCollection();
        var assembly = new ControlledAssembly(new[]
        {
            typeof(ScannedProfile),
            typeof(AbstractProfile),
            typeof(NoContractProfile),
            typeof(OpenGenericProfile<>)
        });

        var result = MappingProfileServiceCollectionExtensions.AddMappingProfiles(services, assembly);

        Assert.Same(services, result);
        using var provider = services.BuildServiceProvider();
        var registry = provider.GetRequiredService<IMappingProfileRegistry>();
        Assert.True(registry.TryGetDescriptor(typeof(ScannedProfile).FullName, MappingDirection.Import,
            typeof(ProfileModel), out _));
        Assert.False(registry.TryGetDescriptor(typeof(AbstractProfile).FullName, MappingDirection.Import,
            typeof(ProfileModel), out _));
        Assert.False(registry.TryGetDescriptor(typeof(NoContractProfile).FullName, MappingDirection.Import,
            typeof(ProfileModel), out _));
        Assert.False(registry.TryGetDescriptor(typeof(OpenGenericProfile<>).FullName, MappingDirection.Import,
            typeof(ProfileModel), out _));
    }

    /// <summary>
    /// 验证程序集扫描部分类型加载失败时仍保留可加载的 Profile。
    /// </summary>
    [Fact]
    public void AssemblyScan_WhenSomeTypesFailToLoad_ShouldKeepLoadableProfiles()
    {
        var services = new ServiceCollection();
        var assembly = new ControlledAssembly(new[] { typeof(ScannedProfile), null },
            new Exception[] { new TypeLoadException("缺少扫描依赖") });

        var result = MappingProfileServiceCollectionExtensions.AddMappingProfiles(services, assembly);

        Assert.Same(services, result);
        using var provider = services.BuildServiceProvider();
        var registry = provider.GetRequiredService<IMappingProfileRegistry>();
        Assert.True(registry.TryGetDescriptor(typeof(ScannedProfile).FullName, MappingDirection.Import,
            typeof(ProfileModel), out _));
    }

    /// <summary>
    /// 验证程序集扫描没有可加载类型时会在任何注册前失败。
    /// </summary>
    [Fact]
    public void AssemblyScan_WhenNoTypesCanLoad_ShouldFailBeforeAnyRegistration()
    {
        var services = new ServiceCollection();
        var count = services.Count;
        var assembly = new ControlledAssembly(new Type[] { null },
            new Exception[] { new TypeLoadException("所有类型均不可加载") });

        var exception = Assert.Throws<InvalidOperationException>(() =>
            MappingProfileServiceCollectionExtensions.AddMappingProfiles(services, assembly));

        Assert.Contains("所有类型均不可加载", exception.Message);
        Assert.IsType<ReflectionTypeLoadException>(exception.InnerException);
        Assert.Equal(count, services.Count);
        Assert.DoesNotContain(services, descriptor => descriptor.ServiceType == typeof(ScannedProfile));
    }

    /// <summary>
    /// 验证程序集扫描发现重复 Profile 时会失败且不追加服务。
    /// </summary>
    [Fact]
    public void AssemblyScan_DuplicateProfile_ShouldFailWithoutAppendingServices()
    {
        var services = new ServiceCollection();
        var assembly = new ControlledAssembly(new[] { typeof(ScannedProfile) });
        MappingProfileServiceCollectionExtensions.AddMappingProfiles(services, assembly);
        var count = services.Count;

        var exception = Assert.Throws<InvalidOperationException>(() =>
            MappingProfileServiceCollectionExtensions.AddMappingProfiles(services, assembly));

        Assert.Contains("Profile 注册名称重复", exception.Message);
        Assert.Equal(count, services.Count);
        Assert.Equal(1, services.Count(descriptor => descriptor.ServiceType == typeof(ScannedProfile)));
    }

    /// <summary>
    /// 断言无部分注册。
    /// </summary>
    /// <typeparam name="TProfile">泛型参数 TProfile 表示方法处理的数据类型。</typeparam>
    /// <param name="register">注册操作。</param>
    private static void AssertNoPartialRegistration<TProfile>(Action<IServiceCollection> register)
        where TProfile : class
    {
        var services = new ServiceCollection();
        var count = services.Count;

        Assert.Throws<ArgumentException>(() => register(services));

        Assert.Equal(count, services.Count);
        Assert.DoesNotContain(services, descriptor => descriptor.ServiceType == typeof(TProfile));
    }

    /// <summary>
    /// 表示 Profile 扫描测试使用的模型。
    /// </summary>
    public sealed class ProfileModel
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 提供测试场景使用的映射 Profile。
    /// </summary>
    public sealed class ValidProfile : IMappingProfile<ProfileModel, ProfileModel>
    {
        /// <inheritdoc />
        public void Configure(FluentSetting<ProfileModel, ProfileModel> setting)
        {
        }
    }

    /// <summary>
    /// 提供测试场景使用的映射 Profile。
    /// </summary>
    public sealed class NamedProfile : IMappingProfile<ProfileModel, ProfileModel>
    {
        /// <inheritdoc />
        public void Configure(FluentSetting<ProfileModel, ProfileModel> setting)
        {
        }
    }

    /// <summary>
    /// 提供测试场景使用的映射 Profile。
    /// </summary>
    public sealed class ScannedProfile : IMappingProfile<ProfileModel, ProfileModel>
    {
        /// <inheritdoc />
        public void Configure(FluentSetting<ProfileModel, ProfileModel> setting)
        {
        }
    }

    /// <summary>
    /// 提供测试场景使用的映射 Profile。
    /// </summary>
    public sealed class NoContractProfile
    {
    }

    /// <summary>
    /// 提供测试场景使用的映射 Profile。
    /// </summary>
    public abstract class AbstractProfile : IMappingProfile<ProfileModel, ProfileModel>
    {
        /// <inheritdoc />
        public abstract void Configure(FluentSetting<ProfileModel, ProfileModel> setting);
    }

    /// <summary>
    /// 提供测试场景使用的映射 Profile。
    /// </summary>
    /// <typeparam name="T">导入和导出共用的模型类型。</typeparam>
    public sealed class OpenGenericProfile<T> : IMappingProfile<T, T> where T : class, new()
    {
        /// <inheritdoc />
        public void Configure(FluentSetting<T, T> setting)
        {
        }
    }

    /// <summary>
    /// 提供测试场景使用的映射 Profile。
    /// </summary>
    /// <typeparam name="T">用于构造冲突映射契约的模型类型。</typeparam>
    public sealed class ConflictingProfile<T> : IImportMappingProfile<T>, IMappingProfile<T, T>
        where T : class, new()
    {
        /// <inheritdoc />
        public void Configure(ImportMappingBuilder<T> setting)
        {
        }

        /// <inheritdoc />
        public void Configure(FluentSetting<T, T> setting)
        {
        }
    }

    /// <summary>
    /// 提供可控类型和加载异常的程序集替身。
    /// </summary>
    private sealed class ControlledAssembly : Assembly
    {
        /// <summary>
        /// 模拟程序集返回的类型集合。
        /// </summary>
        private readonly Type[] _types;

        /// <summary>
        /// 模拟程序集加载失败时返回的异常集合。
        /// </summary>
        private readonly Exception[] _loaderExceptions;

        /// <summary>
        /// 初始化一个 <see cref="ControlledAssembly" /> 类型的实例。
        /// </summary>
        /// <param name="types">程序集公开的类型数组。</param>
        /// <param name="loaderExceptions">模拟的类型加载异常数组；null 表示正常返回类型。</param>
        public ControlledAssembly(Type[] types, Exception[] loaderExceptions = null)
        {
            _types = types;
            _loaderExceptions = loaderExceptions;
        }

        /// <inheritdoc />
        public override string FullName => "Bing.Offices.Tests.ControlledMappingProfileAssembly";

        /// <inheritdoc />
        public override Type[] GetTypes()
        {
            if (_loaderExceptions == null)
                return _types;
            throw new ReflectionTypeLoadException(_types, _loaderExceptions);
        }
    }
}
