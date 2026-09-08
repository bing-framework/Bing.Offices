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

    private static void AssertNoPartialRegistration<TProfile>(Action<IServiceCollection> register)
        where TProfile : class
    {
        var services = new ServiceCollection();
        var count = services.Count;

        Assert.Throws<ArgumentException>(() => register(services));

        Assert.Equal(count, services.Count);
        Assert.DoesNotContain(services, descriptor => descriptor.ServiceType == typeof(TProfile));
    }

    public sealed class ProfileModel
    {
        public string Name { get; set; }
    }

    public sealed class ValidProfile : IMappingProfile<ProfileModel, ProfileModel>
    {
        public void Configure(FluentSetting<ProfileModel, ProfileModel> setting)
        {
        }
    }

    public sealed class NamedProfile : IMappingProfile<ProfileModel, ProfileModel>
    {
        public void Configure(FluentSetting<ProfileModel, ProfileModel> setting)
        {
        }
    }

    public sealed class ScannedProfile : IMappingProfile<ProfileModel, ProfileModel>
    {
        public void Configure(FluentSetting<ProfileModel, ProfileModel> setting)
        {
        }
    }

    public sealed class NoContractProfile
    {
    }

    public abstract class AbstractProfile : IMappingProfile<ProfileModel, ProfileModel>
    {
        public abstract void Configure(FluentSetting<ProfileModel, ProfileModel> setting);
    }

    public sealed class OpenGenericProfile<T> : IMappingProfile<T, T> where T : class, new()
    {
        public void Configure(FluentSetting<T, T> setting)
        {
        }
    }

    public sealed class ConflictingProfile<T> : IImportMappingProfile<T>, IMappingProfile<T, T>
        where T : class, new()
    {
        public void Configure(ImportMappingBuilder<T> setting)
        {
        }

        public void Configure(FluentSetting<T, T> setting)
        {
        }
    }

    private sealed class ControlledAssembly : Assembly
    {
        private readonly Type[] _types;
        private readonly Exception[] _loaderExceptions;

        public ControlledAssembly(Type[] types, Exception[] loaderExceptions = null)
        {
            _types = types;
            _loaderExceptions = loaderExceptions;
        }

        public override string FullName => "Bing.Offices.Tests.ControlledMappingProfileAssembly";

        public override Type[] GetTypes()
        {
            if (_loaderExceptions == null)
                return _types;
            throw new ReflectionTypeLoadException(_types, _loaderExceptions);
        }
    }
}
