using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using Bing.Offices.Configurations;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// Mapping Profile 的 DI 注册扩展。
/// </summary>
public static class MappingProfileServiceCollectionExtensions
{
    /// <summary>
    /// 注册一个包含受支持 Profile 契约的具体类型。
    /// </summary>
    /// <typeparam name="TProfile">Profile 实现类型。</typeparam>
    /// <param name="services">DI 服务集合。</param>
    public static IServiceCollection AddMappingProfile<TProfile>(this IServiceCollection services)
        where TProfile : class
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));
        ValidateProfileType(typeof(TProfile));
        if (!ProfileDescriptorFactory.HasSupportedContract(typeof(TProfile)))
            throw new ArgumentException($"Profile 类型未实现受支持契约: {typeof(TProfile).FullName}", nameof(TProfile));
        AddMappingProfilesCore(services, new[] { typeof(TProfile) });
        return services;
    }

    /// <summary>
    /// 扫描程序集并注册其中所有受支持的具体 Profile 类型。
    /// </summary>
    /// <param name="services">DI 服务集合。</param>
    /// <param name="assembly">待扫描程序集。</param>
    public static IServiceCollection AddMappingProfiles(this IServiceCollection services, Assembly assembly)
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));
        if (assembly == null)
            throw new ArgumentNullException(nameof(assembly));
        var profileTypes = assembly.GetTypes()
                     .Where(type => !type.IsAbstract && !type.IsInterface && !type.ContainsGenericParameters)
                     .Where(type => ProfileDescriptorFactory.HasSupportedContract(type))
                     .OrderBy(type => type.FullName, StringComparer.Ordinal)
                     .ToArray();
        AddMappingProfilesCore(services, profileTypes);
        return services;
    }

    private static void AddMappingProfilesCore(IServiceCollection services, IReadOnlyList<Type> profileTypes)
    {
        var registrations = profileTypes.Select(type => new MappingProfileRegistration(type)).ToArray();
        var existingRegistrations = services
            .Where(descriptor => descriptor.ServiceType == typeof(MappingProfileRegistration)
                && descriptor.ImplementationInstance is MappingProfileRegistration)
            .Select(descriptor => (MappingProfileRegistration)descriptor.ImplementationInstance)
            .ToArray();
        var names = new HashSet<string>(existingRegistrations.Select(item => item.ProfileName),
            StringComparer.OrdinalIgnoreCase);
        var keys = new HashSet<string>(existingRegistrations.SelectMany(item => item.Keys),
            StringComparer.OrdinalIgnoreCase);
        foreach (var registration in registrations)
        {
            if (!names.Add(registration.ProfileName))
                throw new InvalidOperationException($"Profile 注册名称重复: {registration.ProfileName}");
            foreach (var key in registration.Keys)
            {
                if (!keys.Add(key))
                    throw new InvalidOperationException($"Profile 注册键重复: {key}");
            }
        }
        foreach (var registration in registrations)
        {
            services.TryAdd(ServiceDescriptor.Singleton(registration.ProfileType, registration.ProfileType));
            services.AddSingleton(registration);
        }
        AddRegistry(services);
    }

    private static void AddRegistry(IServiceCollection services)
    {
        services.TryAddSingleton<MappingProfileRegistry>(provider =>
        {
            var registry = new MappingProfileRegistry();
            foreach (var registration in provider.GetServices<MappingProfileRegistration>())
            {
                foreach (var descriptor in registration.Create(provider))
                    registry.Register(descriptor);
            }
            return registry;
        });
        services.TryAddSingleton<IMappingProfileRegistry>(provider =>
            provider.GetRequiredService<MappingProfileRegistry>());
    }

    private static void ValidateProfileType(Type profileType)
    {
        if (profileType.IsAbstract || profileType.IsInterface || profileType.ContainsGenericParameters)
            throw new ArgumentException($"Profile 类型必须是封闭的具体类型: {profileType.FullName}", nameof(profileType));
    }

    private sealed class MappingProfileRegistration
    {
        private readonly Type _profileType;

        public MappingProfileRegistration(Type profileType)
        {
            _profileType = profileType;
            Keys = ProfileDescriptorFactory.GetKeys(profileType, ProfileName);
        }

        public Type ProfileType => _profileType;
        public string ProfileName => _profileType.FullName;
        public IReadOnlyList<string> Keys { get; }

        public IReadOnlyList<ProfileDescriptor> Create(IServiceProvider provider) =>
            ProfileDescriptorFactory.Create(provider.GetRequiredService(_profileType), _profileType,
                ProfileName);
    }
}
