using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.Configurations;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Bing.Offices.Extensions;

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
    /// <returns>已添加 Profile 注册的服务集合。</returns>
    public static IServiceCollection AddMappingProfile<TProfile>(this IServiceCollection services)
        where TProfile : class
        => AddMappingProfile<TProfile>(services, null);

    /// <summary>
    /// 注册一个包含受支持 Profile 契约的具体类型，并使用调用方提供的稳定名称。
    /// </summary>
    /// <typeparam name="TProfile">Profile 实现类型。</typeparam>
    /// <param name="services">DI 服务集合。</param>
    /// <param name="profileName">稳定 Profile 名称；为空时兼容使用类型 FullName。</param>
    /// <returns>已添加 Profile 注册的服务集合。</returns>
    public static IServiceCollection AddMappingProfile<TProfile>(this IServiceCollection services,
        string profileName)
        where TProfile : class
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));
        ValidateProfileType(typeof(TProfile));
        if (!ProfileDescriptorFactory.HasSupportedContract(typeof(TProfile)))
            throw new ArgumentException($"Profile 类型未实现受支持契约: {typeof(TProfile).FullName}", nameof(TProfile));
        AddMappingProfilesCore(services, new[] { new ProfileRegistrationType(typeof(TProfile), profileName) });
        return services;
    }

    /// <summary>
    /// 扫描程序集并注册其中所有受支持的具体 Profile 类型。
    /// </summary>
    /// <param name="services">服务集合。</param>
    /// <param name="assembly">待扫描程序集。</param>
    /// <returns>已添加扫描结果注册的服务集合。</returns>
    public static IServiceCollection AddMappingProfiles(this IServiceCollection services, Assembly assembly)
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));
        if (assembly == null)
            throw new ArgumentNullException(nameof(assembly));
        var profileTypes = GetLoadableTypes(assembly)
                     .Where(type => !type.IsAbstract && !type.IsInterface && !type.ContainsGenericParameters)
                     .Where(type => ProfileDescriptorFactory.HasSupportedContract(type))
                     .OrderBy(type => type.FullName, StringComparer.Ordinal)
                     .Select(type => new ProfileRegistrationType(type, null))
                     .ToArray();
        AddMappingProfilesCore(services, profileTypes);
        return services;
    }

    /// <summary>
    /// 获取程序集可加载的类型；部分类型加载失败时保留其余可扫描类型并报告诊断信息。
    /// </summary>
    /// <param name="assembly">待扫描的程序集。</param>
    /// <returns>可成功加载的类型集合。</returns>
    private static IReadOnlyList<Type> GetLoadableTypes(Assembly assembly)
    {
        try
        {
            return assembly.GetTypes();
        }
        catch (ReflectionTypeLoadException exception)
        {
            var loadableTypes = exception.Types.Where(type => type != null).ToArray();
            var diagnostics = string.Join("; ", exception.LoaderExceptions
                .Where(error => error != null)
                .Select(error => error.Message));
            if (loadableTypes.Length == 0)
                throw new InvalidOperationException($"程序集没有可加载的 Profile 类型: {assembly.FullName}。{diagnostics}",
                    exception);
            if (!string.IsNullOrWhiteSpace(diagnostics))
                System.Diagnostics.Trace.TraceWarning(
                    $"程序集部分类型加载失败，已跳过不可加载类型: {assembly.FullName}。{diagnostics}");
            return loadableTypes;
        }
    }

    /// <summary>校验 Profile 注册名称和方向键后写入 DI 服务集合。</summary>
    /// <param name="services">DI 服务集合。</param>
    /// <param name="profileTypes">待注册的 Profile 类型及可选名称。</param>
    private static void AddMappingProfilesCore(IServiceCollection services,
        IReadOnlyList<ProfileRegistrationType> profileTypes)
    {
        var registrations = profileTypes.Select(item =>
            new MappingProfileRegistration(item.ProfileType, item.ProfileName)).ToArray();
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

    /// <summary>向 DI 容器注册 Profile 描述符仓储和解析器。</summary>
    /// <param name="services">DI 服务集合。</param>
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
        services.TryAddSingleton<IMappingProfileResolver>(provider =>
            provider.GetRequiredService<MappingProfileRegistry>());
    }

    /// <summary>验证 Profile 类型是可实例化且没有未绑定泛型参数的具体类型。</summary>
    /// <param name="profileType">待验证的 Profile 类型。</param>
    private static void ValidateProfileType(Type profileType)
    {
        if (profileType.IsAbstract || profileType.IsInterface || profileType.ContainsGenericParameters)
            throw new ArgumentException($"Profile 类型必须是封闭的具体类型: {profileType.FullName}", nameof(profileType));
    }

    /// <summary>保存一个 Profile 类型、稳定名称及其方向描述符键。</summary>
    private sealed class MappingProfileRegistration
    {
        /// <summary>已注册 Profile 的运行时类型。</summary>
        private readonly Type _profileType;

        /// <summary>创建包含方向描述符键的 Profile 注册记录。</summary>
        /// <param name="profileType">Profile 实现类型。</param>
        /// <param name="profileName">稳定 Profile 名称；为空时使用类型全名。</param>
        public MappingProfileRegistration(Type profileType, string profileName)
        {
            _profileType = profileType;
            ProfileName = string.IsNullOrWhiteSpace(profileName) ? profileType.FullName : profileName;
            Keys = ProfileDescriptorFactory.GetKeys(profileType, ProfileName);
        }

        /// <summary>获取已注册 Profile 的运行时类型。</summary>
        public Type ProfileType => _profileType;
        /// <summary>获取 Profile 的稳定注册名称。</summary>
        public string ProfileName { get; }
        /// <summary>获取 Profile 产生的方向和模型描述符键。</summary>
        public IReadOnlyList<string> Keys { get; }

        /// <summary>从 DI 容器解析 Profile 实例并创建方向描述符。</summary>
        /// <param name="provider">DI 服务提供程序。</param>
        /// <returns>该 Profile 产生的方向描述符集合。</returns>
        public IReadOnlyList<ProfileDescriptor> Create(IServiceProvider provider) =>
            ProfileDescriptorFactory.Create(provider.GetRequiredService(_profileType), _profileType,
                ProfileName);
    }

    /// <summary>保存程序集扫描阶段发现的 Profile 类型及可选名称。</summary>
    private sealed class ProfileRegistrationType
    {
        /// <summary>创建待注册的 Profile 类型记录。</summary>
        /// <param name="profileType">Profile 实现类型。</param>
        /// <param name="profileName">可选的稳定 Profile 名称。</param>
        public ProfileRegistrationType(Type profileType, string profileName)
        {
            ProfileType = profileType;
            ProfileName = profileName;
        }

        /// <summary>获取待注册的 Profile 类型。</summary>
        public Type ProfileType { get; }
        /// <summary>获取可选的稳定 Profile 名称。</summary>
        public string ProfileName { get; }
    }
}
