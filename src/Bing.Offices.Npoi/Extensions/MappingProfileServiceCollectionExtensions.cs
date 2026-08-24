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
    /// 显式注册一个双模型 Profile。
    /// </summary>
    public static IServiceCollection AddMappingProfile<TProfile, TImport, TExport>(
        this IServiceCollection services, string profileName = null)
        where TProfile : class, IMappingProfile<TImport, TExport>
        where TImport : class, new()
        where TExport : class, new()
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));
        var name = profileName ?? typeof(TProfile).FullName;
        ValidateUniqueRegistration(services, name, typeof(TImport), typeof(TExport));
        services.TryAddSingleton<TProfile>();
        services.AddSingleton(new MappingProfileRegistration(name, typeof(TImport), typeof(TExport), provider =>
            new ExcelMappingProfile<TImport, TExport>(provider.GetRequiredService<TProfile>())));
        AddRegistry(services);
        return services;
    }

    /// <summary>
    /// 扫描一个程序集并注册其中的非抽象 Profile 类型。
    /// </summary>
    public static IServiceCollection AddMappingProfilesFromAssembly(this IServiceCollection services,
        Assembly assembly)
    {
        if (assembly == null)
            throw new ArgumentNullException(nameof(assembly));
        foreach (var type in assembly.GetTypes()
                 .Where(type => !type.IsAbstract && !type.IsInterface && !type.ContainsGenericParameters)
                 .OrderBy(type => type.FullName, StringComparer.Ordinal))
        {
            var contract = type.GetInterfaces().FirstOrDefault(item =>
                item.IsGenericType && item.GetGenericTypeDefinition() == typeof(IMappingProfile<,>));
            if (contract == null)
                continue;
            var arguments = contract.GetGenericArguments();
            var method = typeof(MappingProfileServiceCollectionExtensions).GetMethod(
                nameof(AddMappingProfile), BindingFlags.Public | BindingFlags.Static);
            try
            {
                method.MakeGenericMethod(type, arguments[0], arguments[1]).Invoke(null,
                    new object[] { services, type.FullName });
            }
            catch (TargetInvocationException exception) when (exception.InnerException != null)
            {
                ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
                throw;
            }
        }
        return services;
    }

    /// <summary>
    /// 扫描多个程序集并注册其中的 Profile 类型。
    /// </summary>
    public static IServiceCollection AddMappingProfilesFromAssemblies(this IServiceCollection services,
        params Assembly[] assemblies)
    {
        if (assemblies == null)
            throw new ArgumentNullException(nameof(assemblies));
        foreach (var assembly in assemblies.Where(assembly => assembly != null)
                     .OrderBy(assembly => assembly.FullName, StringComparer.Ordinal))
            AddMappingProfilesFromAssembly(services, assembly);
        return services;
    }

    private static void AddRegistry(IServiceCollection services)
    {
        services.TryAddSingleton<MappingProfileRegistry>(provider =>
        {
            var registry = new MappingProfileRegistry();
            foreach (var registration in provider.GetServices<MappingProfileRegistration>())
                registration.Apply(registry, provider);
            return registry;
        });
        services.TryAddSingleton<IMappingProfileRegistry>(provider =>
            provider.GetRequiredService<MappingProfileRegistry>());
    }

    private static void ValidateUniqueRegistration(IServiceCollection services, string name,
        Type importType, Type exportType)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Profile 名称不能为空。", nameof(name));
        if (services.Any(descriptor => descriptor.ServiceType == typeof(MappingProfileRegistration)
            && descriptor.ImplementationInstance is MappingProfileRegistration registration
            && registration.Conflicts(name, importType, exportType)))
            throw new InvalidOperationException($"Profile 注册键重复: {name}");
    }

    private sealed class MappingProfileRegistration
    {
        private readonly string _name;
        private readonly Type _importType;
        private readonly Type _exportType;
        private readonly Func<IServiceProvider, IMappingProfileSnapshot> _factory;

        public MappingProfileRegistration(string name, Type importType, Type exportType,
            Func<IServiceProvider, IMappingProfileSnapshot> factory)
        {
            _name = name;
            _factory = factory;
            _importType = importType;
            _exportType = exportType;
        }

        public bool Conflicts(string name, Type importType, Type exportType) =>
            string.Equals(_name, name, StringComparison.OrdinalIgnoreCase)
            && (_importType == importType || _exportType == exportType);

        public void Apply(MappingProfileRegistry registry, IServiceProvider provider)
        {
            var snapshot = _factory(provider);
            registry.Register(_name, snapshot);
        }
    }
}
