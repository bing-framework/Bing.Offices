using System;
using System.Linq;
using System.Reflection;
using Bing.Extensions;
using Bing.Offices.Configurations;
using Bing.Offices.Internals;

namespace Bing.Offices
{
    /// <summary>
    /// 设置
    /// </summary>
    public static class FluentSettings
    {
        /// <summary>
        /// 映射配置文件的配置方法名
        /// </summary>
        private const string MappingProfileConfigureMethodName = "Configure";

        /// <summary>
        /// 映射配置文件的泛型类型定义
        /// </summary>
        private static readonly Type ProfileGenericTypeDefinition = typeof(IMappingProfile<>);

        /// <summary>
        /// 获取指定实体配置入口点
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        public static IExcelConfiguration<TEntity> For<TEntity>() => InternalHelper.GetExcelConfigurationMapping<TEntity>();

        /// <summary>
        /// 加载映射配置文件
        /// </summary>
        /// <param name="assemblies">程序集数组</param>
        public static void LoadMappingProfiles(params Assembly[] assemblies)
        {
            if (assemblies == null)
                throw new ArgumentNullException(nameof(assemblies));
            if (assemblies.Length == 0)
                assemblies = AppDomain.CurrentDomain.GetAssemblies();
            LoadMappingProfiles(assemblies.SelectMany(ass => ass.GetExportedTypes()).ToArray());
        }

        /// <summary>
        /// 加载映射配置文件
        /// </summary>
        /// <param name="types">映射配置文件类型数组</param>
        public static void LoadMappingProfiles(params Type[] types)
        {
            if (types == null)
                throw new ArgumentNullException(nameof(types));
            foreach (var type in types.Where(x=>x.IsBaseOn<IMappingProfile>()))
            {
                if (Activator.CreateInstance(type) is IMappingProfile profile)
                    LoadMappingProfile(profile);
            }
        }

        /// <summary>
        /// 加载映射配置文件
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        /// <typeparam name="TMappingProfile">映射配置文件类型</typeparam>
        public static void LoadMappingProfile<TEntity, TMappingProfile>()
            where TMappingProfile : IMappingProfile<TEntity>, new()
        {
            var profile = new TMappingProfile();
            profile.Configure(InternalHelper.GetExcelConfigurationMapping<TEntity>());
        }

        /// <summary>
        /// 加载映射配置文件
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        /// <param name="profile">映射配置文件</param>
        public static void LoadMappingProfile<TEntity>(IMappingProfile<TEntity> profile)
        {
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));
            profile.Configure(InternalHelper.GetExcelConfigurationMapping<TEntity>());
        }

        /// <summary>
        /// 加载映射配置文件
        /// </summary>
        /// <typeparam name="TMappingProfile">映射配置文件类型</typeparam>
        public static void LoadMappingProfile<TMappingProfile>() where TMappingProfile : IMappingProfile, new() =>
            LoadMappingProfile(new TMappingProfile());

        /// <summary>
        /// 加载映射配置文件
        /// </summary>
        /// <param name="profile">映射配置文件</param>
        private static void LoadMappingProfile(IMappingProfile profile)
        {
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));
            var profileInterfaceType = profile.GetType()
                .GetTypeInfo()
                .ImplementedInterfaces
                .FirstOrDefault(x => x.IsGenericType && x.GetGenericTypeDefinition() == ProfileGenericTypeDefinition);
            if (profileInterfaceType == null)
                return;
            var entityType = profileInterfaceType.GetGenericArguments()[0];
            var configuration = InternalHelper.GetExcelConfigurationMapping(entityType);
            var method = profileInterfaceType.GetMethod(MappingProfileConfigureMethodName,
                new[] {typeof(IExcelConfiguration<>).MakeGenericType(entityType)});
            method?.Invoke(profile, new object[] {configuration});
        }
    }
}
